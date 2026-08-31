#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("build", "validate", "ppam")]
    [string]$Action,

    [string]$InputPath,

    [string]$OutputPath,

    [string]$PpamOutputPath,

    [switch]$Force,

    [switch]$Visible,

    [switch]$NoValidation,

    [ValidateRange(10, 600)]
    [int]$CompileTimeoutSeconds = 60,

    [ValidateRange(10, 600)]
    [int]$OfficeTimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "Office artifact automation requires Windows PowerPoint."
}

if (
    [Threading.Thread]::CurrentThread.ApartmentState -ne
        [Threading.ApartmentState]::STA
) {
    throw "Run office-build.ps1 in an STA PowerShell process."
}

$MSO_TRUE = -1
$MSO_FALSE = 0
$MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
$PP_ALERTS_NONE = 1
$PP_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED = 25
$PP_SAVE_AS_OPEN_XML_ADDIN = 30
$VBEXT_CT_STDMODULE = 1
$VBEXT_CT_CLASSMODULE = 2
$VBEXT_CT_MSFORM = 3
$VBEXT_CT_DOCUMENT = 100
$SCRIPTING_RUNTIME_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$officeModulePath = Join-Path $PSScriptRoot "lib\IguanaTex.Office.psm1"
$packageModulePath = Join-Path $PSScriptRoot "lib\IguanaTex.Package.psm1"
$compileModulePath = Join-Path $PSScriptRoot "lib\IguanaTex.Compile.psm1"

foreach ($modulePath in @(
    $officeModulePath,
    $packageModulePath,
    $compileModulePath
)) {
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        throw "Required build module not found: $modulePath"
    }

    Import-Module -Name $modulePath -Force -DisableNameChecking
}

if (-not ("IguanaTexBuildNative" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class IguanaTexBuildNative
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId
    );
}

public sealed class IguanaTexPowerPointWatchdog : IDisposable
{
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_COMMAND = 0x0111;
    private const uint IDOK = 1;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    private readonly int processId;
    private readonly long startTimeUtcTicks;
    private readonly int timeoutMilliseconds;
    private readonly ManualResetEvent stopEvent = new ManualResetEvent(false);
    private readonly HashSet<long> originalDialogs = new HashSet<long>();
    private Thread thread;

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId
    );

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(
        IntPtr hWnd,
        StringBuilder text,
        int maximumCount
    );

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(
        IntPtr hWnd,
        StringBuilder className,
        int maximumCount
    );

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        uint message,
        UIntPtr wParam,
        IntPtr lParam,
        uint flags,
        uint timeout,
        out UIntPtr result
    );

    public bool DialogDetected { get; private set; }
    public bool DialogDismissed { get; private set; }
    public bool TimedOut { get; private set; }
    public bool ProcessKilled { get; private set; }
    public bool ShutdownFailed { get; private set; }
    public string DialogDescription { get; private set; }

    public IguanaTexPowerPointWatchdog(
        int processId,
        long startTimeUtcTicks,
        int timeoutSeconds
    )
    {
        this.processId = processId;
        this.startTimeUtcTicks = startTimeUtcTicks;
        this.timeoutMilliseconds = checked(timeoutSeconds * 1000);
        DialogDescription = String.Empty;
    }

    private bool IsExactProcessRunning()
    {
        try
        {
            using (Process process = Process.GetProcessById(processId))
            {
                return String.Equals(
                        process.ProcessName,
                        "POWERPNT",
                        StringComparison.OrdinalIgnoreCase) &&
                    process.StartTime.ToUniversalTime().Ticks == startTimeUtcTicks &&
                    !process.HasExited;
            }
        }
        catch
        {
            return false;
        }
    }

    private static string ReadWindowText(IntPtr hWnd)
    {
        StringBuilder value = new StringBuilder(1024);
        GetWindowText(hWnd, value, value.Capacity);
        return value.ToString();
    }

    private static string ReadClassName(IntPtr hWnd)
    {
        StringBuilder value = new StringBuilder(256);
        GetClassName(hWnd, value, value.Capacity);
        return value.ToString();
    }

    private bool IsCandidateDialog(IntPtr hWnd)
    {
        uint ownerProcessId;
        GetWindowThreadProcessId(hWnd, out ownerProcessId);

        if (ownerProcessId != (uint)processId || !IsWindowVisible(hWnd))
        {
            return false;
        }

        string className = ReadClassName(hWnd);
        return String.Equals(className, "#32770", StringComparison.Ordinal) ||
            String.Equals(className, "NUIDialog", StringComparison.Ordinal) ||
            className.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase);
    }

    private List<IntPtr> FindDialogs()
    {
        List<IntPtr> result = new List<IntPtr>();

        EnumWindows(delegate(IntPtr hWnd, IntPtr ignored)
        {
            if (IsCandidateDialog(hWnd))
            {
                result.Add(hWnd);
            }

            return true;
        }, IntPtr.Zero);

        return result;
    }

    private static bool DismissDialog(IntPtr hWnd)
    {
        UIntPtr ignored;
        SendMessageTimeout(
            hWnd,
            WM_COMMAND,
            new UIntPtr(IDOK),
            IntPtr.Zero,
            SMTO_ABORTIFHUNG,
            750,
            out ignored
        );

        for (int attempt = 0; attempt < 10 && IsWindow(hWnd); attempt++)
        {
            Thread.Sleep(50);
        }

        if (IsWindow(hWnd))
        {
            SendMessageTimeout(
                hWnd,
                WM_CLOSE,
                UIntPtr.Zero,
                IntPtr.Zero,
                SMTO_ABORTIFHUNG,
                750,
                out ignored
            );
        }

        for (int attempt = 0; attempt < 20 && IsWindow(hWnd); attempt++)
        {
            Thread.Sleep(50);
        }

        return !IsWindow(hWnd);
    }

    private bool KillExactProcess()
    {
        try
        {
            using (Process process = Process.GetProcessById(processId))
            {
                if (!String.Equals(
                        process.ProcessName,
                        "POWERPNT",
                        StringComparison.OrdinalIgnoreCase) ||
                    process.StartTime.ToUniversalTime().Ticks != startTimeUtcTicks)
                {
                    return false;
                }

                process.Kill();
                return process.WaitForExit(5000);
            }
        }
        catch
        {
            return false;
        }
    }

    private void Run()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();

        while (!stopEvent.WaitOne(100))
        {
            if (!IsExactProcessRunning())
            {
                return;
            }

            foreach (IntPtr dialog in FindDialogs())
            {
                if (originalDialogs.Contains(dialog.ToInt64()))
                {
                    continue;
                }

                DialogDetected = true;
                DialogDescription = String.Format(
                    "title='{0}', class='{1}'",
                    ReadWindowText(dialog),
                    ReadClassName(dialog)
                );
                DialogDismissed = DismissDialog(dialog);
            }

            if (stopwatch.ElapsedMilliseconds >= timeoutMilliseconds)
            {
                TimedOut = true;
                ProcessKilled = KillExactProcess();
                return;
            }
        }
    }

    public void Start()
    {
        foreach (IntPtr dialog in FindDialogs())
        {
            originalDialogs.Add(dialog.ToInt64());
        }

        thread = new Thread(Run);
        thread.IsBackground = true;
        thread.Name = "IguanaTex PowerPoint COM watchdog";
        thread.Start();
    }

    public bool Stop()
    {
        Thread activeThread = thread;

        if (activeThread == null)
        {
            return true;
        }

        Thread.Sleep(250);
        stopEvent.Set();

        if (!activeThread.Join(3000))
        {
            ShutdownFailed = true;
            ProcessKilled = KillExactProcess();

            if (!activeThread.Join(3000))
            {
                return false;
            }
        }

        thread = null;
        return true;
    }

    public void Dispose()
    {
        if (Stop())
        {
            stopEvent.Dispose();
        }
    }
}
"@
}

function ConvertTo-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [switch]$MustExist
    )

    $absolute = [System.IO.Path]::GetFullPath(
        $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    )

    if ($MustExist -and -not (Test-Path -LiteralPath $absolute -PathType Leaf)) {
        throw "File not found: $absolute"
    }

    return $absolute
}

function Get-CanonicalLayout {
    $sourceDirectory = Join-Path $projectRoot "src"
    $projectMetadataPath = Join-Path $projectRoot "office\project\project.json"
    $referencesPath = Join-Path $projectRoot "office\project\references.json"
    $ribbonDirectory = Join-Path $projectRoot "office\ribbon"

    foreach ($directory in @($sourceDirectory, $ribbonDirectory)) {
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
            throw "Canonical input directory not found: $directory"
        }
    }

    foreach ($file in @(
        $projectMetadataPath,
        $referencesPath,
        (Join-Path $ribbonDirectory "customUI.xml"),
        (Join-Path $ribbonDirectory "customUI14.xml")
    )) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
            throw "Canonical input file not found: $file"
        }
    }

    return [PSCustomObject]@{
        ProjectRoot = $projectRoot
        SourceDirectory = $sourceDirectory
        ProjectMetadataPath = $projectMetadataPath
        ReferencesPath = $referencesPath
        RibbonDirectory = $ribbonDirectory
    }
}

function Read-CanonicalConfiguration {
    param([object]$Layout)

    $metadata = Get-Content -LiteralPath $Layout.ProjectMetadataPath -Raw |
        ConvertFrom-Json
    $referenceConfig = Get-Content -LiteralPath $Layout.ReferencesPath -Raw |
        ConvertFrom-Json

    foreach ($propertyName in @(
        "name",
        "description",
        "protection",
        "conditionalCompilationArguments"
    )) {
        if ($metadata.PSObject.Properties.Name -notcontains $propertyName) {
            throw "Missing project metadata property: $propertyName"
        }
    }

    if ([string]::IsNullOrWhiteSpace([string]$metadata.name)) {
        throw "Project metadata name must not be empty."
    }

    if ([int]$metadata.protection -ne 0) {
        throw "Only an unprotected fresh VBProject is supported (protection = 0)."
    }

    if ([string]$metadata.conditionalCompilationArguments -ne "") {
        throw (
            "PowerPoint VBIDE does not expose project conditional compilation " +
            "arguments. This fresh build supports only the canonical empty value."
        )
    }

    if ($referenceConfig.PSObject.Properties.Name -notcontains "explicit") {
        throw "references.json must contain an explicit array."
    }

    $explicitReferences = @($referenceConfig.explicit)

    foreach ($reference in $explicitReferences) {
        foreach ($propertyName in @("name", "guid", "major", "minor")) {
            if ($reference.PSObject.Properties.Name -notcontains $propertyName) {
                throw "Explicit reference is missing property: $propertyName"
            }
        }

        $parsedGuid = [Guid]::Empty

        if (-not [Guid]::TryParse([string]$reference.guid, [ref]$parsedGuid)) {
            throw "Invalid explicit reference GUID: $($reference.guid)"
        }

        if (
            $parsedGuid -eq [Guid]$SCRIPTING_RUNTIME_GUID -or
            [string]$reference.name -eq "Scripting"
        ) {
            throw (
                "Microsoft Scripting Runtime must not be declared as an " +
                "explicit VBA project reference."
            )
        }
    }

    return [PSCustomObject]@{
        Metadata = $metadata
        ExplicitReferences = $explicitReferences
    }
}

function Get-PowerPointExecutablePath {
    $progidKey = Get-Item -LiteralPath (
        "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID"
    ) -ErrorAction Stop
    $classId = [string]$progidKey.GetValue("")

    if ([string]::IsNullOrWhiteSpace($classId)) {
        throw "PowerPoint.Application has no registered CLSID."
    }

    $serverKey = Get-Item -LiteralPath (
        "Registry::HKEY_CLASSES_ROOT\CLSID\$classId\LocalServer32"
    ) -ErrorAction Stop
    $serverCommand = [string]$serverKey.GetValue("")
    $match = [regex]::Match(
        $serverCommand,
        '(?i)^\s*"?(.*?\.exe)"?(?:\s|$)'
    )

    if (-not $match.Success) {
        throw "Could not parse the PowerPoint LocalServer32 command."
    }

    $executablePath = $match.Groups[1].Value

    if (-not (Test-Path -LiteralPath $executablePath -PathType Leaf)) {
        throw "Registered PowerPoint executable not found: $executablePath"
    }

    return $executablePath
}

function Get-ExactOwnedPowerPointProcess {
    param([object]$ApplicationInfo)

    if ($null -eq $ApplicationInfo -or -not $ApplicationInfo.Owned) {
        return $null
    }

    $process = Get-Process -Id ([int]$ApplicationInfo.ProcessId) `
        -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return $null
    }

    try {
        if (
            $process.ProcessName -ne "POWERPNT" -or
            [long]$process.StartTime.ToUniversalTime().Ticks -ne
                [long]$ApplicationInfo.StartTimeUtcTicks
        ) {
            $process.Dispose()
            return $null
        }

        [void]$process.Handle
        return $process
    }
    catch {
        $process.Dispose()
        return $null
    }
}

function Stop-ExactOwnedPowerPointProcess {
    param([object]$ApplicationInfo)

    $process = Get-ExactOwnedPowerPointProcess -ApplicationInfo $ApplicationInfo

    if ($null -eq $process) {
        return $false
    }

    try {
        $process.Kill()
        [void]$process.WaitForExit(5000)
        return $process.HasExited
    }
    catch {
        return $false
    }
    finally {
        $process.Dispose()
    }
}

function Invoke-OwnedPowerPointOperation {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ApplicationInfo,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $watchdog = New-Object `
        -TypeName IguanaTexPowerPointWatchdog `
        -ArgumentList @(
            [int]$ApplicationInfo.ProcessId,
            [long]$ApplicationInfo.StartTimeUtcTicks,
            [int]$OfficeTimeoutSeconds
        )
    $operationError = $null
    $value = $null
    $watchdogStopped = $false

    try {
        $watchdog.Start()

        try {
            $value = & $Operation
        }
        catch {
            $operationError = $_
        }
        finally {
            $watchdogStopped = $watchdog.Stop()
        }

        if ($watchdog.TimedOut) {
            throw (
                "$Description timed out after $OfficeTimeoutSeconds seconds. " +
                "Exact owned process termination succeeded: " +
                ([bool]$watchdog.ProcessKilled)
            )
        }

        if (-not $watchdogStopped -or $watchdog.ShutdownFailed) {
            throw (
                "$Description could not stop its exact-PID watchdog cleanly. " +
                "Exact owned process termination succeeded: " +
                ([bool]$watchdog.ProcessKilled)
            )
        }

        if ($watchdog.DialogDetected) {
            throw (
                "$Description produced a PowerPoint modal dialog (" +
                $watchdog.DialogDescription + "). Dismissed: " +
                ([bool]$watchdog.DialogDismissed)
            )
        }

        if ($null -ne $operationError) {
            throw $operationError
        }

        Write-Output -NoEnumerate $value
    }
    finally {
        $watchdog.Dispose()
    }
}

function New-OwnedPowerPointApplication {
    param([switch]$ShowWindow)

    $launcher = $null
    $application = $null
    $applicationInfo = $null
    $launchRecord = $null

    try {
        $powerPointPath = Get-PowerPointExecutablePath
        $launcher = Start-Process `
            -FilePath $powerPointPath `
            -ArgumentList "/AUTOMATION" `
            -WindowStyle Hidden `
            -PassThru
        [void]$launcher.Handle

        try {
            [void]$launcher.WaitForInputIdle(5000)
        }
        catch {
            Start-Sleep -Milliseconds 750
        }

        if ($launcher.HasExited) {
            throw "The explicitly launched PowerPoint automation process exited early."
        }

        $launchStartTicks = [long]$launcher.StartTime.ToUniversalTime().Ticks
        $launchRecord = [PSCustomObject]@{
            Application = $null
            ProcessId = [int]$launcher.Id
            StartTimeUtcTicks = $launchStartTicks
            Owned = $true
        }

        $application = Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $launchRecord `
            -Description "Connect to explicitly launched PowerPoint" `
            -Operation { New-Object -ComObject PowerPoint.Application }

        $processId = [uint32]0
        $windowDeadline = [DateTime]::UtcNow.AddSeconds(5)

        while ($processId -eq 0 -and [DateTime]::UtcNow -lt $windowDeadline) {
            try {
                $windowHandle = [IntPtr]([long]$application.HWND)
                [void][IguanaTexBuildNative]::GetWindowThreadProcessId(
                    $windowHandle,
                    [ref]$processId
                )
            }
            catch {
            }

            if ($processId -eq 0) {
                Start-Sleep -Milliseconds 100
            }
        }

        if ($processId -eq 0) {
            throw "Could not resolve the PowerPoint process from its window handle."
        }

        if ([int]$processId -ne [int]$launcher.Id) {
            throw (
                "PowerPoint COM did not bind to the explicitly launched automation " +
                "process (launched $($launcher.Id), bound $processId). No existing " +
                "process will be used or terminated."
            )
        }

        $process = Get-Process -Id ([int]$processId) -ErrorAction Stop

        try {
            $startTimeUtcTicks = [long]$process.StartTime.ToUniversalTime().Ticks
        }
        finally {
            $process.Dispose()
        }

        $processRecord = Get-CimInstance -ClassName Win32_Process `
            -Filter "ProcessId = $processId" -ErrorAction Stop
        $commandLine = [string]$processRecord.CommandLine

        if (
            $startTimeUtcTicks -ne $launchStartTicks -or
            $commandLine -notmatch '(?i)/AUTOMATION'
        ) {
            throw (
                "The bound process identity or command line did not match the " +
                "PowerPoint automation process explicitly launched by this build."
            )
        }

        $applicationInfo = [PSCustomObject]@{
            Application = $application
            ProcessId = [int]$processId
            StartTimeUtcTicks = $startTimeUtcTicks
            Owned = $true
        }

        if ($ShowWindow) {
            $application.Visible = $MSO_TRUE
        }

        try {
            $application.AutomationSecurity = $MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        }
        catch {
            throw "Could not disable macros for Office automation: $($_.Exception.Message)"
        }

        try {
            $application.DisplayAlerts = $PP_ALERTS_NONE
        }
        catch {
            throw "Could not disable PowerPoint alerts: $($_.Exception.Message)"
        }

        return $applicationInfo
    }
    catch {
        if ($null -ne $applicationInfo) {
            try {
                [void](Close-OwnedPowerPointApplication $applicationInfo)
            }
            catch {
            }
        }
        elseif ($null -ne $application) {
            Release-ComObject $application
        }

        if ($null -ne $launchRecord) {
            [void](Stop-ExactOwnedPowerPointProcess $launchRecord)
        }

        Invoke-ComGarbageCollection
        throw
    }
    finally {
        if ($null -ne $launcher) {
            $launcher.Dispose()
        }
    }
}

function Close-OwnedPowerPointApplication {
    param([object]$ApplicationInfo)

    if ($null -eq $ApplicationInfo) {
        return $true
    }

    $application = $ApplicationInfo.Application

    $quitSucceeded = $true

    if ($null -ne $application -and $ApplicationInfo.Owned) {
        try {
            [void](Invoke-OwnedPowerPointOperation `
                -ApplicationInfo $ApplicationInfo `
                -Description "PowerPoint Quit" `
                -Operation { $application.Quit() })
        }
        catch {
            $quitSucceeded = $false
        }
    }

    Release-ComObject $application
    $ApplicationInfo.Application = $null
    Invoke-ComGarbageCollection

    $currentProcess = Get-Process -Id ([int]$ApplicationInfo.ProcessId) `
        -ErrorAction SilentlyContinue

    if ($null -eq $currentProcess) {
        return $quitSucceeded
    }

    try {
        if (
            $currentProcess.ProcessName -ne "POWERPNT" -or
            [long]$currentProcess.StartTime.ToUniversalTime().Ticks -ne
                [long]$ApplicationInfo.StartTimeUtcTicks
        ) {
            return $false
        }

        [void]$currentProcess.WaitForExit(5000)

        if ($currentProcess.HasExited) {
            return $quitSucceeded
        }
    }
    finally {
        $currentProcess.Dispose()
    }

    $stopped = Stop-ExactOwnedPowerPointProcess -ApplicationInfo $ApplicationInfo
    return ($quitSucceeded -and $stopped)
}

function Get-VbaReferenceSnapshot {
    param([object]$Project)

    $result = @()
    $references = $null

    try {
        $references = $Project.References

        for ($index = 1; $index -le $references.Count; $index++) {
            $reference = $null

            try {
                $reference = $references.Item($index)
                $isBroken = [bool]$reference.IsBroken
                $name = $null
                $guid = $null
                $major = $null
                $minor = $null

                try { $name = [string]$reference.Name } catch { }
                try { $guid = [string]$reference.Guid } catch { }
                try { $major = [int]$reference.Major } catch { }
                try { $minor = [int]$reference.Minor } catch { }

                $result += [PSCustomObject]@{
                    Name = $name
                    Guid = $guid
                    Major = $major
                    Minor = $minor
                    IsBroken = $isBroken
                }
            }
            finally {
                Release-ComObject $reference
            }
        }
    }
    finally {
        Release-ComObject $references
    }

    return @($result)
}

function Add-CanonicalExplicitReferences {
    param(
        [object]$Project,
        [object[]]$ExplicitReferences
    )

    $references = $null

    try {
        $references = $Project.References

        foreach ($expected in $ExplicitReferences) {
            $added = $null

            try {
                $added = $references.AddFromGuid(
                    [string]$expected.guid,
                    [int]$expected.major,
                    [int]$expected.minor
                )

                if ([bool]$added.IsBroken) {
                    throw "Reference was broken immediately after AddFromGuid: $($expected.name)"
                }

                if (
                    [string]$added.Guid -ne [string]$expected.guid -or
                    [int]$added.Major -ne [int]$expected.major -or
                    [int]$added.Minor -ne [int]$expected.minor
                ) {
                    throw "Reference identity mismatch after AddFromGuid: $($expected.name)"
                }

                Write-Host (
                    "Reference added: {0} {1} {2}.{3}" -f
                    ([string]$added.Name),
                    ([string]$added.Guid),
                    ([int]$added.Major),
                    ([int]$added.Minor)
                )
            }
            finally {
                Release-ComObject $added
            }
        }
    }
    finally {
        Release-ComObject $references
    }
}

function Assert-VbaReferences {
    param(
        [object]$Project,
        [object[]]$ExplicitReferences,
        [switch]$RequireMsForms
    )

    $actual = @(Get-VbaReferenceSnapshot -Project $Project)
    $broken = @($actual | Where-Object { $_.IsBroken })

    if ($broken.Count -gt 0) {
        $names = @($broken | ForEach-Object {
            if ($null -ne $_.Name) { $_.Name } else { $_.Guid }
        }) -join ", "
        throw "Broken VBA reference(s): $names"
    }

    $scriptingRuntime = @($actual | Where-Object {
        $_.Guid -eq $SCRIPTING_RUNTIME_GUID -or $_.Name -eq "Scripting"
    })

    if ($scriptingRuntime.Count -gt 0) {
        throw "Microsoft Scripting Runtime must not be a VBA project reference."
    }

    foreach ($defaultName in @("VBA", "PowerPoint", "stdole", "Office")) {
        $matches = @($actual | Where-Object {
            $_.Name -eq $defaultName -and -not $_.IsBroken
        })

        if ($matches.Count -ne 1) {
            throw "Expected one healthy default VBA reference named '$defaultName'."
        }
    }

    foreach ($expected in $ExplicitReferences) {
        $matches = @($actual | Where-Object {
            $_.Guid -eq [string]$expected.guid -and
            $_.Major -eq [int]$expected.major -and
            $_.Minor -eq [int]$expected.minor -and
            -not $_.IsBroken
        })

        if ($matches.Count -ne 1) {
            throw (
                "Expected one healthy explicit reference: {0} {1} {2}.{3}" -f
                ([string]$expected.name),
                ([string]$expected.guid),
                ([int]$expected.major),
                ([int]$expected.minor)
            )
        }
    }

    if ($RequireMsForms) {
        $msForms = @($actual | Where-Object {
            $_.Name -eq "MSForms" -and -not $_.IsBroken
        })

        if ($msForms.Count -ne 1) {
            throw (
                "MSForms was not added automatically by UserForm import, " +
                "or the resulting reference is broken."
            )
        }
    }

    return [PSCustomObject]@{
        Count = $actual.Count
        Names = @($actual | ForEach-Object { $_.Name })
    }
}

function Assert-VbaProjectMetadata {
    param(
        [object]$Project,
        [object]$Metadata
    )

    if ([string]$Project.Name -ne [string]$Metadata.name) {
        throw "VBProject name does not match canonical metadata."
    }

    if ([string]$Project.Description -ne [string]$Metadata.description) {
        throw "VBProject description does not match canonical metadata."
    }

    if ([int]$Project.Protection -ne [int]$Metadata.protection) {
        throw "VBProject protection state does not match canonical metadata."
    }
}

function Set-AndAssert-VbaProjectMetadata {
    param(
        [object]$Project,
        [object]$Metadata
    )

    $Project.Name = [string]$Metadata.name
    $Project.Description = [string]$Metadata.description
    Assert-VbaProjectMetadata -Project $Project -Metadata $Metadata
}

function Assert-VbaProjectComponents {
    param(
        [object]$Project,
        [string]$SourceDirectory
    )

    $expected = @{}

    foreach ($file in @(Get-VbaSourceFiles -Directory $SourceDirectory)) {
        $type = switch ($file.Extension.ToLowerInvariant()) {
            ".bas" { $VBEXT_CT_STDMODULE }
            ".cls" { $VBEXT_CT_CLASSMODULE }
            ".frm" { $VBEXT_CT_MSFORM }
            default { throw "Unsupported canonical VBA extension: $($file.Extension)" }
        }

        $expected[$file.BaseName] = $type
    }

    $actual = @{}
    $components = $null

    try {
        $components = $Project.VBComponents

        for ($index = 1; $index -le $components.Count; $index++) {
            $component = $null

            try {
                $component = $components.Item($index)
                $name = [string]$component.Name
                $type = [int]$component.Type

                if ($type -eq $VBEXT_CT_DOCUMENT) {
                    throw "Unexpected PowerPoint host/document component: $name"
                }

                if ($type -notin @(
                    $VBEXT_CT_STDMODULE,
                    $VBEXT_CT_CLASSMODULE,
                    $VBEXT_CT_MSFORM
                )) {
                    throw "Unsupported VBComponent in fresh project: $name (type $type)"
                }

                $actual[$name] = $type
            }
            finally {
                Release-ComObject $component
            }
        }
    }
    finally {
        Release-ComObject $components
    }

    if ($actual.Count -ne $expected.Count) {
        throw (
            "Fresh VBProject component count mismatch: expected {0}, actual {1}." -f
            $expected.Count,
            $actual.Count
        )
    }

    foreach ($name in $expected.Keys) {
        if (-not $actual.ContainsKey($name)) {
            throw "Canonical VBA component was not imported: $name"
        }

        if ([int]$actual[$name] -ne [int]$expected[$name]) {
            throw "Imported VBA component type mismatch: $name"
        }
    }

    return [PSCustomObject]@{
        Count = $actual.Count
        Names = @($actual.Keys | Sort-Object)
    }
}

function Assert-OpenVbaProject {
    param(
        [object]$Project,
        [object]$Configuration,
        [object]$Layout
    )

    Assert-VbaProjectMetadata `
        -Project $Project `
        -Metadata $Configuration.Metadata

    [void](Assert-VbaProjectComponents `
        -Project $Project `
        -SourceDirectory $Layout.SourceDirectory)

    [void](Assert-VbaReferences `
        -Project $Project `
        -ExplicitReferences $Configuration.ExplicitReferences `
        -RequireMsForms)
}

function New-FreshPptmScaffold {
    param(
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout,
        [switch]$ShowWindow
    )

    [void](Assert-VbaSourceClosure -SourceDirectory $Layout.SourceDirectory)

    $applicationInfo = $null
    $presentations = $null
    $presentation = $null
    $project = $null
    $saved = $false

    try {
        Write-Host "Creating a fresh PowerPoint presentation (no PPTM input)..."
        $applicationInfo = New-OwnedPowerPointApplication -ShowWindow:$ShowWindow
        $presentations = $applicationInfo.Application.Presentations
        $presentation = Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Create fresh presentation" `
            -Operation { $presentations.Add($MSO_FALSE) }
        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Initial PPTM SaveAs" `
            -Operation {
                $presentation.SaveAs(
                    $ArtifactPath,
                    $PP_SAVE_AS_OPEN_XML_PRESENTATION_MACRO_ENABLED
                )
            })

        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Configure and import canonical VBA" `
            -Operation {
                $buildProject = $null

                try {
                    try {
                        $buildProject = $presentation.VBProject
                    }
                    catch {
                        throw (
                            "Cannot access VBProject. Enable 'Trust access to the VBA " +
                            "project object model'. $($_.Exception.Message)"
                        )
                    }

                    Set-AndAssert-VbaProjectMetadata `
                        -Project $buildProject `
                        -Metadata $Configuration.Metadata

                    Add-CanonicalExplicitReferences `
                        -Project $buildProject `
                        -ExplicitReferences $Configuration.ExplicitReferences

                    Import-VbaSourceTree `
                        -Project $buildProject `
                        -SourceDirectory $Layout.SourceDirectory

                    [void](Assert-VbaProjectComponents `
                        -Project $buildProject `
                        -SourceDirectory $Layout.SourceDirectory)

                    [void](Assert-VbaReferences `
                        -Project $buildProject `
                        -ExplicitReferences $Configuration.ExplicitReferences `
                        -RequireMsForms)
                }
                finally {
                    Release-ComObject $buildProject
                }
            })

        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Save imported PPTM" `
            -Operation { $presentation.Save() })
        $saved = $true
        Write-Host "Fresh VBA container saved."
    }
    finally {
        $cleanupFailure = $null
        Release-ComObject $project

        if ($null -ne $presentation) {
            if (-not $saved) {
                try { $presentation.Saved = $MSO_TRUE } catch { }
            }

            try {
                [void](Invoke-OwnedPowerPointOperation `
                    -ApplicationInfo $applicationInfo `
                    -Description "Close fresh PPTM" `
                    -Operation { $presentation.Close() })
            }
            catch {
                $cleanupFailure = $_
            }

            Release-ComObject $presentation
        }

        Release-ComObject $presentations
        if (-not (Close-OwnedPowerPointApplication $applicationInfo)) {
            $cleanupFailure = "The build-owned PowerPoint process did not exit cleanly."
        }

        if ($null -ne $cleanupFailure) {
            throw $cleanupFailure
        }
    }
}

function Open-Assert-AndClosePresentation {
    param(
        [object]$Presentations,
        [object]$ApplicationInfo,
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout,
        [switch]$Save,
        [switch]$ReadOnly,
        [switch]$ShowWindow
    )

    $presentation = $null
    $project = $null

    try {
        $readOnlyValue = $MSO_FALSE
        if ($ReadOnly) { $readOnlyValue = $MSO_TRUE }

        $windowValue = $MSO_FALSE
        if ($ShowWindow) { $windowValue = $MSO_TRUE }

        $presentation = Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $ApplicationInfo `
            -Description "Open PowerPoint artifact" `
            -Operation {
                $Presentations.Open(
                    $ArtifactPath,
                    $readOnlyValue,
                    $MSO_FALSE,
                    $windowValue
                )
            }

        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $ApplicationInfo `
            -Description "Validate reopened VBProject" `
            -Operation {
                $reopenedProject = $null

                try {
                    try {
                        $reopenedProject = $presentation.VBProject
                    }
                    catch {
                        throw (
                            "Reopened presentation has no accessible VBProject: " +
                            $_.Exception.Message
                        )
                    }

                    Assert-OpenVbaProject `
                        -Project $reopenedProject `
                        -Configuration $Configuration `
                        -Layout $Layout
                }
                finally {
                    Release-ComObject $reopenedProject
                }
            })

        if ($Save) {
            [void](Invoke-OwnedPowerPointOperation `
                -ApplicationInfo $ApplicationInfo `
                -Description "Save reopened PPTM" `
                -Operation { $presentation.Save() })
        }
    }
    finally {
        $cleanupFailure = $null
        Release-ComObject $project

        if ($null -ne $presentation) {
            if (-not $Save) {
                try { $presentation.Saved = $MSO_TRUE } catch { }
            }

            try {
                [void](Invoke-OwnedPowerPointOperation `
                    -ApplicationInfo $ApplicationInfo `
                    -Description "Close reopened PowerPoint artifact" `
                    -Operation { $presentation.Close() })
            }
            catch {
                $cleanupFailure = $_
            }

            Release-ComObject $presentation
        }

        if ($null -ne $cleanupFailure) {
            throw $cleanupFailure
        }
    }
}

function Invoke-PowerPointRoundTripValidation {
    param(
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout,
        [switch]$SaveFirstOpen,
        [switch]$ShowWindow
    )

    $applicationInfo = $null
    $presentations = $null

    try {
        Write-Host "PowerPoint open/close/reopen validation..."
        $applicationInfo = New-OwnedPowerPointApplication -ShowWindow:$ShowWindow
        $presentations = $applicationInfo.Application.Presentations

        Open-Assert-AndClosePresentation `
            -Presentations $presentations `
            -ApplicationInfo $applicationInfo `
            -ArtifactPath $ArtifactPath `
            -Configuration $Configuration `
            -Layout $Layout `
            -Save:$SaveFirstOpen `
            -ReadOnly:(-not $SaveFirstOpen) `
            -ShowWindow:$ShowWindow

        [void](Assert-IguanaTexOfficePackage `
            -PackagePath $ArtifactPath `
            -RibbonDirectory $Layout.RibbonDirectory `
            -SourceDirectory $Layout.SourceDirectory)

        Open-Assert-AndClosePresentation `
            -Presentations $presentations `
            -ApplicationInfo $applicationInfo `
            -ArtifactPath $ArtifactPath `
            -Configuration $Configuration `
            -Layout $Layout `
            -ReadOnly `
            -ShowWindow:$ShowWindow
    }
    finally {
        Release-ComObject $presentations
        if (-not (Close-OwnedPowerPointApplication $applicationInfo)) {
            throw "The reopen-validation PowerPoint process did not exit cleanly."
        }
    }

    [void](Assert-IguanaTexOfficePackage `
        -PackagePath $ArtifactPath `
        -RibbonDirectory $Layout.RibbonDirectory `
        -SourceDirectory $Layout.SourceDirectory)
}

function Open-Assert-AndRemoveAddIn {
    param(
        [object]$Application,
        [object]$ApplicationInfo,
        [object]$AddIns,
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout
    )

    $addIn = $null
    $addInName = $null
    $vbe = $null
    $projects = $null
    $project = $null

    try {
        $addIn = Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $ApplicationInfo `
            -Description "Register generated PPAM" `
            -Operation { $AddIns.Add($ArtifactPath) }
        $addInName = [string]$addIn.Name
        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $ApplicationInfo `
            -Description "Load generated PPAM" `
            -Operation { $addIn.Loaded = $MSO_TRUE })

        if (-not [bool]$addIn.Loaded) {
            throw "PowerPoint did not load the generated PPAM."
        }

        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $ApplicationInfo `
            -Description "Validate loaded PPAM VBProject" `
            -Operation {
                $validationVbe = $null
                $validationProjects = $null
                $validationProject = $null

                try {
                    $validationVbe = $Application.VBE
                    $validationProjects = $validationVbe.VBProjects
                    $matchingIndexes = @()

                    for (
                        $index = 1;
                        $index -le $validationProjects.Count;
                        $index++
                    ) {
                        $candidate = $null

                        try {
                            $candidate = $validationProjects.Item($index)
                            $candidatePath = $null

                            try {
                                $candidatePath = [System.IO.Path]::GetFullPath(
                                    [string]$candidate.FileName
                                )
                            }
                            catch {
                            }

                            if (
                                $null -ne $candidatePath -and
                                $candidatePath.Equals(
                                    $ArtifactPath,
                                    [System.StringComparison]::OrdinalIgnoreCase
                                )
                            ) {
                                $matchingIndexes += $index
                            }
                        }
                        finally {
                            Release-ComObject $candidate
                        }
                    }

                    if ($matchingIndexes.Count -ne 1) {
                        throw (
                            "Expected one loaded VBProject for the PPAM, found " +
                            $matchingIndexes.Count + "."
                        )
                    }

                    $validationProject = $validationProjects.Item(
                        [int]$matchingIndexes[0]
                    )

                    if (
                        [string]$validationProject.Name -ne
                            [string]$Configuration.Metadata.name
                    ) {
                        throw (
                            "Loaded PPAM VBProject name does not match " +
                            "canonical metadata."
                        )
                    }

                    if (
                        [string]$validationProject.Description -ne
                            [string]$Configuration.Metadata.description
                    ) {
                        throw (
                            "Loaded PPAM VBProject description does not match " +
                            "canonical metadata."
                        )
                    }

                    # A loaded PPAM is exposed as locked/unviewable even when
                    # its source PPTM was unprotected. References remain visible.
                    [void](Assert-VbaReferences `
                        -Project $validationProject `
                        -ExplicitReferences $Configuration.ExplicitReferences `
                        -RequireMsForms)
                }
                finally {
                    Release-ComObject $validationProject
                    Release-ComObject $validationProjects
                    Release-ComObject $validationVbe
                }
            })
    }
    finally {
        $cleanupFailure = $null
        Release-ComObject $project
        Release-ComObject $projects
        Release-ComObject $vbe

        if ($null -ne $addIn) {
            try {
                [void](Invoke-OwnedPowerPointOperation `
                    -ApplicationInfo $ApplicationInfo `
                    -Description "Unload generated PPAM" `
                    -Operation { $addIn.Loaded = $MSO_FALSE })

                if ([bool]$addIn.Loaded) {
                    throw "The generated PPAM remained loaded after unload."
                }
            }
            catch {
                $cleanupFailure = $_
            }

            if (-not [string]::IsNullOrWhiteSpace($addInName)) {
                try {
                    [void](Invoke-OwnedPowerPointOperation `
                        -ApplicationInfo $ApplicationInfo `
                        -Description "Remove generated PPAM registration" `
                        -Operation { $AddIns.Remove($addInName) })
                }
                catch {
                    $cleanupFailure = $_
                }
            }

            Release-ComObject $addIn
        }

        if (
            $null -eq $cleanupFailure -and
            -not [string]::IsNullOrWhiteSpace($addInName)
        ) {
            for ($index = 1; $index -le $AddIns.Count; $index++) {
                $remainingAddIn = $null

                try {
                    $remainingAddIn = $AddIns.Item($index)

                    if (
                        [string]$remainingAddIn.Name -eq $addInName -or
                        [string]$remainingAddIn.FullName -eq $ArtifactPath
                    ) {
                        $cleanupFailure = (
                            "The generated PPAM registration remained after removal."
                        )
                        break
                    }
                }
                finally {
                    Release-ComObject $remainingAddIn
                }
            }
        }

        if ($null -ne $cleanupFailure) {
            throw $cleanupFailure
        }
    }
}

function Invoke-PowerPointAddInValidation {
    param(
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout,
        [switch]$ShowWindow
    )

    $applicationInfo = $null
    $addIns = $null

    try {
        Write-Host "PowerPoint add-in load/unload/reload validation..."
        $applicationInfo = New-OwnedPowerPointApplication -ShowWindow:$ShowWindow
        $addIns = $applicationInfo.Application.AddIns

        for ($attempt = 1; $attempt -le 2; $attempt++) {
            Open-Assert-AndRemoveAddIn `
                -Application $applicationInfo.Application `
                -ApplicationInfo $applicationInfo `
                -AddIns $addIns `
                -ArtifactPath $ArtifactPath `
                -Configuration $Configuration `
                -Layout $Layout

            [void](Assert-IguanaTexOfficePackage `
                -PackagePath $ArtifactPath `
                -RibbonDirectory $Layout.RibbonDirectory `
                -SourceDirectory $Layout.SourceDirectory)
        }
    }
    finally {
        Release-ComObject $addIns
        if (-not (Close-OwnedPowerPointApplication $applicationInfo)) {
            throw "The add-in-validation PowerPoint process did not exit cleanly."
        }
    }
}

function Convert-PptmToFreshPpam {
    param(
        [string]$PptmPath,
        [string]$PpamPath,
        [switch]$ShowWindow
    )

    $applicationInfo = $null
    $presentations = $null
    $presentation = $null
    $saved = $false

    try {
        Write-Host "Converting PPTM to PowerPoint add-in format..."
        $applicationInfo = New-OwnedPowerPointApplication -ShowWindow:$ShowWindow
        $presentations = $applicationInfo.Application.Presentations

        $windowValue = $MSO_FALSE
        if ($ShowWindow) { $windowValue = $MSO_TRUE }

        $presentation = Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Open PPTM for PPAM conversion" `
            -Operation {
                $presentations.Open(
                    $PptmPath,
                    $MSO_FALSE,
                    $MSO_FALSE,
                    $windowValue
                )
            }
        [void](Invoke-OwnedPowerPointOperation `
            -ApplicationInfo $applicationInfo `
            -Description "Save PowerPoint add-in" `
            -Operation {
                $presentation.SaveAs(
                    $PpamPath,
                    $PP_SAVE_AS_OPEN_XML_ADDIN
                )
            })
        $saved = $true
    }
    finally {
        $cleanupFailure = $null

        if ($null -ne $presentation) {
            if (-not $saved) {
                try { $presentation.Saved = $MSO_TRUE } catch { }
            }

            try {
                [void](Invoke-OwnedPowerPointOperation `
                    -ApplicationInfo $applicationInfo `
                    -Description "Close PPAM conversion presentation" `
                    -Operation { $presentation.Close() })
            }
            catch {
                $cleanupFailure = $_
            }

            Release-ComObject $presentation
        }

        Release-ComObject $presentations
        if (-not (Close-OwnedPowerPointApplication $applicationInfo)) {
            $cleanupFailure = (
                "The PPAM-conversion PowerPoint process did not exit cleanly."
            )
        }

        if ($null -ne $cleanupFailure) {
            throw $cleanupFailure
        }
    }
}

function New-StagingArtifactPath {
    param(
        [string]$DestinationPath,
        [string]$RequiredExtension
    )

    if ([System.IO.Path]::GetExtension($DestinationPath) -ne $RequiredExtension) {
        throw "Output must use the $RequiredExtension extension: $DestinationPath"
    }

    $parent = Split-Path -Parent $DestinationPath

    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw "Output path must have a parent directory: $DestinationPath"
    }

    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $parent)
    }

    # PowerPoint AddIns rejects some otherwise valid long filenames. Keep the
    # disposable name short while retaining enough entropy for parallel builds.
    $token = [Guid]::NewGuid().ToString("N").Substring(0, 16)
    $name = "IT-{0}{1}" -f (
        $token,
        $RequiredExtension
    )

    return Join-Path $parent $name
}

function Assert-OutputMayBeWritten {
    param(
        [string]$DestinationPath,
        [switch]$Overwrite
    )

    if (Test-Path -LiteralPath $DestinationPath -PathType Container) {
        throw "Output path is a directory: $DestinationPath"
    }

    if ((Test-Path -LiteralPath $DestinationPath -PathType Leaf) -and -not $Overwrite) {
        throw "Output already exists. Pass -Force to replace it: $DestinationPath"
    }
}

function Publish-StagingArtifact {
    param(
        [string]$StagingPath,
        [string]$DestinationPath,
        [switch]$Overwrite
    )

    Assert-OutputMayBeWritten -DestinationPath $DestinationPath -Overwrite:$Overwrite
    Move-Item -LiteralPath $StagingPath -Destination $DestinationPath -Force:$Overwrite
}

function Remove-GeneratedStagingFile {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        $leaf = [System.IO.Path]::GetFileName($Path)

        if ($leaf -notmatch '^IT-[0-9a-f]{16}\.(?:pptm|ppam)$') {
            throw "Refusing to remove a path that is not a build staging file: $Path"
        }

        if ([System.IO.File]::Exists($Path)) {
            [System.IO.File]::Delete($Path)
        }
    }
    catch {
        Write-Warning "Could not remove staging file '$Path': $($_.Exception.Message)"
    }
}

function Remove-ValidationTemporaryDirectory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        $tempRoot = [System.IO.Path]::GetFullPath(
            [System.IO.Path]::GetTempPath()
        )
        $resolved = [System.IO.Path]::GetFullPath($Path)
        $leaf = Split-Path -Leaf $resolved

        if (
            -not $resolved.StartsWith(
                $tempRoot,
                [System.StringComparison]::OrdinalIgnoreCase
            ) -or
            -not $leaf.StartsWith(
                "IguanaTex-office-validate-",
                [System.StringComparison]::Ordinal
            )
        ) {
            throw "Refusing to recursively remove an unverified path: $resolved"
        }

        if (Test-Path -LiteralPath $resolved -PathType Container) {
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
    catch {
        Write-Warning (
            "Could not remove validation directory '$Path': " +
            $_.Exception.Message
        )
    }
}

function Assert-CompilePassed {
    param([object]$Result)

    if (-not $Result.Passed) {
        throw (
            "VBE compile validation failed [{0}]: {1}" -f
            ([string]$Result.Status),
            ([string]$Result.Message)
        )
    }

    Write-Host (
        "VBE compile PASS ({0}; Enabled {1} -> {2})" -f
        ([string]$Result.Status),
        ([string]$Result.EnabledBefore),
        ([string]$Result.EnabledAfter)
    )
}

function Invoke-FullPptmBuild {
    param(
        [string]$DestinationPath,
        [string]$OptionalPpamDestination,
        [object]$Configuration,
        [object]$Layout,
        [switch]$SkipValidation
    )

    Assert-OutputMayBeWritten -DestinationPath $DestinationPath -Overwrite:$Force

    if (-not [string]::IsNullOrWhiteSpace($OptionalPpamDestination)) {
        Assert-OutputMayBeWritten `
            -DestinationPath $OptionalPpamDestination `
            -Overwrite:$Force
    }

    $pptmStage = New-StagingArtifactPath `
        -DestinationPath $DestinationPath `
        -RequiredExtension ".pptm"
    $ppamStage = $null

    try {
        New-FreshPptmScaffold `
            -ArtifactPath $pptmStage `
            -Configuration $Configuration `
            -Layout $Layout `
            -ShowWindow:$Visible

        if (-not $SkipValidation) {
            $compileResult = Invoke-VbeCompileValidation `
                -PresentationPath $pptmStage `
                -TimeoutSeconds $CompileTimeoutSeconds
            Assert-CompilePassed $compileResult

            [void](Assert-VbaPackageClosure -PackagePath $pptmStage)
        }

        [void](Add-IguanaTexRibbon `
            -PackagePath $pptmStage `
            -RibbonDirectory $Layout.RibbonDirectory)

        if (-not $SkipValidation) {
            [void](Assert-IguanaTexOfficePackage `
                -PackagePath $pptmStage `
                -RibbonDirectory $Layout.RibbonDirectory `
                -SourceDirectory $Layout.SourceDirectory)

            Invoke-PowerPointRoundTripValidation `
                -ArtifactPath $pptmStage `
                -Configuration $Configuration `
                -Layout $Layout `
                -SaveFirstOpen `
                -ShowWindow:$Visible
        }

        if (-not [string]::IsNullOrWhiteSpace($OptionalPpamDestination)) {
            $ppamStage = New-StagingArtifactPath `
                -DestinationPath $OptionalPpamDestination `
                -RequiredExtension ".ppam"

            Invoke-PpamBuild `
                -SourcePptmPath $pptmStage `
                -PpamStagePath $ppamStage `
                -Configuration $Configuration `
                -Layout $Layout `
                -SkipValidation:$SkipValidation
        }

        Publish-StagingArtifact `
            -StagingPath $pptmStage `
            -DestinationPath $DestinationPath `
            -Overwrite:$Force
        $pptmStage = $null

        if ($null -ne $ppamStage) {
            Publish-StagingArtifact `
                -StagingPath $ppamStage `
                -DestinationPath $OptionalPpamDestination `
                -Overwrite:$Force
            $ppamStage = $null
        }
    }
    finally {
        Remove-GeneratedStagingFile $pptmStage
        Remove-GeneratedStagingFile $ppamStage
    }
}

function Invoke-PpamBuild {
    param(
        [string]$SourcePptmPath,
        [string]$PpamStagePath,
        [object]$Configuration,
        [object]$Layout,
        [switch]$SkipValidation
    )

    Convert-PptmToFreshPpam `
        -PptmPath $SourcePptmPath `
        -PpamPath $PpamStagePath `
        -ShowWindow:$Visible

    if (-not $SkipValidation) {
        [void](Assert-VbaPackageClosure -PackagePath $PpamStagePath)
    }

    # SaveAs may preserve, remove, or rewrite customUI parts depending on the
    # Office build. Every final Office save is deliberately followed by an
    # idempotent canonical injection.
    [void](Add-IguanaTexRibbon `
        -PackagePath $PpamStagePath `
        -RibbonDirectory $Layout.RibbonDirectory)

    if (-not $SkipValidation) {
        [void](Assert-IguanaTexOfficePackage `
            -PackagePath $PpamStagePath `
            -RibbonDirectory $Layout.RibbonDirectory `
            -SourceDirectory $Layout.SourceDirectory)

        Invoke-PowerPointAddInValidation `
            -ArtifactPath $PpamStagePath `
            -Configuration $Configuration `
            -Layout $Layout `
            -ShowWindow:$Visible
    }
}

function Invoke-IsolatedArtifactValidation {
    param(
        [string]$ArtifactPath,
        [object]$Configuration,
        [object]$Layout
    )

    [void](Assert-IguanaTexOfficePackage `
        -PackagePath $ArtifactPath `
        -RibbonDirectory $Layout.RibbonDirectory `
        -SourceDirectory $Layout.SourceDirectory)

    $temporaryRoot = Join-Path (
        [System.IO.Path]::GetTempPath()
    ) ("IguanaTex-office-validate-" + [Guid]::NewGuid().ToString("N"))
    $artifactExtension = [System.IO.Path]::GetExtension($ArtifactPath)
    $temporaryName = [System.IO.Path]::GetFileName($ArtifactPath)

    if ($artifactExtension -eq ".ppam") {
        # AddIns.Remove addresses registrations by name. A randomized validation
        # basename prevents collision with an unrelated installed add-in.
        $temporaryName = "IT-{0}.ppam" -f (
            [Guid]::NewGuid().ToString("N").Substring(0, 16)
        )
    }

    $temporaryArtifact = Join-Path $temporaryRoot $temporaryName

    try {
        [void](New-Item -ItemType Directory -Path $temporaryRoot)
        Copy-Item -LiteralPath $ArtifactPath -Destination $temporaryArtifact

        if ($artifactExtension -eq ".pptm") {
            $compileResult = Invoke-VbeCompileValidation `
                -PresentationPath $temporaryArtifact `
                -TimeoutSeconds $CompileTimeoutSeconds `
                -AllowAlreadyCompiled
            Assert-CompilePassed $compileResult

            Invoke-PowerPointRoundTripValidation `
                -ArtifactPath $temporaryArtifact `
                -Configuration $Configuration `
                -Layout $Layout `
                -SaveFirstOpen `
                -ShowWindow:$Visible
        }
        else {
            Invoke-PowerPointAddInValidation `
                -ArtifactPath $temporaryArtifact `
                -Configuration $Configuration `
                -Layout $Layout `
                -ShowWindow:$Visible
        }
    }
    finally {
        Remove-ValidationTemporaryDirectory $temporaryRoot
    }
}

$exitCode = 0

try {
    $layout = Get-CanonicalLayout
    $configuration = Read-CanonicalConfiguration -Layout $layout
    [void](Assert-VbaSourceClosure -SourceDirectory $layout.SourceDirectory)

    $defaultOutput = Join-Path $projectRoot ".build\office\IguanaTeX.pptm"
    $defaultPpamOutput = Join-Path $projectRoot ".build\office\IguanaTeX.ppam"

    switch ($Action) {
        "build" {
            if (-not [string]::IsNullOrWhiteSpace($InputPath)) {
                throw "-InputPath is not valid for a fresh build."
            }

            if ([string]::IsNullOrWhiteSpace($OutputPath)) {
                $OutputPath = $defaultOutput
            }

            $destination = ConvertTo-AbsolutePath $OutputPath
            $ppamDestination = $null

            if (-not [string]::IsNullOrWhiteSpace($PpamOutputPath)) {
                $ppamDestination = ConvertTo-AbsolutePath $PpamOutputPath
            }

            Write-Host "Action: fresh PPTM build"
            Write-Host "Source: $($layout.SourceDirectory)"
            Write-Host "Output: $destination"
            if ($null -ne $ppamDestination) {
                Write-Host "PPAM:   $ppamDestination"
            }
            if ($NoValidation) {
                Write-Warning (
                    "Artifact validation is disabled. Outputs are unverified " +
                    "and should not be distributed."
                )
            }
            Write-Host ""

            Invoke-FullPptmBuild `
                -DestinationPath $destination `
                -OptionalPpamDestination $ppamDestination `
                -Configuration $configuration `
                -Layout $layout `
                -SkipValidation:$NoValidation

            Write-Host ""
            if ($NoValidation) {
                Write-Host "BUILD OK (VALIDATION SKIPPED): $destination"
            }
            else {
                Write-Host "BUILD OK: $destination"
            }
            if ($null -ne $ppamDestination) {
                if ($NoValidation) {
                    Write-Host "PPAM OK (VALIDATION SKIPPED): $ppamDestination"
                }
                else {
                    Write-Host "PPAM OK:  $ppamDestination"
                }
            }
        }

        "validate" {
            if ([string]::IsNullOrWhiteSpace($InputPath)) {
                throw "validate requires -InputPath."
            }

            if (
                -not [string]::IsNullOrWhiteSpace($OutputPath) -or
                -not [string]::IsNullOrWhiteSpace($PpamOutputPath) -or
                $Force -or
                $NoValidation
            ) {
                throw (
                    "validate does not accept output, overwrite, or " +
                    "-NoValidation options."
                )
            }

            $input = ConvertTo-AbsolutePath $InputPath -MustExist
            $extension = [System.IO.Path]::GetExtension($input).ToLowerInvariant()

            if ($extension -notin @(".pptm", ".ppam")) {
                throw "Validation input must be a .pptm or .ppam file: $input"
            }

            Write-Host "Action: isolated Office artifact validation"
            Write-Host "Input:  $input"
            Write-Host ""

            Invoke-IsolatedArtifactValidation `
                -ArtifactPath $input `
                -Configuration $configuration `
                -Layout $layout

            Write-Host ""
            Write-Host "VALIDATE OK: $input"
        }

        "ppam" {
            if ([string]::IsNullOrWhiteSpace($InputPath)) {
                throw "ppam requires -InputPath pointing to a PPTM."
            }

            if (-not [string]::IsNullOrWhiteSpace($PpamOutputPath)) {
                throw "Use -OutputPath, not -PpamOutputPath, with the ppam action."
            }

            if ([string]::IsNullOrWhiteSpace($OutputPath)) {
                $OutputPath = $defaultPpamOutput
            }

            $input = ConvertTo-AbsolutePath $InputPath -MustExist
            $destination = ConvertTo-AbsolutePath $OutputPath

            if ([System.IO.Path]::GetExtension($input) -ne ".pptm") {
                throw "PPAM source must be a .pptm file: $input"
            }

            Assert-OutputMayBeWritten -DestinationPath $destination -Overwrite:$Force

            Write-Host "Action: PPAM build"
            Write-Host "Input:  $input"
            Write-Host "Output: $destination"
            if ($NoValidation) {
                Write-Warning (
                    "Artifact validation is disabled. The output is unverified " +
                    "and should not be distributed."
                )
            }
            Write-Host ""

            if (-not $NoValidation) {
                Invoke-IsolatedArtifactValidation `
                    -ArtifactPath $input `
                    -Configuration $configuration `
                    -Layout $layout
            }

            $stage = New-StagingArtifactPath `
                -DestinationPath $destination `
                -RequiredExtension ".ppam"

            try {
                Invoke-PpamBuild `
                    -SourcePptmPath $input `
                    -PpamStagePath $stage `
                    -Configuration $configuration `
                    -Layout $layout `
                    -SkipValidation:$NoValidation

                Publish-StagingArtifact `
                    -StagingPath $stage `
                    -DestinationPath $destination `
                    -Overwrite:$Force
                $stage = $null
            }
            finally {
                Remove-GeneratedStagingFile $stage
            }

            Write-Host ""
            if ($NoValidation) {
                Write-Host "PPAM OK (VALIDATION SKIPPED): $destination"
            }
            else {
                Write-Host "PPAM OK: $destination"
            }
        }
    }
}
catch {
    Write-Error $_
    $exitCode = 1
}

exit $exitCode

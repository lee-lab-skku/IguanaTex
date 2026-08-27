#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PresentationPath,
    [Parameter(Mandatory = $true)][string]$StateDirectory,
    [Parameter(Mandatory = $true)][string]$PreexistingProcessesPath,
    [Parameter(Mandatory = $true)][string]$WatcherPath,
    [ValidateRange(5, 3600)][int]$TimeoutSeconds = 120,
    [switch]$AllowAlreadyCompiled
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class IguanaTexCompileWorkerNative
{
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId
    );
}
'@

function ConvertTo-PowerShellSingleQuotedLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return "'" + $Value.Replace("'", "''") + "'"
}

function Write-JsonFileAtomic {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $temporaryPath = $Path + "." + $PID + ".tmp"
    $json = $Value | ConvertTo-Json -Depth 8 -Compress
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($temporaryPath, $json, $encoding)
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
}

function Write-MarkerFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, [string][DateTime]::UtcNow.Ticks, $encoding)
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }

    try {
        return [IO.File]::ReadAllText($Path) | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Get-PowerPointProcessIds {
    return @(
        Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    [int]$_.Id
                }
                finally {
                    $_.Dispose()
                }
            }
    )
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

function Release-ComObjectSafely {
    param([object]$Object)

    if ($null -eq $Object) {
        return $true
    }

    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }

        return $true
    }
    catch {
        return $false
    }
}

function Invoke-ComOperationWithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    $lastException = $null

    while ([DateTime]::UtcNow -lt $DeadlineUtc) {
        try {
            return [pscustomobject]@{
                Succeeded = $true
                Value     = (& $Operation)
                Exception = $null
            }
        }
        catch {
            $lastException = $_.Exception
            Start-Sleep -Milliseconds 150
        }
    }

    return [pscustomobject]@{
        Succeeded = $false
        Value     = $null
        Exception = $lastException
    }
}

function Get-ApplicationWindowProcessId {
    param(
        [Parameter(Mandatory = $true)][object]$Application,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc
    )

    while ([DateTime]::UtcNow -lt $DeadlineUtc) {
        try {
            $rawHandle = [Int64]$Application.HWND

            if ($rawHandle -lt 0) {
                $rawHandle = ([Int64][UInt32]::MaxValue + 1) + $rawHandle
            }

            if ($rawHandle -ne 0) {
                $processId = [UInt32]0
                [void][IguanaTexCompileWorkerNative]::GetWindowThreadProcessId(
                    [IntPtr]$rawHandle,
                    [ref]$processId
                )

                if ($processId -ne 0) {
                    return [int]$processId
                }
            }
        }
        catch {
        }

        Start-Sleep -Milliseconds 100
    }

    return $null
}

function Get-ExactOwnedPowerPointProcess {
    param(
        [int]$ProcessId,
        [Int64]$StartTimeUtcTicks,
        [bool]$Owned
    )

    if (-not $Owned -or $ProcessId -le 0) {
        return $null
    }

    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return $null
    }

    try {
        if ($process.ProcessName -ne "POWERPNT" -or
            [Int64]$process.StartTime.ToUniversalTime().Ticks -ne $StartTimeUtcTicks) {
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
    param(
        [int]$ProcessId,
        [Int64]$StartTimeUtcTicks,
        [bool]$Owned
    )

    $process = Get-ExactOwnedPowerPointProcess `
        -ProcessId $ProcessId `
        -StartTimeUtcTicks $StartTimeUtcTicks `
        -Owned $Owned

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

function Test-ExactOwnedPowerPointExited {
    param(
        [int]$ProcessId,
        [Int64]$StartTimeUtcTicks,
        [bool]$Owned,
        [int]$WaitMilliseconds = 0
    )

    if (-not $Owned) {
        return $true
    }

    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return $true
    }

    try {
        if ($process.ProcessName -ne "POWERPNT" -or
            [Int64]$process.StartTime.ToUniversalTime().Ticks -ne
                $StartTimeUtcTicks) {
            # PID reuse or an unverifiable identity is never considered a clean
            # validator shutdown and is never eligible for termination.
            return $false
        }

        if ($WaitMilliseconds -gt 0) {
            [void]$process.WaitForExit($WaitMilliseconds)
        }

        return $process.HasExited
    }
    catch {
        return $false
    }
    finally {
        $process.Dispose()
    }
}

function Start-CompileWatcher {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$Directory,
        [int]$PowerPointProcessId,
        [Int64]$PowerPointStartTimeUtcTicks,
        [int]$Timeout
    )

    $scriptLiteral = ConvertTo-PowerShellSingleQuotedLiteral $ScriptPath
    $stateLiteral = ConvertTo-PowerShellSingleQuotedLiteral $Directory
    $command = (
        "& $scriptLiteral -StateDirectory $stateLiteral " +
        "-PowerPointProcessId $PowerPointProcessId " +
        "-PowerPointStartTimeUtcTicks $PowerPointStartTimeUtcTicks " +
        "-TimeoutSeconds $Timeout"
    )
    $encodedCommand = [Convert]::ToBase64String(
        [Text.Encoding]::Unicode.GetBytes($command)
    )
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $PSHOME "powershell.exe"
    $startInfo.Arguments = (
        "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass " +
        "-STA -EncodedCommand $encodedCommand"
    )
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    if (-not $process.Start()) {
        $process.Dispose()
        throw "Failed to start compile dialog watcher."
    }

    return $process
}

function Wait-ForFileUntil {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][DateTime]$DeadlineUtc,
        [System.Diagnostics.Process]$Process
    )

    while ([DateTime]::UtcNow -lt $DeadlineUtc) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return $true
        }

        if ($null -ne $Process -and $Process.HasExited) {
            return $false
        }

        Start-Sleep -Milliseconds 100
    }

    return $false
}

$resultPath = Join-Path $StateDirectory "result.json"
$ownerPath = Join-Path $StateDirectory "powerpoint-owner.json"
$watcherReadyPath = Join-Path $StateDirectory "watcher.ready"
$watcherOutcomePath = Join-Path $StateDirectory "watcher-outcome.json"
$watcherStopPath = Join-Path $StateDirectory "watcher.stop"
$compileStartPath = Join-Path $StateDirectory "compile.start"
$executeReturnedPath = Join-Path $StateDirectory "execute.returned"
$deadlineUtc = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$result = [ordered]@{
    Passed                       = $false
    Status                       = "WorkerFailed"
    EnabledBefore                = $null
    EnabledAfter                 = $null
    DialogDetected               = $false
    CleanupSucceeded             = $false
    OwnedPowerPointProcessId      = $null
    Message                      = "Compile worker did not complete."
}
$ppt = $null
$presentations = $null
$presentation = $null
$vbe = $null
$mainWindow = $null
$commandBars = $null
$command = $null
$watcher = $null
$launcher = $null
$powerPointApplicationCreated = $false
$mappedProcessWasProtected = $false
$ownedPowerPoint = $false
$comBoundToOwnedPowerPoint = $false
$powerPointProcessId = 0
$boundPowerPointProcessId = 0
$powerPointStartTimeUtcTicks = [Int64]0
$activationStartedUtcTicks = [Int64]0
$explicitLaunchVerified = $false
$automationCommandLineVerified = $false
$compileSucceeded = $false
$cleanupOperationsSucceeded = $true

try {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne
        [Threading.ApartmentState]::STA) {
        throw "Compile worker must run in an STA Windows PowerShell process."
    }

    if (-not [IO.Path]::IsPathRooted($PresentationPath) -or
        -not (Test-Path -LiteralPath $PresentationPath -PathType Leaf)) {
        throw "Presentation does not exist or is not absolute: $PresentationPath"
    }

    if (-not (Test-Path -LiteralPath $WatcherPath -PathType Leaf)) {
        throw "Compile dialog watcher was not found: $WatcherPath"
    }

    $controllerSnapshot = Read-JsonFile $PreexistingProcessesPath
    $protectedProcessIds = @()

    if ($null -ne $controllerSnapshot) {
        $protectedProcessIds += @(
            $controllerSnapshot.Processes | ForEach-Object { [int]$_.Id }
        )
    }

    # This second snapshot closes the controller-to-worker launch race.
    $protectedProcessIds += @(Get-PowerPointProcessIds)
    $protectedProcessIds = @($protectedProcessIds | Sort-Object -Unique)

    $activationStartedUtcTicks = [Int64][DateTime]::UtcNow.Ticks
    $powerPointPath = Get-PowerPointExecutablePath
    $launcher = Start-Process `
        -FilePath $powerPointPath `
        -ArgumentList "/AUTOMATION" `
        -WindowStyle Hidden `
        -PassThru
    [void]$launcher.Handle

    if ($launcher.HasExited) {
        $result.Status = "IsolationFailed"
        $result.Message = "The explicitly launched PowerPoint process exited early."
        throw $result.Message
    }

    $powerPointProcessId = [int]$launcher.Id
    $powerPointStartTimeUtcTicks = (
        [Int64]$launcher.StartTime.ToUniversalTime().Ticks
    )
    $explicitLaunchVerified = $true
    $ownedPowerPoint = $true
    $result.OwnedPowerPointProcessId = [int]$powerPointProcessId

    if ($protectedProcessIds -contains $powerPointProcessId) {
        $result.Status = "IsolationFailed"
        $result.Message = (
            "The explicitly launched PowerPoint PID collided with a protected " +
            "preexisting-process snapshot."
        )
        throw $result.Message
    }

    # Publish exact ownership before COM activation. If activation blocks, the
    # controller may terminate only this explicit Start-Process PID/start pair.
    Write-JsonFileAtomic -Path $ownerPath -Value ([pscustomobject][ordered]@{
        Owned                 = $true
        ExplicitLaunchVerified = $explicitLaunchVerified
        LaunchArguments       = "/AUTOMATION"
        ProcessId             = [int]$powerPointProcessId
        StartTimeUtcTicks     = [string]$powerPointStartTimeUtcTicks
        ActivationStartedUtcTicks = [string]$activationStartedUtcTicks
        AutomationCommandLineVerified = $false
        WorkerProcessId       = [int]$PID
        Reason                = "Explicit PowerPoint launch; COM binding pending."
    })

    $processMetadata = Get-CimInstance `
        -ClassName Win32_Process `
        -Filter ("ProcessId = " + [int]$powerPointProcessId) `
        -ErrorAction Stop
    $automationCommandLine = [string]$processMetadata.CommandLine
    $automationCommandLineVerified = (
        $automationCommandLine -match
            '(?i)(?:^|\s)[/-](?:automation|embedding)(?:\s|$)'
    )

    if (-not $automationCommandLineVerified) {
        $result.Status = "IsolationFailed"
        $result.Message = (
            "The explicitly launched PowerPoint command line did not retain " +
            "the /AUTOMATION argument."
        )
        throw $result.Message
    }

    try {
        [void]$launcher.WaitForInputIdle(5000)
    }
    catch {
        Start-Sleep -Milliseconds 750
    }

    if ($launcher.HasExited) {
        $result.Status = "IsolationFailed"
        $result.Message = (
            "The explicitly launched PowerPoint process exited before COM binding."
        )
        throw $result.Message
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    $powerPointApplicationCreated = $true
    $boundPowerPointProcessId = Get-ApplicationWindowProcessId `
        -Application $ppt `
        -DeadlineUtc ([DateTime]::UtcNow.AddSeconds(5))

    if ($null -eq $boundPowerPointProcessId) {
        $result.Status = "IsolationFailed"
        $result.Message = "Could not map the PowerPoint application HWND to a process."
        throw $result.Message
    }

    $mappedProcessWasProtected = (
        $protectedProcessIds -contains $boundPowerPointProcessId
    )

    if ([int]$boundPowerPointProcessId -ne [int]$powerPointProcessId) {
        $result.Status = "IsolationFailed"
        $result.Message = (
            "PowerPoint COM did not bind to the explicitly launched automation " +
            "process (launched $powerPointProcessId, bound " +
            "$boundPowerPointProcessId). The bound process will not be opened, " +
            "modified, quit, or terminated."
        )
        throw $result.Message
    }

    $boundProcess = Get-Process -Id $boundPowerPointProcessId -ErrorAction Stop

    try {
        if (
            $boundProcess.ProcessName -ne "POWERPNT" -or
            [Int64]$boundProcess.StartTime.ToUniversalTime().Ticks -ne
                $powerPointStartTimeUtcTicks
        ) {
            throw "The COM-bound PowerPoint process identity changed."
        }
    }
    finally {
        $boundProcess.Dispose()
    }

    $comBoundToOwnedPowerPoint = $true
    Write-JsonFileAtomic -Path $ownerPath -Value ([pscustomobject][ordered]@{
        Owned                 = $true
        ExplicitLaunchVerified = $explicitLaunchVerified
        LaunchArguments       = "/AUTOMATION"
        ProcessId             = [int]$powerPointProcessId
        StartTimeUtcTicks     = [string]$powerPointStartTimeUtcTicks
        ActivationStartedUtcTicks = [string]$activationStartedUtcTicks
        AutomationCommandLineVerified = $automationCommandLineVerified
        WorkerProcessId       = [int]$PID
        Reason                = (
            "COM HWND matched the explicitly launched PowerPoint PID/start time."
        )
    })

    $ppt.Visible = -1

    try {
        # msoAutomationSecurityForceDisable. The validator must never run target
        # macros merely by opening the artifact under test.
        $ppt.AutomationSecurity = 3
    }
    catch {
        $result.Status = "AutomationSecurityFailed"
        $result.Message = (
            "Could not force-disable macros before opening the target: " +
            $_.Exception.Message
        )
        throw $result.Message
    }

    $presentations = $ppt.Presentations
    $presentation = $presentations.Open($PresentationPath, 0, 0, -1)
    $vbe = $ppt.VBE
    $mainWindow = $vbe.MainWindow
    $mainWindow.Visible = $true
    $commandBars = $vbe.CommandBars
    $command = $commandBars.FindControl(1, 578)

    if ($null -eq $command) {
        $result.Status = "CompileCommandMissing"
        $result.Message = "VBE compile command bar control ID 578 was not found."
        throw $result.Message
    }

    $beforeRead = Invoke-ComOperationWithRetry `
        -Operation { [bool]$command.Enabled } `
        -DeadlineUtc ([DateTime]::UtcNow.AddSeconds(5))

    if (-not $beforeRead.Succeeded) {
        $result.Status = "ComRecoveryFailed"
        $result.Message = "Could not read the VBE compile command pre-state."
        throw $result.Message
    }

    $result.EnabledBefore = [bool]$beforeRead.Value

    if (-not $result.EnabledBefore) {
        $result.EnabledAfter = $false

        if (-not $AllowAlreadyCompiled) {
            $result.Status = "CompileCommandNotEnabled"
            $result.Message = (
                "The VBE compile command was disabled before execution; " +
                "a fresh compile could not be proven."
            )
        }
        else {
            try {
                $presentation.Save()
                $compileSucceeded = $true
                $result.Passed = $true
                $result.Status = "AlreadyCompiled"
                $result.Message = (
                    "The VBE compile command was already disabled and " +
                    "AllowAlreadyCompiled was specified."
                )
            }
            catch {
                $result.Status = "SaveFailed"
                $result.Message = (
                    "The already-compiled presentation could not be saved: " +
                    $_.Exception.Message
                )
            }
        }
    }
    else {
        $remainingSeconds = [Math]::Max(
            5,
            [int][Math]::Ceiling(($deadlineUtc - [DateTime]::UtcNow).TotalSeconds)
        )
        $watcher = Start-CompileWatcher `
            -ScriptPath $WatcherPath `
            -Directory $StateDirectory `
            -PowerPointProcessId $powerPointProcessId `
            -PowerPointStartTimeUtcTicks $powerPointStartTimeUtcTicks `
            -Timeout $remainingSeconds

        if (-not (Wait-ForFileUntil `
            -Path $watcherReadyPath `
            -DeadlineUtc $deadlineUtc `
            -Process $watcher)) {
            $result.Status = "WatcherFailed"
            $result.Message = "Compile dialog watcher did not become ready."
            throw $result.Message
        }

        Write-MarkerFile $compileStartPath
        $executeException = $null

        try {
            [void]$command.Execute()
        }
        catch {
            $executeException = $_.Exception
        }
        finally {
            Write-MarkerFile $executeReturnedPath
        }

        if (-not (Wait-ForFileUntil `
            -Path $watcherOutcomePath `
            -DeadlineUtc $deadlineUtc `
            -Process $watcher)) {
            $result.Status = "Timeout"
            $result.Message = "Compile dialog watcher did not produce an outcome."
            throw $result.Message
        }

        $watcherOutcome = Read-JsonFile $watcherOutcomePath

        if ($null -eq $watcherOutcome) {
            $result.Status = "WatcherFailed"
            $result.Message = "Compile dialog watcher produced an invalid outcome."
            throw $result.Message
        }

        $result.DialogDetected = [bool]$watcherOutcome.DialogDetected

        if ([string]$watcherOutcome.Status -eq "Timeout") {
            $result.Status = "Timeout"
            $result.Message = [string]$watcherOutcome.Message
            throw $result.Message
        }

        if ([string]$watcherOutcome.Status -in @(
            "ProcessExited",
            "WatcherFailed"
        )) {
            $result.Status = "ComRecoveryFailed"
            $result.Message = [string]$watcherOutcome.Message
            throw $result.Message
        }

        if ($null -ne $executeException -and -not $result.DialogDetected) {
            $result.Status = "CompileCommandFailed"
            $result.Message = (
                "The VBE compile command invocation failed: " +
                $executeException.Message
            )
            throw $result.Message
        }

        $afterRead = Invoke-ComOperationWithRetry `
            -Operation { [bool]$command.Enabled } `
            -DeadlineUtc $deadlineUtc

        if (-not $afterRead.Succeeded) {
            $result.Status = "ComRecoveryFailed"
            $result.Message = (
                "VBE COM state could not be recovered after compile validation."
            )
            throw $result.Message
        }

        $result.EnabledAfter = [bool]$afterRead.Value

        if ($result.DialogDetected) {
            $result.Status = "CompileErrorDialog"
            $result.Message = [string]$watcherOutcome.Message
        }
        elseif ($result.EnabledAfter) {
            $result.Status = "CompileCommandStillEnabled"
            $result.Message = (
                "The VBE compile command remained enabled after execution."
            )
        }
        else {
            try {
                $presentation.Save()
                $compileSucceeded = $true
                $result.Passed = $true
                $result.Status = "Passed"
                $result.Message = (
                    "VBE compile completed and command control ID 578 became disabled."
                )
            }
            catch {
                $result.Status = "SaveFailed"
                $result.Message = (
                    "VBE compile succeeded, but the presentation could not be saved: " +
                    $_.Exception.Message
                )
            }
        }
    }
}
catch {
    if ($result.Status -eq "WorkerFailed") {
        $result.Status = "WorkerFailed"
        $result.Message = $_.Exception.Message
    }
}
finally {
    if ($null -ne $watcher) {
        try {
            Write-MarkerFile $watcherStopPath

            if (-not $watcher.WaitForExit(2000)) {
                [void]$watcher.Handle
                $watcher.Kill()
                [void]$watcher.WaitForExit(2000)
            }
        }
        catch {
        }
        finally {
            try {
                $watcher.Dispose()
            }
            catch {
                $cleanupOperationsSucceeded = $false
            }
            $watcher = $null
        }
    }

    if ($null -ne $mainWindow) {
        try {
            $mainWindow.Visible = $false
        }
        catch {
            $cleanupOperationsSucceeded = $false
        }
    }

    if (-not (Release-ComObjectSafely $command)) {
        $cleanupOperationsSucceeded = $false
    }
    $command = $null

    if (-not (Release-ComObjectSafely $commandBars)) {
        $cleanupOperationsSucceeded = $false
    }
    $commandBars = $null

    if (-not (Release-ComObjectSafely $mainWindow)) {
        $cleanupOperationsSucceeded = $false
    }
    $mainWindow = $null

    if (-not (Release-ComObjectSafely $vbe)) {
        $cleanupOperationsSucceeded = $false
    }
    $vbe = $null

    if ($null -ne $presentation) {
        if (-not $compileSucceeded) {
            try {
                $presentation.Saved = -1
            }
            catch {
                $cleanupOperationsSucceeded = $false
            }
        }

        try {
            $presentation.Close()
        }
        catch {
            $cleanupOperationsSucceeded = $false
        }

        if (-not (Release-ComObjectSafely $presentation)) {
            $cleanupOperationsSucceeded = $false
        }
        $presentation = $null
    }

    if (-not (Release-ComObjectSafely $presentations)) {
        $cleanupOperationsSucceeded = $false
    }
    $presentations = $null

    if ($null -ne $ppt) {
        if ($comBoundToOwnedPowerPoint) {
            try {
                $ppt.Quit()
            }
            catch {
                $cleanupOperationsSucceeded = $false
            }
        }

        if (-not (Release-ComObjectSafely $ppt)) {
            $cleanupOperationsSucceeded = $false
        }
        $ppt = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($ownedPowerPoint) {
        $exitedNormally = Test-ExactOwnedPowerPointExited `
            -ProcessId $powerPointProcessId `
            -StartTimeUtcTicks $powerPointStartTimeUtcTicks `
            -Owned $ownedPowerPoint `
            -WaitMilliseconds 5000

        if ($exitedNormally) {
            $result.CleanupSucceeded = $true
        }
        else {
            $cleanupOperationsSucceeded = $false
            $result.CleanupSucceeded = Stop-ExactOwnedPowerPointProcess `
                -ProcessId $powerPointProcessId `
                -StartTimeUtcTicks $powerPointStartTimeUtcTicks `
                -Owned $ownedPowerPoint
        }
    }
    else {
        # A mapped pre-existing application is deliberately left running after
        # releasing this RCW. An unverified newly-created candidate is also never
        # terminated, but that conservative safety outcome is not clean recovery.
        $result.CleanupSucceeded = (
            -not $powerPointApplicationCreated -or $mappedProcessWasProtected
        )

        if (-not $result.CleanupSucceeded) {
            $result.Message += (
                " PowerPoint ownership was not proven, so no Quit or fallback " +
                "termination was attempted."
            )
        }
    }

    if ($null -ne $launcher) {
        try {
            if (-not $ownedPowerPoint -and -not $launcher.HasExited) {
                $launcher.Kill()
                [void]$launcher.WaitForExit(5000)
            }
            $launcher.Dispose()
        }
        catch {
            $cleanupOperationsSucceeded = $false
        }
        $launcher = $null
    }

    if (-not $cleanupOperationsSucceeded) {
        if ($result.Passed) {
            $result.Passed = $false
            $result.Status = "ComRecoveryFailed"
        }

        $result.Message += (
            " COM cleanup did not complete normally; " +
            "only the verified validator-owned process was eligible for fallback termination."
        )
    }

    try {
        Write-JsonFileAtomic -Path $resultPath -Value ([pscustomobject]$result)
    }
    catch {
        [Console]::Error.WriteLine("Could not write compile result: " + $_.Exception.Message)
    }
}

if ($result.Passed) {
    exit 0
}

exit 1

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$StateDirectory,
    [Parameter(Mandatory = $true)][int]$PowerPointProcessId,
    [Parameter(Mandatory = $true)][Int64]$PowerPointStartTimeUtcTicks,
    [ValidateRange(5, 3600)][int]$TimeoutSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public sealed class IguanaTexCompileDialogWindow
{
    public IntPtr Handle;
    public string Title;
    public string ClassName;
}

public static class IguanaTexCompileWatcherNative
{
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_COMMAND = 0x0111;
    private const uint IDOK = 1;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hWnd,
        out uint processId
    );

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

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

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr hWnd,
        uint message,
        UIntPtr wParam,
        IntPtr lParam,
        uint flags,
        uint timeout,
        out UIntPtr result
    );

    private static string ReadWindowText(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        StringBuilder text = new StringBuilder(Math.Max(length + 1, 2));
        GetWindowText(hWnd, text, text.Capacity);
        return text.ToString();
    }

    private static string ReadClassName(IntPtr hWnd)
    {
        StringBuilder className = new StringBuilder(256);
        GetClassName(hWnd, className, className.Capacity);
        return className.ToString();
    }

    public static IguanaTexCompileDialogWindow[] FindVisualBasicDialogs(
        uint expectedProcessId
    )
    {
        List<IguanaTexCompileDialogWindow> dialogs =
            new List<IguanaTexCompileDialogWindow>();

        EnumWindows(delegate(IntPtr hWnd, IntPtr ignored)
        {
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);

            if (processId != expectedProcessId || !IsWindowVisible(hWnd))
            {
                return true;
            }

            string title = ReadWindowText(hWnd);
            string className = ReadClassName(hWnd);

            // Exact PID plus a newly-created visible top-level window is the
            // safety boundary. #32770 also covers localized Office dialog titles.
            if (title.StartsWith(
                    "Microsoft Visual Basic",
                    StringComparison.OrdinalIgnoreCase) ||
                String.Equals(className, "#32770", StringComparison.Ordinal))
            {
                dialogs.Add(new IguanaTexCompileDialogWindow
                {
                    Handle = hWnd,
                    Title = title,
                    ClassName = className
                });
            }

            return true;
        }, IntPtr.Zero);

        return dialogs.ToArray();
    }

    public static bool DismissDialog(IntPtr hWnd)
    {
        if (!IsWindow(hWnd))
        {
            return true;
        }

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
}
'@

function Write-JsonFileAtomic {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $temporaryPath = $Path + "." + $PID + ".tmp"
    $json = $Value | ConvertTo-Json -Depth 6 -Compress
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($temporaryPath, $json, $encoding)
    Move-Item -LiteralPath $temporaryPath -Destination $Path -Force
}

function Write-MarkerFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, [string][DateTime]::UtcNow.Ticks, $encoding)
}

function Test-ExactPowerPointProcess {
    $process = Get-Process -Id $PowerPointProcessId -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return $false
    }

    try {
        return (
            $process.ProcessName -eq "POWERPNT" -and
            [Int64]$process.StartTime.ToUniversalTime().Ticks -eq
                $PowerPointStartTimeUtcTicks
        )
    }
    catch {
        return $false
    }
    finally {
        $process.Dispose()
    }
}

$readyPath = Join-Path $StateDirectory "watcher.ready"
$ownerPath = Join-Path $StateDirectory "watcher-owner.json"
$outcomePath = Join-Path $StateDirectory "watcher-outcome.json"
$stopPath = Join-Path $StateDirectory "watcher.stop"
$compileStartPath = Join-Path $StateDirectory "compile.start"
$executeReturnedPath = Join-Path $StateDirectory "execute.returned"
$deadlineUtc = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$outcomeWritten = $false
$watcherProcess = Get-Process -Id $PID

try {
    $watcherStartTicks = [Int64]$watcherProcess.StartTime.ToUniversalTime().Ticks
}
finally {
    $watcherProcess.Dispose()
}

try {
    Write-JsonFileAtomic -Path $ownerPath -Value ([pscustomobject][ordered]@{
        ProcessId             = [int]$PID
        StartTimeUtcTicks     = [string]$watcherStartTicks
    })

    if (-not (Test-ExactPowerPointProcess)) {
        throw "Validator-owned PowerPoint process identity was not valid."
    }

    $preexistingDialogHandles = New-Object 'System.Collections.Generic.HashSet[Int64]'

    foreach ($dialog in @(
        [IguanaTexCompileWatcherNative]::FindVisualBasicDialogs(
            [UInt32]$PowerPointProcessId
        )
    )) {
        [void]$preexistingDialogHandles.Add($dialog.Handle.ToInt64())
    }

    Write-MarkerFile $readyPath

    while (-not (Test-Path -LiteralPath $compileStartPath -PathType Leaf)) {
        if (Test-Path -LiteralPath $stopPath -PathType Leaf) {
            exit 0
        }

        if ([DateTime]::UtcNow -ge $deadlineUtc) {
            Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                Status          = "Timeout"
                DialogDetected  = $false
                Dismissed       = $false
                Message         = "Compile watcher timed out before execution began."
            })
            $outcomeWritten = $true
            exit 3
        }

        if (-not (Test-ExactPowerPointProcess)) {
            Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                Status          = "ProcessExited"
                DialogDetected  = $false
                Dismissed       = $false
                Message         = "Validator-owned PowerPoint exited before compile began."
            })
            $outcomeWritten = $true
            exit 2
        }

        Start-Sleep -Milliseconds 100
    }

    $executeReturnObservedUtc = $null

    while ([DateTime]::UtcNow -lt $deadlineUtc) {
        if (Test-Path -LiteralPath $stopPath -PathType Leaf) {
            exit 0
        }

        if (-not (Test-ExactPowerPointProcess)) {
            Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                Status          = "ProcessExited"
                DialogDetected  = $false
                Dismissed       = $false
                Message         = "Validator-owned PowerPoint exited during compile."
            })
            $outcomeWritten = $true
            exit 2
        }

        $newDialogs = @(
            [IguanaTexCompileWatcherNative]::FindVisualBasicDialogs(
                [UInt32]$PowerPointProcessId
            ) | Where-Object {
                -not $preexistingDialogHandles.Contains($_.Handle.ToInt64())
            }
        )

        if ($newDialogs.Count -gt 0) {
            $allDismissed = $true
            $descriptions = @()

            foreach ($dialog in $newDialogs) {
                $descriptions += (
                    "title='" + $dialog.Title + "', class='" +
                    $dialog.ClassName + "'"
                )

                if (-not [IguanaTexCompileWatcherNative]::DismissDialog(
                    $dialog.Handle
                )) {
                    $allDismissed = $false
                }
            }

            $status = "CompileErrorDialog"
            $message = (
                "Detected a PID-scoped Microsoft Visual Basic compile dialog (" +
                ($descriptions -join "; ") + ")."
            )

            if (-not $allDismissed) {
                $status = "CompileErrorDialogNotDismissed"
                $message += " The dialog could not be dismissed cleanly."
            }

            Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                Status          = $status
                DialogDetected  = $true
                Dismissed       = $allDismissed
                Message         = $message
            })
            $outcomeWritten = $true
            exit 2
        }

        if (Test-Path -LiteralPath $executeReturnedPath -PathType Leaf) {
            if ($null -eq $executeReturnObservedUtc) {
                $executeReturnObservedUtc = [DateTime]::UtcNow
            }
            elseif (([DateTime]::UtcNow - $executeReturnObservedUtc).TotalMilliseconds -ge
                1200) {
                Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                    Status          = "Clear"
                    DialogDetected  = $false
                    Dismissed       = $false
                    Message         = "No PID-scoped VBE compile dialog was detected."
                })
                $outcomeWritten = $true
                exit 0
            }
        }

        Start-Sleep -Milliseconds 100
    }

    Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
        Status          = "Timeout"
        DialogDetected  = $false
        Dismissed       = $false
        Message         = "Compile dialog watcher timed out."
    })
    $outcomeWritten = $true
    exit 3
}
catch {
    if (-not $outcomeWritten) {
        try {
            Write-JsonFileAtomic -Path $outcomePath -Value ([pscustomobject][ordered]@{
                Status          = "WatcherFailed"
                DialogDetected  = $false
                Dismissed       = $false
                Message         = $_.Exception.Message
            })
        }
        catch {
        }
    }

    [Console]::Error.WriteLine($_.Exception.Message)
    exit 2
}

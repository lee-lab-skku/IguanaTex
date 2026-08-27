#requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function ConvertTo-PowerShellSingleQuotedLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return "'" + $Value.Replace("'", "''") + "'"
}

function Write-CompileJsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object]$Value
    )

    $json = $Value | ConvertTo-Json -Depth 8 -Compress
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $json, $encoding)
}

function Read-CompileJsonFile {
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

function Get-PowerPointProcessSnapshot {
    $snapshot = @()

    foreach ($process in @(Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue)) {
        try {
            $snapshot += [pscustomobject][ordered]@{
                Id                    = [int]$process.Id
                StartTimeUtcTicks     = [string]$process.StartTime.ToUniversalTime().Ticks
            }
        }
        catch {
            # A process that cannot be identified precisely is still protected by PID.
            $snapshot += [pscustomobject][ordered]@{
                Id                    = [int]$process.Id
                StartTimeUtcTicks     = $null
            }
        }
        finally {
            $process.Dispose()
        }
    }

    return @($snapshot)
}

function Get-WindowsPowerShellPath {
    $path = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Windows PowerShell 5.1 was not found: $path"
    }

    return $path
}

function Start-StaPowerShellScript {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$InvocationArguments
    )

    $shellPath = Get-WindowsPowerShellPath
    $scriptLiteral = ConvertTo-PowerShellSingleQuotedLiteral $ScriptPath
    $command = (
        '$ProgressPreference = ''SilentlyContinue''; ' +
        "& $scriptLiteral $InvocationArguments"
    )
    $encodedCommand = [Convert]::ToBase64String(
        [Text.Encoding]::Unicode.GetBytes($command)
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $shellPath
    $startInfo.Arguments = (
        "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass " +
        "-STA -EncodedCommand $encodedCommand"
    )
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo

    if (-not $process.Start()) {
        $process.Dispose()
        throw "Failed to start compile validation worker."
    }

    # Drain redirected streams asynchronously so a verbose failure cannot deadlock.
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()

    return [pscustomobject]@{
        Process    = $process
        StdoutTask = $stdoutTask
        StderrTask = $stderrTask
    }
}

function Get-ExactPowerPointProcessState {
    param(
        [object]$OwnerRecord,
        [int[]]$ProtectedProcessIds
    )

    if ($null -eq $OwnerRecord -or -not [bool]$OwnerRecord.Owned) {
        return [pscustomobject]@{ Safe = $false; Running = $false; Process = $null }
    }

    try {
        $explicitLaunchVerified = [bool]$OwnerRecord.ExplicitLaunchVerified
        $activationTicks = [Int64]([string]$OwnerRecord.ActivationStartedUtcTicks)
        $recordedStartTicks = [Int64]([string]$OwnerRecord.StartTimeUtcTicks)

        if (-not $explicitLaunchVerified -or $recordedStartTicks -lt $activationTicks) {
            return [pscustomobject]@{ Safe = $false; Running = $false; Process = $null }
        }
    }
    catch {
        return [pscustomobject]@{ Safe = $false; Running = $false; Process = $null }
    }

    $processId = [int]$OwnerRecord.ProcessId

    if ($ProtectedProcessIds -contains $processId) {
        return [pscustomobject]@{ Safe = $false; Running = $true; Process = $null }
    }

    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return [pscustomobject]@{ Safe = $true; Running = $false; Process = $null }
    }

    try {
        $expectedTicks = [Int64]([string]$OwnerRecord.StartTimeUtcTicks)
        $actualTicks = [Int64]$process.StartTime.ToUniversalTime().Ticks

        if ($process.ProcessName -ne "POWERPNT" -or $actualTicks -ne $expectedTicks) {
            $process.Dispose()
            return [pscustomobject]@{ Safe = $false; Running = $true; Process = $null }
        }

        # Force acquisition of an OS process handle before returning it.
        [void]$process.Handle
        return [pscustomobject]@{ Safe = $true; Running = $true; Process = $process }
    }
    catch {
        $process.Dispose()
        return [pscustomobject]@{ Safe = $false; Running = $true; Process = $null }
    }
}

function Stop-ExactOwnedPowerPointProcess {
    param(
        [object]$OwnerRecord,
        [int[]]$ProtectedProcessIds
    )

    $state = Get-ExactPowerPointProcessState `
        -OwnerRecord $OwnerRecord `
        -ProtectedProcessIds $ProtectedProcessIds

    if (-not $state.Safe) {
        return $false
    }

    if (-not $state.Running) {
        return $true
    }

    try {
        $state.Process.Kill()
        [void]$state.Process.WaitForExit(5000)
        return $state.Process.HasExited
    }
    catch {
        return $false
    }
    finally {
        $state.Process.Dispose()
    }
}

function Stop-HelperProcessFromRecord {
    param([object]$ProcessRecord)

    if ($null -eq $ProcessRecord) {
        return
    }

    $process = Get-Process -Id ([int]$ProcessRecord.ProcessId) -ErrorAction SilentlyContinue

    if ($null -eq $process) {
        return
    }

    try {
        $expectedTicks = [Int64]([string]$ProcessRecord.StartTimeUtcTicks)
        $actualTicks = [Int64]$process.StartTime.ToUniversalTime().Ticks

        if ($actualTicks -eq $expectedTicks) {
            [void]$process.Handle

            # Allow an in-progress dialog dismissal and atomic outcome write to
            # finish before terminating the watcher helper.
            if (-not $process.WaitForExit(2000)) {
                $process.Kill()
                [void]$process.WaitForExit(3000)
            }
        }
    }
    catch {
    }
    finally {
        $process.Dispose()
    }
}

function New-CompileValidationResult {
    param(
        [bool]$Passed,
        [string]$Status,
        [object]$EnabledBefore,
        [object]$EnabledAfter,
        [bool]$DialogDetected,
        [bool]$CleanupSucceeded,
        [object]$OwnedPowerPointProcessId,
        [string]$Message
    )

    return [pscustomobject][ordered]@{
        Passed                       = $Passed
        Status                       = $Status
        EnabledBefore                = $EnabledBefore
        EnabledAfter                 = $EnabledAfter
        DialogDetected               = $DialogDetected
        CleanupSucceeded             = $CleanupSucceeded
        OwnedPowerPointProcessId      = $OwnedPowerPointProcessId
        Message                      = $Message
    }
}

function Invoke-VbeCompileValidation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$PresentationPath,

        [ValidateRange(5, 3600)]
        [int]$TimeoutSeconds = 120,

        [switch]$AllowAlreadyCompiled
    )

    if (-not [IO.Path]::IsPathRooted($PresentationPath)) {
        throw "PresentationPath must be absolute: $PresentationPath"
    }

    $presentationItem = Get-Item -LiteralPath $PresentationPath -ErrorAction Stop

    if ($presentationItem.PSIsContainer) {
        throw "PresentationPath must identify a file: $PresentationPath"
    }

    if ($presentationItem.Extension -ne ".pptm") {
        return New-CompileValidationResult `
            -Passed $false `
            -Status "UnsupportedPresentationType" `
            -EnabledBefore $null `
            -EnabledAfter $null `
            -DialogDetected $false `
            -CleanupSucceeded $true `
            -OwnedPowerPointProcessId $null `
            -Message (
                "Isolated VBE command validation currently supports .pptm only. " +
                "PPAM requires a separate AddIns.Add/Loaded validation workflow."
            )
    }

    $workerPath = Join-Path $PSScriptRoot "IguanaTex.Compile.Worker.ps1"
    $watcherPath = Join-Path $PSScriptRoot "IguanaTex.Compile.Watcher.ps1"

    foreach ($requiredPath in @($workerPath, $watcherPath)) {
        if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            throw "Compile validation helper was not found: $requiredPath"
        }
    }

    $stateDirectory = Join-Path (
        [IO.Path]::GetTempPath()
    ) ("IguanaTexCompile-" + [Guid]::NewGuid().ToString("N"))
    [void](New-Item -ItemType Directory -Path $stateDirectory)

    $snapshotPath = Join-Path $stateDirectory "preexisting-processes.json"
    $resultPath = Join-Path $stateDirectory "result.json"
    $ownerPath = Join-Path $stateDirectory "powerpoint-owner.json"
    $watcherOwnerPath = Join-Path $stateDirectory "watcher-owner.json"
    $watcherOutcomePath = Join-Path $stateDirectory "watcher-outcome.json"
    $protectedProcesses = @(Get-PowerPointProcessSnapshot)
    $protectedIds = @($protectedProcesses | ForEach-Object { [int]$_.Id })
    $workerBundle = $null
    $timedOut = $false
    $workerExitCode = $null
    $workerStdout = ""
    $workerStderr = ""
    $finalResult = $null

    try {
        Write-CompileJsonFile -Path $snapshotPath -Value ([pscustomobject][ordered]@{
            CapturedAtUtcTicks = [string][DateTime]::UtcNow.Ticks
            Processes          = @($protectedProcesses)
        })

        $arguments = @(
            "-PresentationPath " + (ConvertTo-PowerShellSingleQuotedLiteral $presentationItem.FullName)
            "-StateDirectory " + (ConvertTo-PowerShellSingleQuotedLiteral $stateDirectory)
            "-PreexistingProcessesPath " + (ConvertTo-PowerShellSingleQuotedLiteral $snapshotPath)
            "-WatcherPath " + (ConvertTo-PowerShellSingleQuotedLiteral $watcherPath)
            "-TimeoutSeconds $TimeoutSeconds"
        ) -join " "

        if ($AllowAlreadyCompiled) {
            $arguments += " -AllowAlreadyCompiled"
        }

        $workerBundle = Start-StaPowerShellScript `
            -ScriptPath $workerPath `
            -InvocationArguments $arguments
        # TimeoutSeconds bounds compile work inside the worker. Grant a fixed
        # grace period for COM cleanup and atomic result writing.
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds + 15)

        while (-not $workerBundle.Process.HasExited) {
            if ([DateTime]::UtcNow -ge $deadline) {
                $timedOut = $true
                break
            }

            Start-Sleep -Milliseconds 100
        }

        if ($timedOut) {
            $watcherOwner = Read-CompileJsonFile $watcherOwnerPath
            Stop-HelperProcessFromRecord $watcherOwner
            # Re-read only after the watcher has stopped so an atomic dialog
            # outcome written during dismissal cannot be missed.
            $watcherOutcome = Read-CompileJsonFile $watcherOutcomePath

            try {
                [void]$workerBundle.Process.Handle
                $workerBundle.Process.Kill()
                [void]$workerBundle.Process.WaitForExit(3000)
            }
            catch {
            }

            try {
                $workerStdout = $workerBundle.StdoutTask.GetAwaiter().GetResult()
                $workerStderr = $workerBundle.StderrTask.GetAwaiter().GetResult()
            }
            catch {
            }

            $owner = Read-CompileJsonFile $ownerPath
            $cleanupSucceeded = Stop-ExactOwnedPowerPointProcess `
                -OwnerRecord $owner `
                -ProtectedProcessIds $protectedIds
            $ownedId = $null

            if ($null -ne $owner -and [bool]$owner.Owned) {
                $ownedId = [int]$owner.ProcessId
            }

            $timeoutStatus = "Timeout"
            $dialogDetected = $false
            $timeoutMessage = "VBE compile validation exceeded $TimeoutSeconds seconds."

            if ($null -ne $watcherOutcome -and
                [bool]$watcherOutcome.DialogDetected) {
                $timeoutStatus = "CompileErrorDialog"
                $dialogDetected = $true
                $timeoutMessage = (
                    [string]$watcherOutcome.Message +
                    " The worker did not complete within its cleanup grace period."
                )
            }

            if (-not [string]::IsNullOrWhiteSpace($workerStderr)) {
                $timeoutMessage += " Worker stderr: " + $workerStderr.Trim()
            }

            if (-not [string]::IsNullOrWhiteSpace($workerStdout)) {
                $timeoutMessage += " Worker stdout: " + $workerStdout.Trim()
            }

            $finalResult = New-CompileValidationResult `
                -Passed $false `
                -Status $timeoutStatus `
                -EnabledBefore $null `
                -EnabledAfter $null `
                -DialogDetected $dialogDetected `
                -CleanupSucceeded $cleanupSucceeded `
                -OwnedPowerPointProcessId $ownedId `
                -Message $timeoutMessage
        }
        else {
            $workerBundle.Process.WaitForExit()
            $workerExitCode = $workerBundle.Process.ExitCode
            $workerStdout = $workerBundle.StdoutTask.GetAwaiter().GetResult()
            $workerStderr = $workerBundle.StderrTask.GetAwaiter().GetResult()
            $workerResult = Read-CompileJsonFile $resultPath

            if ($null -eq $workerResult) {
                $message = "Compile worker exited without a valid result (exit $workerExitCode)."
                $watcherOutcome = Read-CompileJsonFile $watcherOutcomePath
                $status = "WorkerFailed"
                $dialogDetected = $false

                if ($null -ne $watcherOutcome -and
                    [bool]$watcherOutcome.DialogDetected) {
                    $status = "CompileErrorDialog"
                    $dialogDetected = $true
                    $message = (
                        [string]$watcherOutcome.Message +
                        " The worker exited before writing its final result."
                    )
                }

                if (-not [string]::IsNullOrWhiteSpace($workerStderr)) {
                    $message += " " + $workerStderr.Trim()
                }
                elseif (-not [string]::IsNullOrWhiteSpace($workerStdout)) {
                    $message += " " + $workerStdout.Trim()
                }

                $owner = Read-CompileJsonFile $ownerPath
                $cleanupSucceeded = Stop-ExactOwnedPowerPointProcess `
                    -OwnerRecord $owner `
                    -ProtectedProcessIds $protectedIds
                $ownedId = $null

                if ($null -ne $owner -and [bool]$owner.Owned) {
                    $ownedId = [int]$owner.ProcessId
                }

                $finalResult = New-CompileValidationResult `
                    -Passed $false `
                    -Status $status `
                    -EnabledBefore $null `
                    -EnabledAfter $null `
                    -DialogDetected $dialogDetected `
                    -CleanupSucceeded $cleanupSucceeded `
                    -OwnedPowerPointProcessId $ownedId `
                    -Message $message
            }
            else {
                $finalResult = New-CompileValidationResult `
                    -Passed ([bool]$workerResult.Passed) `
                    -Status ([string]$workerResult.Status) `
                    -EnabledBefore $workerResult.EnabledBefore `
                    -EnabledAfter $workerResult.EnabledAfter `
                    -DialogDetected ([bool]$workerResult.DialogDetected) `
                    -CleanupSucceeded ([bool]$workerResult.CleanupSucceeded) `
                    -OwnedPowerPointProcessId $workerResult.OwnedPowerPointProcessId `
                    -Message ([string]$workerResult.Message)

                if (-not $finalResult.Passed) {
                    if (-not [string]::IsNullOrWhiteSpace($workerStderr)) {
                        $finalResult.Message += " Worker stderr: " + $workerStderr.Trim()
                    }

                    if (-not [string]::IsNullOrWhiteSpace($workerStdout)) {
                        $finalResult.Message += " Worker stdout: " + $workerStdout.Trim()
                    }
                }

                if (($finalResult.Passed -and $workerExitCode -ne 0) -or
                    (-not $finalResult.Passed -and $workerExitCode -eq 0)) {
                    $finalResult.Passed = $false
                    $finalResult.Status = "WorkerProtocolError"
                    $finalResult.Message = (
                        "Worker exit code $workerExitCode did not match its result."
                    )
                }

                $owner = Read-CompileJsonFile $ownerPath
                $ownerState = Get-ExactPowerPointProcessState `
                    -OwnerRecord $owner `
                    -ProtectedProcessIds $protectedIds

                if ($ownerState.Safe -and $ownerState.Running) {
                    $ownerState.Process.Dispose()
                    $stopped = Stop-ExactOwnedPowerPointProcess `
                        -OwnerRecord $owner `
                        -ProtectedProcessIds $protectedIds
                    $finalResult.CleanupSucceeded = $stopped

                    if ($finalResult.Passed) {
                        $finalResult.Passed = $false
                        $finalResult.Status = "ComRecoveryFailed"
                    }

                    $finalResult.Message += (
                        " Validator-owned PowerPoint remained after worker exit; " +
                        "exact-PID fallback termination attempted."
                    )
                }
                elseif ($ownerState.Safe -and -not $ownerState.Running) {
                    # The worker released its process normally.
                }
                elseif ($null -ne $owner -and [bool]$owner.Owned) {
                    $finalResult.Passed = $false
                    $finalResult.CleanupSucceeded = $false
                    $finalResult.Status = "OwnershipRecoveryFailed"
                    $finalResult.Message += (
                        " PowerPoint ownership could not be revalidated; no process was killed."
                    )
                }
            }
        }
    }
    catch {
        $owner = Read-CompileJsonFile $ownerPath
        $cleanupSucceeded = Stop-ExactOwnedPowerPointProcess `
            -OwnerRecord $owner `
            -ProtectedProcessIds $protectedIds
        $ownedId = $null

        if ($null -ne $owner -and [bool]$owner.Owned) {
            $ownedId = [int]$owner.ProcessId
        }

        $finalResult = New-CompileValidationResult `
            -Passed $false `
            -Status "ControllerFailed" `
            -EnabledBefore $null `
            -EnabledAfter $null `
            -DialogDetected $false `
            -CleanupSucceeded $cleanupSucceeded `
            -OwnedPowerPointProcessId $ownedId `
            -Message $_.Exception.Message
    }
    finally {
        if ($null -ne $workerBundle) {
            if (-not $workerBundle.Process.HasExited) {
                try {
                    $workerBundle.Process.Kill()
                    [void]$workerBundle.Process.WaitForExit(3000)
                }
                catch {
                }
            }

            $workerBundle.Process.Dispose()
        }

        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        $resolvedState = [IO.Path]::GetFullPath($stateDirectory)

        if ($resolvedState.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
            (Split-Path -Leaf $resolvedState).StartsWith("IguanaTexCompile-")) {
            Remove-Item -LiteralPath $resolvedState -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    return $finalResult
}

Export-ModuleMember -Function Invoke-VbeCompileValidation

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$SkipWatch,

    [switch]$KeepArtifacts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$modulePath = Join-Path $PSScriptRoot "lib\IguanaTex.UserForms.psm1"
$artifactsRoot = Join-Path $projectRoot ".build\userforms-verify"
$runRoot = Join-Path $artifactsRoot ("run-" + [DateTime]::UtcNow.ToString("yyyyMMddTHHmmssZ") + "-" + [Guid]::NewGuid().ToString("N"))
$sourceStage = Join-Path $runRoot "canonical-create"
$latestSummaryPath = Join-Path $artifactsRoot "summary.json"
$expectedExecutablePath = [IO.Path]::GetFullPath((Join-Path $projectRoot ".build\frxedit\frxedit.exe"))

function Write-JsonFile([string]$Path, [object]$Value, [int]$Depth = 30) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $parent = Split-Path -Parent $fullPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    [IO.File]::WriteAllText(
        $fullPath,
        ($Value | ConvertTo-Json -Depth $Depth),
        [Text.UTF8Encoding]::new($false))
}

function Get-FileUriReferences([object]$Value) {
    if ($null -eq $Value) {
        return
    }
    if ($Value -is [string]) {
        if ($Value.StartsWith("file://", [StringComparison]::OrdinalIgnoreCase)) {
            Write-Output $Value
        }
        return
    }
    if ($Value -is [Management.Automation.PSCustomObject]) {
        foreach ($property in $Value.PSObject.Properties) {
            Get-FileUriReferences $property.Value
        }
        return
    }
    if ($Value -is [Collections.IDictionary]) {
        foreach ($entryValue in $Value.Values) {
            Get-FileUriReferences $entryValue
        }
        return
    }
    if ($Value -is [Collections.IEnumerable]) {
        foreach ($item in $Value) {
            Get-FileUriReferences $item
        }
    }
}

function Assert-DocumentAssetReferences(
    [string]$DocumentPath,
    [string]$AllowedRoot,
    [bool]$RequireAtLeastOne) {
    $document = Get-Content -Raw -LiteralPath $DocumentPath | ConvertFrom-Json
    $references = @(Get-FileUriReferences $document)
    if ($RequireAtLeastOne -and $references.Count -eq 0) {
        throw "Expected at least one relative file:// asset reference: $DocumentPath"
    }

    $documentDirectory = [IO.Path]::GetFullPath((Split-Path -Parent $DocumentPath))
    $allowedDirectory = [IO.Path]::GetFullPath($AllowedRoot)
    $allowedPrefix = $allowedDirectory.TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    foreach ($reference in $references) {
        $referencePath = [Uri]::UnescapeDataString($reference.Substring("file://".Length))
        if ([string]::IsNullOrWhiteSpace($referencePath) -or
            $reference.StartsWith("file:///", [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::IsPathRooted($referencePath)) {
            throw "Asset URI must be relative to its JSON document: $reference"
        }

        $platformPath = $referencePath.Replace('/', [IO.Path]::DirectorySeparatorChar)
        $resolved = [IO.Path]::GetFullPath((Join-Path $documentDirectory $platformPath))
        if (-not $resolved.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Asset URI escapes its allowed directory '$allowedDirectory': $reference"
        }
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "Asset URI does not resolve to a file beside its JSON document: $reference"
        }
    }
    return $references
}

function Assert-RebuildReport([string]$ReportPath) {
    if (-not (Test-Path -LiteralPath $ReportPath -PathType Leaf)) {
        throw "FrxEdit did not write its rebuild report: $ReportPath"
    }
    $report = Get-Content -Raw -LiteralPath $ReportPath | ConvertFrom-Json
    $semanticMatchProperty = $report.PSObject.Properties["semanticMatch"]
    if ($null -eq $semanticMatchProperty -or -not [bool]$semanticMatchProperty.Value) {
        throw "FrxEdit's internal rebuild comparison did not match: $ReportPath"
    }
}

function Assert-SemanticReport([object]$Report, [string]$Description) {
    if ($null -eq $Report) {
        throw "Semantic comparator returned no report for $Description."
    }
    $semanticEqualProperty = $Report.PSObject.Properties["semanticEqual"]
    if ($null -eq $semanticEqualProperty -or -not [bool]$semanticEqualProperty.Value) {
        throw "Semantic comparator did not confirm equality for $Description."
    }
}

function Get-VbaSourceText([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "VBA source not found: $Path"
    }

    $encoding = [Text.Encoding]::GetEncoding(1252)
    return $encoding.GetString([IO.File]::ReadAllBytes($Path))
}

function Assert-VbaFilesEqual(
    [string]$ExpectedPath,
    [string]$ActualPath,
    [string]$Description) {
    $expected = Get-VbaSourceText $ExpectedPath
    $actual = Get-VbaSourceText $ActualPath
    if ($actual -cne $expected) {
        throw ("UserForm VBA sidecar mismatch for {0}: {1}" -f
            $Description,
            $ActualPath)
    }
}

function Assert-FormVbaEqual(
    [string]$FormPath,
    [string]$ExpectedVbaPath,
    [string]$Description) {
    $formText = Get-VbaSourceText $FormPath
    $attributeBlock = [regex]::Match(
        $formText,
        '^(Attribute\s+VB_[A-Za-z0-9_]+\s*=\s*.*?\r?\n)+',
        [Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $attributeBlock.Success) {
        throw ("Generated UserForm has no VBA attribute boundary for {0}: {1}" -f
            $Description,
            $FormPath)
    }

    $actualVba = $formText.Substring(
        $attributeBlock.Index + $attributeBlock.Length)
    $expectedVba = Get-VbaSourceText $ExpectedVbaPath
    if ($actualVba -cne $expectedVba) {
        throw ("Generated UserForm VBA mismatch for {0}: {1}" -f
            $Description,
            $FormPath)
    }
}

function ConvertTo-ProcessArgument([string]$Value) {
    if ($Value.Length -eq 0) {
        return '""'
    }
    if ($Value -notmatch '[\s"]') {
        return $Value
    }

    # Apply the Windows CommandLineToArgvW quoting rules used by ProcessStartInfo.Arguments.
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq [char]92) {
            $backslashCount++
            continue
        }
        if ($character -eq [char]34) {
            [void]$builder.Append([string]::new([char]92, (($backslashCount * 2) + 1)))
            [void]$builder.Append('"')
            $backslashCount = 0
            continue
        }
        if ($backslashCount -gt 0) {
            [void]$builder.Append([string]::new([char]92, $backslashCount))
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }
    if ($backslashCount -gt 0) {
        [void]$builder.Append([string]::new([char]92, ($backslashCount * 2)))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function New-WatchStartInfo(
    [string]$ExecutablePath,
    [string[]]$Arguments,
    [string]$WorkingDirectory) {
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = [IO.Path]::GetFullPath($ExecutablePath)
    $startInfo.WorkingDirectory = [IO.Path]::GetFullPath($WorkingDirectory)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    if ($startInfo.PSObject.Properties.Name -contains "ArgumentList") {
        foreach ($argument in $Arguments) {
            [void]$startInfo.ArgumentList.Add($argument)
        }
    }
    else {
        $startInfo.Arguments = (@($Arguments | ForEach-Object { ConvertTo-ProcessArgument $_ }) -join " ")
    }
    return $startInfo
}

function Invoke-BoundedWatch(
    [string]$ExecutablePath,
    [string]$InputForm,
    [string]$PatchPath,
    [string]$OutputForm,
    [string]$WorkingDirectory,
    [string]$ArtifactDirectory) {
    New-Item -ItemType Directory -Force -Path $WorkingDirectory, $ArtifactDirectory | Out-Null

    $stdoutPath = Join-Path $ArtifactDirectory "watch.stdout.log"
    $stderrPath = Join-Path $ArtifactDirectory "watch.stderr.log"
    $statePath = Join-Path $ArtifactDirectory "watch-state.json"
    $outputFrx = [IO.Path]::ChangeExtension($OutputForm, ".frx")
    $arguments = @(
        "watch", $InputForm, $PatchPath,
        "--out", $OutputForm,
        "--mode", "strict",
        "--stream-mode", "full-patch")

    $process = $null
    $processStarted = $false
    $stdoutTask = $null
    $stderrTask = $null
    $stdout = ""
    $stderr = ""
    $ownedProcessId = $null
    $initialOutputReady = $false
    $runningAtTrigger = $false
    $triggered = $false
    $regenerated = $false
    $survivedTrigger = $false
    $terminated = $false
    $ownedStillRunning = $false

    try {
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = New-WatchStartInfo $ExecutablePath $arguments $WorkingDirectory
        [void]$process.Start()
        $processStarted = $true
        $ownedProcessId = $process.Id
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()

        $initialDeadline = [DateTime]::UtcNow.AddSeconds(30)
        while ([DateTime]::UtcNow -lt $initialDeadline) {
            $process.Refresh()
            if ($process.HasExited) {
                break
            }
            if ((Test-Path -LiteralPath $OutputForm -PathType Leaf) -and
                (Test-Path -LiteralPath $outputFrx -PathType Leaf)) {
                $initialOutputReady = $true
                break
            }
            Start-Sleep -Milliseconds 100
        }

        $process.Refresh()
        $runningAtTrigger = -not $process.HasExited
        if ($initialOutputReady -and $runningAtTrigger) {
            # Let the initial atomic output pair settle before taking the
            # baseline hashes used to prove a subsequent watch regeneration.
            Start-Sleep -Milliseconds 250
            $beforeFrm = Get-Item -LiteralPath $OutputForm
            $beforeFrx = Get-Item -LiteralPath $outputFrx
            $beforeFrmWrite = $beforeFrm.LastWriteTimeUtc
            $beforeFrxWrite = $beforeFrx.LastWriteTimeUtc
            $beforeFrmHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $OutputForm).Hash
            $beforeFrxHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFrx).Hash

            $patchBeforeWrite = (Get-Item -LiteralPath $PatchPath).LastWriteTimeUtc
            $patchTriggerWrite = [DateTime]::UtcNow
            if ($patchTriggerWrite -le $patchBeforeWrite) {
                $patchTriggerWrite = $patchBeforeWrite.AddSeconds(1)
            }
            [IO.File]::SetLastWriteTimeUtc($PatchPath, $patchTriggerWrite)
            $triggered = (Get-Item -LiteralPath $PatchPath).LastWriteTimeUtc -gt $patchBeforeWrite

            $regenerationDeadline = [DateTime]::UtcNow.AddSeconds(30)
            while ($triggered -and [DateTime]::UtcNow -lt $regenerationDeadline) {
                $process.Refresh()
                if ($process.HasExited) {
                    break
                }
                if ((Test-Path -LiteralPath $OutputForm -PathType Leaf) -and
                    (Test-Path -LiteralPath $outputFrx -PathType Leaf)) {
                    try {
                        $afterFrm = Get-Item -LiteralPath $OutputForm
                        $afterFrx = Get-Item -LiteralPath $outputFrx
                        $afterFrmHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $OutputForm).Hash
                        $afterFrxHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $outputFrx).Hash
                        $regenerated =
                            $afterFrm.LastWriteTimeUtc -gt $beforeFrmWrite -or
                            $afterFrx.LastWriteTimeUtc -gt $beforeFrxWrite -or
                            $afterFrmHash -cne $beforeFrmHash -or
                            $afterFrxHash -cne $beforeFrxHash
                    }
                    catch [IO.IOException] {
                        # File replacement can make the output pair briefly unavailable.
                        $regenerated = $false
                    }
                    if ($regenerated) {
                        break
                    }
                }
                Start-Sleep -Milliseconds 100
            }

            if ($regenerated) {
                Start-Sleep -Milliseconds 250
                $process.Refresh()
                $survivedTrigger = -not $process.HasExited
            }
        }
    }
    catch {
        $stderr += ($_ | Out-String)
    }
    finally {
        if ($null -ne $process) {
            if ($processStarted) {
                try {
                    $process.Refresh()
                    if (-not $process.HasExited) {
                        $process.Kill()
                        $terminated = $process.WaitForExit(10000)
                    }
                }
                catch {
                    $stderr += ($_ | Out-String)
                }

                try {
                    $process.Refresh()
                    if ($process.HasExited) {
                        if ($null -ne $stdoutTask) {
                            $stdout += $stdoutTask.Result
                        }
                        if ($null -ne $stderrTask) {
                            $stderr += $stderrTask.Result
                        }
                    }
                }
                catch {
                    $stderr += ($_ | Out-String)
                }
            }
            $process.Dispose()
        }
        [IO.File]::WriteAllText($stdoutPath, $stdout, [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText($stderrPath, $stderr, [Text.UTF8Encoding]::new($false))
    }

    if ($null -ne $ownedProcessId) {
        $ownedStillRunning = $null -ne (Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue)
    }
    $workingPrefix = [IO.Path]::GetFullPath($WorkingDirectory).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $patchOutsideWorkingDirectory = -not [IO.Path]::GetFullPath($PatchPath).StartsWith(
        $workingPrefix,
        [StringComparison]::OrdinalIgnoreCase)
    $passed =
        $initialOutputReady -and
        $runningAtTrigger -and
        $triggered -and
        $regenerated -and
        $survivedTrigger -and
        $terminated -and
        $patchOutsideWorkingDirectory -and
        -not $ownedStillRunning
    $state = [ordered]@{
        schemaVersion = 1
        inputForm = [IO.Path]::GetFullPath($InputForm)
        patch = [IO.Path]::GetFullPath($PatchPath)
        outputForm = [IO.Path]::GetFullPath($OutputForm)
        workingDirectory = [IO.Path]::GetFullPath($WorkingDirectory)
        patchDocumentOutsideWorkingDirectory = $patchOutsideWorkingDirectory
        ownedProcessId = $ownedProcessId
        initialOutputReady = $initialOutputReady
        runningAtTrigger = $runningAtTrigger
        noOpMtimeTrigger = $triggered
        outputRegeneratedAfterBaseline = $regenerated
        processSurvivedTrigger = $survivedTrigger
        boundedProcessTerminated = $terminated
        ownedProcessStillRunning = $ownedStillRunning
        passed = $passed
    }
    Write-JsonFile $statePath $state
    if (-not $passed) {
        throw "Bounded watch reconstruction failed; see $statePath"
    }
}

function Compare-Form(
    [string]$ExecutablePath,
    [string]$OriginalForm,
    [string]$CandidateForm,
    [string]$ExpectedVbaPath,
    [string]$ArtifactDirectory,
    [string]$Description) {
    Assert-FormVbaEqual `
        -FormPath $CandidateForm `
        -ExpectedVbaPath $ExpectedVbaPath `
        -Description $Description
    $report = Compare-IguanaTexUserFormSemantics `
        -ProjectRoot $projectRoot `
        -FrxEditPath $ExecutablePath `
        -OriginalFormPath $OriginalForm `
        -CandidateFormPath $CandidateForm `
        -ArtifactsDirectory $ArtifactDirectory `
        -ReportPath (Join-Path $ArtifactDirectory "canonical-comparison.json")
    Assert-SemanticReport $report $Description
}

function Test-UserForm([object]$Form, [string]$ExecutablePath) {
    $formName = [string]$Form.Name
    $caseRoot = Join-Path $runRoot $formName
    $sourceRoot = Join-Path $caseRoot "source-copy"
    $exportRoot = Join-Path $caseRoot "exported"
    $unrelatedWorkingDirectory = Join-Path $caseRoot "unrelated-process-cwd"
    New-Item -ItemType Directory -Force -Path $sourceRoot, $exportRoot, $unrelatedWorkingDirectory | Out-Null

    $stagedForm = Join-Path $sourceStage ($formName + ".frm")
    $stagedFrx = Join-Path $sourceStage ($formName + ".frx")
    if (-not (Test-Path -LiteralPath $stagedForm -PathType Leaf) -or
        -not (Test-Path -LiteralPath $stagedFrx -PathType Leaf)) {
        throw "Canonical staging is missing the generated pair for '$formName'."
    }

    $sourceForm = Join-Path $sourceRoot ($formName + ".frm")
    $sourceFrx = Join-Path $sourceRoot ($formName + ".frx")
    Copy-Item -LiteralPath $stagedForm -Destination $sourceForm
    Copy-Item -LiteralPath $stagedFrx -Destination $sourceFrx
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "validate", $sourceForm, "--mode", "strict")

    $canonicalAssetReferences = @(Assert-DocumentAssetReferences `
        -DocumentPath $Form.TemplatePath `
        -AllowedRoot $Form.AssetRoot `
        -RequireAtLeastOne:$false)

    # Recreate directly from canonical JSON while the process CWD is unrelated to
    # the JSON. This verifies strict create and JSON-relative asset resolution.
    $directCreateRoot = Join-Path $caseRoot "direct-canonical-create"
    New-Item -ItemType Directory -Force -Path $directCreateRoot | Out-Null
    $directCreateForm = Join-Path $directCreateRoot ($formName + ".frm")
    Invoke-FrxEditCommand `
        -ExecutablePath $ExecutablePath `
        -Arguments @("create", $directCreateForm, "--name", $formName, "--patch", $Form.TemplatePath) `
        -WorkingDirectory $unrelatedWorkingDirectory
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "validate", $directCreateForm, "--mode", "strict")
    Compare-Form `
        -ExecutablePath $ExecutablePath `
        -OriginalForm $sourceForm `
        -CandidateForm $directCreateForm `
        -ExpectedVbaPath $Form.VbaPath `
        -ArtifactDirectory (Join-Path $directCreateRoot "comparison") `
        -Description "$formName direct canonical create"

    $noOpRoot = Join-Path $caseRoot "no-op"
    New-Item -ItemType Directory -Force -Path $noOpRoot | Out-Null
    $noOpForm = Join-Path $noOpRoot ($formName + ".frm")
    $noOpReport = Join-Path $noOpRoot "rebuild-report.json"
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "build", $sourceForm,
        "--out", $noOpForm,
        "--mode", "strict",
        "--stream-mode", "full-patch",
        "--report-out", $noOpReport)
    Assert-RebuildReport $noOpReport
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "validate", $noOpForm, "--mode", "strict")
    Compare-Form `
        -ExecutablePath $ExecutablePath `
        -OriginalForm $sourceForm `
        -CandidateForm $noOpForm `
        -ExpectedVbaPath $Form.VbaPath `
        -ArtifactDirectory (Join-Path $noOpRoot "comparison") `
        -Description "$formName no-op rebuild"

    $patchPath = Join-Path $exportRoot ($formName + ".patch.json")
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "inspect", $sourceForm,
        "--mode", "strict",
        "--as-patch",
        "--extract-images",
        "--out", $patchPath)
    $patchAssetReferences = @(Assert-DocumentAssetReferences `
        -DocumentPath $patchPath `
        -AllowedRoot $exportRoot `
        -RequireAtLeastOne:($canonicalAssetReferences.Count -gt 0))

    foreach ($patchStyle in @("positional", "option")) {
        $patchRoot = Join-Path $caseRoot ("patch-reapply-" + $patchStyle)
        New-Item -ItemType Directory -Force -Path $patchRoot | Out-Null
        $patchForm = Join-Path $patchRoot ($formName + ".frm")
        $patchReport = Join-Path $patchRoot "rebuild-report.json"
        $patchArguments = if ($patchStyle -ceq "positional") {
            @(
                "build", $sourceForm, $patchPath,
                "--out", $patchForm,
                "--mode", "strict",
                "--stream-mode", "full-patch",
                "--report-out", $patchReport)
        }
        else {
            @(
                "build", $sourceForm,
                "--patch", $patchPath,
                "--out", $patchForm,
                "--mode", "strict",
                "--stream-mode", "full-patch",
                "--report-out", $patchReport)
        }
        Invoke-FrxEditCommand `
            -ExecutablePath $ExecutablePath `
            -Arguments $patchArguments `
            -WorkingDirectory $unrelatedWorkingDirectory
        Assert-RebuildReport $patchReport
        Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
            "validate", $patchForm, "--mode", "strict")
        Compare-Form `
            -ExecutablePath $ExecutablePath `
            -OriginalForm $sourceForm `
            -CandidateForm $patchForm `
            -ExpectedVbaPath $Form.VbaPath `
            -ArtifactDirectory (Join-Path $patchRoot "comparison") `
            -Description "$formName $patchStyle exported-patch reapply"
    }

    if (-not $SkipWatch) {
        $watchRoot = Join-Path $caseRoot "watch-reapply"
        $watchInputRoot = Join-Path $watchRoot "input"
        $watchOutputRoot = Join-Path $watchRoot "output"
        $watchWorkingDirectory = Join-Path $watchRoot "unrelated-process-cwd"
        New-Item -ItemType Directory -Force -Path $watchInputRoot, $watchOutputRoot, $watchWorkingDirectory | Out-Null
        $watchInputForm = Join-Path $watchInputRoot ($formName + ".frm")
        $watchInputFrx = Join-Path $watchInputRoot ($formName + ".frx")
        $watchOutputForm = Join-Path $watchOutputRoot ($formName + ".frm")
        Copy-Item -LiteralPath $sourceForm -Destination $watchInputForm
        Copy-Item -LiteralPath $sourceFrx -Destination $watchInputFrx

        Invoke-BoundedWatch `
            -ExecutablePath $ExecutablePath `
            -InputForm $watchInputForm `
            -PatchPath $patchPath `
            -OutputForm $watchOutputForm `
            -WorkingDirectory $watchWorkingDirectory `
            -ArtifactDirectory $watchRoot
        Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
            "validate", $watchOutputForm, "--mode", "strict")
        Compare-Form `
            -ExecutablePath $ExecutablePath `
            -OriginalForm $sourceForm `
            -CandidateForm $watchOutputForm `
            -ExpectedVbaPath $Form.VbaPath `
            -ArtifactDirectory (Join-Path $watchRoot "comparison") `
            -Description "$formName bounded watch"
    }

    $templatePath = Join-Path $exportRoot ($formName + ".template.json")
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "inspect", $sourceForm,
        "--mode", "strict",
        "--as-template",
        "--extract-images",
        "--out", $templatePath)
    $templateAssetReferences = @(Assert-DocumentAssetReferences `
        -DocumentPath $templatePath `
        -AllowedRoot $exportRoot `
        -RequireAtLeastOne:($canonicalAssetReferences.Count -gt 0))

    # `inspect --as-template` must preserve the native form's separately
    # exported VBA before that sidecar is allowed to drive recreation.
    $templateVbaPath = [IO.Path]::ChangeExtension($templatePath, ".vba")
    Assert-VbaFilesEqual `
        -ExpectedPath $Form.VbaPath `
        -ActualPath $templateVbaPath `
        -Description "$formName exported template"

    $templateRoot = Join-Path $caseRoot "template-recreate"
    New-Item -ItemType Directory -Force -Path $templateRoot | Out-Null
    $templateForm = Join-Path $templateRoot ($formName + ".frm")
    Invoke-FrxEditCommand `
        -ExecutablePath $ExecutablePath `
        -Arguments @("create", $templateForm, "--name", $formName, "--patch", $templatePath) `
        -WorkingDirectory $unrelatedWorkingDirectory
    Invoke-FrxEditCommand -ExecutablePath $ExecutablePath -Arguments @(
        "validate", $templateForm, "--mode", "strict")
    Compare-Form `
        -ExecutablePath $ExecutablePath `
        -OriginalForm $sourceForm `
        -CandidateForm $templateForm `
        -ExpectedVbaPath $Form.VbaPath `
        -ArtifactDirectory (Join-Path $templateRoot "comparison") `
        -Description "$formName exported-template recreate"

    $workingPrefix = [IO.Path]::GetFullPath($unrelatedWorkingDirectory).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $templateOutsideWorkingDirectory = -not [IO.Path]::GetFullPath($Form.TemplatePath).StartsWith(
        $workingPrefix,
        [StringComparison]::OrdinalIgnoreCase)
    if (-not $templateOutsideWorkingDirectory) {
        throw "Canonical template test CWD is not unrelated for '$formName'."
    }

    $assetContract = [ordered]@{
        schemaVersion = 1
        canonicalTemplate = [IO.Path]::GetFullPath($Form.TemplatePath)
        processWorkingDirectory = [IO.Path]::GetFullPath($unrelatedWorkingDirectory)
        templateOutsideWorkingDirectory = $templateOutsideWorkingDirectory
        canonicalAssetReferences = $canonicalAssetReferences
        patchAssetReferences = $patchAssetReferences
        templateAssetReferences = $templateAssetReferences
    }
    Write-JsonFile (Join-Path $caseRoot "asset-path-contract.json") $assetContract

    return [pscustomobject][ordered]@{
        form = $formName
        originalStrict = $true
        canonicalCreateSemantic = $true
        noOpSemantic = $true
        positionalPatchSemantic = $true
        optionPatchSemantic = $true
        watchSemantic = $(if ($SkipWatch) { $null } else { $true })
        templateSemantic = $true
        canonicalAssetReferenceCount = $canonicalAssetReferences.Count
        exportedPatchAssetReferenceCount = $patchAssetReferences.Count
        exportedTemplateAssetReferenceCount = $templateAssetReferences.Count
    }
}

if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
    throw "Missing shared UserForm module: $modulePath"
}

Import-Module $modulePath -Force -DisableNameChecking
New-Item -ItemType Directory -Force -Path $runRoot | Out-Null

$completed = $false
$caughtError = $null
$integrityError = $null
$canonicalBefore = $null
$summary = $null
try {
    $manifest = Assert-IguanaTexUserFormManifest -ProjectRoot $projectRoot
    if (@($manifest.Forms).Count -ne 9) {
        throw "The canonical UserForm manifest must list exactly 9 forms; found $(@($manifest.Forms).Count)."
    }

    $canonicalBefore = Get-IguanaTexCanonicalInputHash -ProjectRoot $projectRoot
    Write-JsonFile (Join-Path $runRoot "canonical-hash.before.json") $canonicalBefore

    # Pinned mode verifies initialization, the superproject gitlink, submodule
    # HEAD, and a clean checkout before publishing the exact codec revision.
    $provenance = Invoke-FrxEditBuild -ProjectRoot $projectRoot -Mode Pinned
    $executablePath = [IO.Path]::GetFullPath([string]$provenance.ExecutablePath)
    if (-not $executablePath.Equals($expectedExecutablePath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Pinned FrxEdit was published to an unexpected path: $executablePath"
    }
    if (-not (Test-Path -LiteralPath $executablePath -PathType Leaf)) {
        throw "Pinned FrxEdit executable was not published: $executablePath"
    }
    Write-JsonFile (Join-Path $runRoot "frxedit-provenance.json") $provenance

    $staging = New-IguanaTexUserFormStaging `
        -ProjectRoot $projectRoot `
        -FrxEditPath $executablePath `
        -OutputDirectory $sourceStage
    Write-JsonFile (Join-Path $runRoot "staging-summary.json") $staging

    $results = [Collections.Generic.List[object]]::new()
    foreach ($form in @($manifest.Forms)) {
        Write-Host "UserForm verification: $($form.Name)"
        $results.Add((Test-UserForm $form $executablePath))
    }

    $canonicalAfter = Get-IguanaTexCanonicalInputHash -ProjectRoot $projectRoot
    Write-JsonFile (Join-Path $runRoot "canonical-hash.after.json") $canonicalAfter
    if ([string]$canonicalBefore.Hash -cne [string]$canonicalAfter.Hash) {
        throw "Canonical UserForm inputs changed during verification."
    }

    $summary = [ordered]@{
        schemaVersion = 1
        status = "passed"
        formCount = $results.Count
        canonicalCreateSemantic = @($results | Where-Object canonicalCreateSemantic).Count
        noOpSemantic = @($results | Where-Object noOpSemantic).Count
        positionalPatchSemantic = @($results | Where-Object positionalPatchSemantic).Count
        optionPatchSemantic = @($results | Where-Object optionPatchSemantic).Count
        watchSemantic = $(if ($SkipWatch) { 0 } else { @($results | Where-Object watchSemantic).Count })
        watchSkipped = [bool]$SkipWatch
        templateSemantic = @($results | Where-Object templateSemantic).Count
        canonicalSourceChanges = 0
        canonicalHash = [string]$canonicalAfter.Hash
        frxEditCommit = [string]$provenance.HeadCommit
        frxEditExecutableSha256 = [string]$provenance.Sha256
        detailedArtifactsRetained = [bool]$KeepArtifacts
        artifactDirectory = $(if ($KeepArtifacts) { [IO.Path]::GetFullPath($runRoot) } else { $null })
        forms = @($results)
    }
    Write-JsonFile (Join-Path $runRoot "summary.json") $summary
    Write-JsonFile $latestSummaryPath $summary
    $completed = $true
}
catch {
    $caughtError = $_
    $failure = [ordered]@{
        schemaVersion = 1
        status = "failed"
        message = $_.Exception.Message
        watchSkipped = [bool]$SkipWatch
        artifactDirectory = [IO.Path]::GetFullPath($runRoot)
    }
    Write-JsonFile (Join-Path $runRoot "failure.json") $failure
    Write-JsonFile $latestSummaryPath $failure
}
finally {
    if ($null -ne $canonicalBefore) {
        try {
            $canonicalFinal = Get-IguanaTexCanonicalInputHash -ProjectRoot $projectRoot
            Write-JsonFile (Join-Path $runRoot "canonical-hash.after.json") $canonicalFinal
            if ([string]$canonicalBefore.Hash -cne [string]$canonicalFinal.Hash) {
                $integrityError = "Canonical UserForm inputs changed during verification."
            }
        }
        catch {
            $integrityError = "Could not prove canonical input immutability after verification: $($_.Exception.Message)"
        }
    }

    if ($null -ne $integrityError) {
        $completed = $false
        $integrityFailure = [ordered]@{
            schemaVersion = 1
            status = "failed"
            message = $integrityError
            watchSkipped = [bool]$SkipWatch
            artifactDirectory = [IO.Path]::GetFullPath($runRoot)
        }
        Write-JsonFile (Join-Path $runRoot "failure.json") $integrityFailure
        Write-JsonFile $latestSummaryPath $integrityFailure
    }

    if ($completed -and -not $KeepArtifacts -and (Test-Path -LiteralPath $runRoot)) {
        Remove-Item -LiteralPath $runRoot -Recurse -Force
    }
}

if ($null -ne $caughtError) {
    if ($null -ne $integrityError) {
        throw "$($caughtError.Exception.Message) Additionally, $integrityError"
    }
    throw $caughtError
}
if ($null -ne $integrityError) {
    throw $integrityError
}

$watchSummary = if ($SkipWatch) { "watch skipped" } else { "bounded watch" }
Write-Host "PASS: strict create, no-op, exported patch, $watchSummary, and template recreation for $($summary.formCount) UserForms."
Write-Host "Summary: $latestSummaryPath"
if ($KeepArtifacts) {
    Write-Host "Detailed artifacts: $runRoot"
}

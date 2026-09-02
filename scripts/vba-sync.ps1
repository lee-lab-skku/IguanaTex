#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("import", "export", "verify")]
    [string]$Action,

    [switch]$Prune,

    [string[]]$UpdateUserForm = @(),

    [string[]]$UpdateFrx = @(),

    [switch]$VerifyFrx,

    [switch]$DryRun,

    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

foreach ($modulePath in @(
    (Join-Path $PSScriptRoot "lib\IguanaTex.Office.psm1"),
    (Join-Path $PSScriptRoot "lib\IguanaTex.UserForms.psm1")
)) {
    Import-Module $modulePath -Force -DisableNameChecking -ErrorAction Stop
}

$VBEXT_CT_DOCUMENT = 100

$MSO_TRUE = -1
$MSO_FALSE = 0

function Get-ProjectLayout {
    param([string]$ScriptDirectory)

    $projectRoot = (Resolve-Path -LiteralPath (
        Join-Path $ScriptDirectory ".."
    )).Path
    $canonicalDirectory = Join-Path $projectRoot "src"
    $sourceDirectory = Join-Path $projectRoot ".build\vba-source"
    $pptmDirectory = Join-Path $projectRoot ".build\office"
    $pptmFiles = @()

    if (Test-Path -LiteralPath $pptmDirectory -PathType Container) {
        $pptmFiles = @(
            Get-ChildItem -LiteralPath $pptmDirectory -Filter "*.pptm" -File |
            Where-Object { -not $_.Name.StartsWith('~$') } |
            Sort-Object Name
        )
    }

    if ($pptmFiles.Count -eq 0) {
        throw "No PPTM container found in build directory: $pptmDirectory"
    }

    if ($pptmFiles.Count -gt 1) {
        $names = ($pptmFiles.Name -join ", ")
        throw "Expected one PPTM container in build directory, found: $names"
    }

    return [PSCustomObject]@{
        ProjectRoot = $projectRoot
        CanonicalDirectory = $canonicalDirectory
        SourceDirectory = $sourceDirectory
        PptmPath = $pptmFiles[0].FullName
    }
}

function Get-ExactFileHash {
    param([string]$Path)

    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Test-FilesEqual {
    param(
        [string]$Left,
        [string]$Right
    )

    if (-not (Test-Path -LiteralPath $Left -PathType Leaf)) {
        return $false
    }

    if (-not (Test-Path -LiteralPath $Right -PathType Leaf)) {
        return $false
    }

    $leftInfo = Get-Item -LiteralPath $Left
    $rightInfo = Get-Item -LiteralPath $Right

    if ($leftInfo.Length -ne $rightInfo.Length) {
        return $false
    }

    return ((Get-ExactFileHash $Left) -eq (Get-ExactFileHash $Right))
}

function Normalize-RequestedNames {
    param([string[]]$Values)

    $result = @()

    foreach ($value in $Values) {
        if ($null -eq $value) {
            continue
        }

        foreach ($part in ($value -split ",")) {
            $trimmed = $part.Trim()

            if ($trimmed.Length -gt 0) {
                $result += $trimmed
            }
        }
    }

    return @($result | Select-Object -Unique)
}

function New-StagingDirectory {
    $name = "IguanaTex-vba-sync-" + [Guid]::NewGuid().ToString("N")
    $path = Join-Path ([System.IO.Path]::GetTempPath()) $name

    [void](New-Item -ItemType Directory -Path $path)

    return $path
}

function Export-ProjectToDirectory {
    param(
        [object]$Project,
        [string]$Directory
    )

    $components = $null

    try {
        $components = $Project.VBComponents

        for ($i = 1; $i -le $components.Count; $i++) {
            $component = $null

            try {
                $component = $components.Item($i)
                $type = [int]$component.Type
                $name = [string]$component.Name
                $extension = Get-VbaComponentExtension $type

                if ($null -eq $extension) {
                    if ($type -ne $VBEXT_CT_DOCUMENT) {
                        Write-Warning "Skipping unsupported component: $name (type $type)"
                    }

                    continue
                }

                $destination = Join-Path $Directory ($name + $extension)
                $component.Export($destination)

                Write-Host ("Exported: {0}" -f ($name + $extension))
            }
            finally {
                Release-ComObject $component
            }
        }
    }
    finally {
        Release-ComObject $components
    }
}

function Copy-FileIfChanged {
    param(
        [string]$Source,
        [string]$Destination,
        [switch]$WhatIfOnly
    )

    if (Test-FilesEqual $Source $Destination) {
        return $false
    }

    $destinationExisted = Test-Path -LiteralPath $Destination -PathType Leaf

    if ($WhatIfOnly) {
        if ($destinationExisted) {
            Write-Host "Would update: $Destination"
        }
        else {
            Write-Host "Would add: $Destination"
        }

        return $true
    }

    Copy-Item -LiteralPath $Source -Destination $Destination -Force

    if ($destinationExisted) {
        Write-Host "Updated: $Destination"
    }
    else {
        Write-Host "Added: $Destination"
    }

    return $true
}

function Remove-FileIfPresent {
    param(
        [string]$Path,
        [switch]$WhatIfOnly
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }

    if ($WhatIfOnly) {
        Write-Host "Would remove: $Path"
        return $true
    }

    Remove-Item -LiteralPath $Path -Force
    Write-Host "Removed: $Path"

    return $true
}

function Get-TextComponentFiles {
    param([string]$Directory)

    $files = @()
    foreach ($pattern in @("*.bas", "*.cls")) {
        $files += @(Get-ChildItem -LiteralPath $Directory -Filter $pattern -File)
    }
    return @($files | Sort-Object Extension, Name)
}

function Sync-ExportedTextComponents {
    param(
        [string]$StagingDirectory,
        [string]$CanonicalDirectory,
        [switch]$PruneFiles,
        [switch]$WhatIfOnly
    )

    $stageFiles = @(Get-TextComponentFiles $StagingDirectory)
    $stageKeys = @{}
    foreach ($file in $stageFiles) {
        $key = $file.Name.ToLowerInvariant()
        $stageKeys[$key] = $true
        [void](Copy-FileIfChanged `
            -Source $file.FullName `
            -Destination (Join-Path $CanonicalDirectory $file.Name) `
            -WhatIfOnly:$WhatIfOnly)
    }

    if ($PruneFiles) {
        foreach ($file in @(Get-TextComponentFiles $CanonicalDirectory)) {
            if (-not $stageKeys.ContainsKey($file.Name.ToLowerInvariant())) {
                [void](Remove-FileIfPresent -Path $file.FullName -WhatIfOnly:$WhatIfOnly)
            }
        }
    }
}

function Get-NormalizedVbaText {
    param([string]$Path)

    $text = Get-Content -Raw -LiteralPath $Path
    $text = (($text -replace "`r`n", "`n") -replace "`r", "`n").TrimEnd([char[]]"`n")

    # The VBE rewrites identifier capitalization to the spelling most recently
    # registered in the project. This can change an otherwise identical export
    # (for example, UseDvi -> UseDVI) after an import/save round trip. VBA
    # identifiers and keywords are case-insensitive, but string literals and
    # apostrophe comments are not comparison noise, so preserve those regions.
    $builder = New-Object System.Text.StringBuilder
    $inString = $false
    $inComment = $false
    $canStartRemComment = $true

    for ($index = 0; $index -lt $text.Length; $index++) {
        $character = $text[$index]

        if ($character -eq "`n") {
            [void]$builder.Append($character)
            $inComment = $false
            $canStartRemComment = $true
            continue
        }

        if ($inComment) {
            [void]$builder.Append($character)
            continue
        }

        if ($inString) {
            [void]$builder.Append($character)
            if ($character -eq '"') {
                if ($index + 1 -lt $text.Length -and $text[$index + 1] -eq '"') {
                    [void]$builder.Append($text[$index + 1])
                    $index++
                }
                else {
                    $inString = $false
                }
            }
            continue
        }

        if ($character -eq "'") {
            [void]$builder.Append($character)
            $inComment = $true
            continue
        }

        if ($character -eq '"') {
            [void]$builder.Append($character)
            $inString = $true
            $canStartRemComment = $false
            continue
        }

        $codePoint = [int]$character
        $isIdentifierStart = (
            ($codePoint -ge [int][char]'A' -and $codePoint -le [int][char]'Z') -or
            ($codePoint -ge [int][char]'a' -and $codePoint -le [int][char]'z') -or
            $character -eq '_'
        )
        if ($isIdentifierStart) {
            $tokenStart = $index
            while ($index + 1 -lt $text.Length) {
                $nextCharacter = $text[$index + 1]
                $nextCodePoint = [int]$nextCharacter
                $isIdentifierPart = (
                    ($nextCodePoint -ge [int][char]'A' -and $nextCodePoint -le [int][char]'Z') -or
                    ($nextCodePoint -ge [int][char]'a' -and $nextCodePoint -le [int][char]'z') -or
                    ($nextCodePoint -ge [int][char]'0' -and $nextCodePoint -le [int][char]'9') -or
                    $nextCharacter -eq '_'
                )
                if (-not $isIdentifierPart) {
                    break
                }
                $index++
            }

            $token = $text.Substring($tokenStart, $index - $tokenStart + 1)
            [void]$builder.Append($token.ToLowerInvariant())
            if ($canStartRemComment -and $token -ieq "Rem") {
                $inComment = $true
            }
            elseif ($token -ieq "Then" -or $token -ieq "Else") {
                $canStartRemComment = $true
            }
            else {
                $canStartRemComment = $false
            }
            continue
        }

        if ([char]::IsWhiteSpace($character)) {
            [void]$builder.Append($character)
            continue
        }

        if ($character -eq ':') {
            [void]$builder.Append($character)
            $canStartRemComment = -not (
                $index + 1 -lt $text.Length -and $text[$index + 1] -eq '='
            )
            continue
        }

        if ($codePoint -ge [int][char]'A' -and $codePoint -le [int][char]'Z') {
            [void]$builder.Append([char]($codePoint + 32))
        }
        else {
            [void]$builder.Append($character)
        }
        if (-not [char]::IsDigit($character) -or -not $canStartRemComment) {
            $canStartRemComment = $false
        }
    }

    return $builder.ToString()
}

function Test-VbaTextEqual {
    param(
        [string]$Left,
        [string]$Right
    )

    if (-not (Test-Path -LiteralPath $Left -PathType Leaf)) {
        return $false
    }
    if (-not (Test-Path -LiteralPath $Right -PathType Leaf)) {
        return $false
    }

    # Preserve the historical byte-exact comparison as the fast and preferred
    # path. Only fall back to VBA-aware comparison for VBE casing churn.
    if (Test-FilesEqual $Left $Right) {
        return $true
    }

    return (Get-NormalizedVbaText $Left) -ceq (Get-NormalizedVbaText $Right)
}

function Assert-ExpectedFormSet {
    param(
        [string]$Directory,
        [string[]]$ExpectedNames
    )

    $expected = @($ExpectedNames | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
    $actual = @(
        Get-ChildItem -LiteralPath $Directory -Filter "*.frm" -File |
        ForEach-Object { $_.BaseName.ToLowerInvariant() } |
        Sort-Object -Unique
    )
    if (($expected -join "`n") -cne ($actual -join "`n")) {
        throw "Exported UserForm set does not match the canonical manifest. Expected: $($expected -join ', '); actual: $($actual -join ', ')"
    }

    foreach ($name in $ExpectedNames) {
        $frxPath = Join-Path $Directory ($name + ".frx")
        if (-not (Test-Path -LiteralPath $frxPath -PathType Leaf)) {
            throw "Exported UserForm is missing its FRX companion: $frxPath"
        }
    }
}

function Verify-ExportedTree {
    param(
        [string]$StagingDirectory,
        [string]$GeneratedDirectory,
        [string[]]$FormNames,
        [string]$FrxEditPath
    )

    $differenceCount = 0
    $stageFiles = @(Get-TextComponentFiles $StagingDirectory)
    $generatedFiles = @(Get-TextComponentFiles $GeneratedDirectory)
    $stageMap = @{}
    $generatedMap = @{}
    foreach ($file in $stageFiles) { $stageMap[$file.Name.ToLowerInvariant()] = $file }
    foreach ($file in $generatedFiles) { $generatedMap[$file.Name.ToLowerInvariant()] = $file }

    $allKeys = @(@($stageMap.Keys) + @($generatedMap.Keys) | Sort-Object -Unique)
    foreach ($key in $allKeys) {
        if (-not $stageMap.ContainsKey($key)) {
            Write-Host "EXTRA in canonical source: $($generatedMap[$key].Name)"
            $differenceCount++
        }
        elseif (-not $generatedMap.ContainsKey($key)) {
            Write-Host "MISSING in canonical source: $($stageMap[$key].Name)"
            $differenceCount++
        }
        elseif (-not (Test-VbaTextEqual $stageMap[$key].FullName $generatedMap[$key].FullName)) {
            Write-Host "DIFF: $($stageMap[$key].Name)"
            $differenceCount++
        }
    }

    Assert-ExpectedFormSet -Directory $StagingDirectory -ExpectedNames $FormNames
    Assert-ExpectedFormSet -Directory $GeneratedDirectory -ExpectedNames $FormNames
    $reportRoot = Join-Path $StagingDirectory ".canonical-comparison"
    [void](New-Item -ItemType Directory -Path $reportRoot -Force)

    foreach ($name in $FormNames) {
        $caseRoot = Join-Path $reportRoot $name
        $generatedRoot = Join-Path $caseRoot "generated"
        $exportedRoot = Join-Path $caseRoot "exported"
        [void](New-Item -ItemType Directory -Path $generatedRoot, $exportedRoot -Force)

        $generatedForm = Join-Path $GeneratedDirectory ($name + ".frm")
        $exportedForm = Join-Path $StagingDirectory ($name + ".frm")
        $generatedTemplate = Join-Path $generatedRoot ($name + ".template.json")
        $exportedTemplate = Join-Path $exportedRoot ($name + ".template.json")

        [void](Compare-IguanaTexUserFormSemantics `
            -ProjectRoot $projectRoot `
            -FrxEditPath $FrxEditPath `
            -OriginalFormPath $generatedForm `
            -CandidateFormPath $exportedForm `
            -ArtifactsDirectory (Join-Path $caseRoot "semantics") `
            -ReportPath (Join-Path $caseRoot "semantic-report.json"))

        Invoke-FrxEditCommand -ExecutablePath $FrxEditPath -Arguments @(
            "inspect", $generatedForm, "--mode", "strict", "--as-template", "--out", $generatedTemplate)
        Invoke-FrxEditCommand -ExecutablePath $FrxEditPath -Arguments @(
            "inspect", $exportedForm, "--mode", "strict", "--as-template", "--out", $exportedTemplate)
        $generatedVba = [IO.Path]::ChangeExtension($generatedTemplate, ".vba")
        $exportedVba = [IO.Path]::ChangeExtension($exportedTemplate, ".vba")
        Assert-IguanaTexGeneratedVba `
            -FormPath $generatedForm `
            -CanonicalVbaPath $generatedVba
        Assert-IguanaTexGeneratedVba `
            -FormPath $exportedForm `
            -CanonicalVbaPath $exportedVba
        if (-not (Test-VbaTextEqual $generatedVba $exportedVba)) {
            Write-Host "DIFF UserForm VBA: $name"
            $differenceCount++
        }
    }

    if ($differenceCount -eq 0) {
        Write-Host "VERIFY OK"
        return 0
    }

    Write-Host ("VERIFY FAILED: {0} difference(s)" -f $differenceCount)
    return 2
}

if ($PSVersionTable.PSEdition -eq "Core" -and -not $IsWindows) {
    throw "This script requires Windows PowerPoint COM automation."
}

if ($UpdateFrx.Count -gt 0) {
    throw "-UpdateFrx was retired when UserForm JSON/VBA became canonical. Use -UpdateUserForm with the export action."
}

if ($VerifyFrx) {
    throw "-VerifyFrx was retired. The verify action now compares UserForms semantically and never uses FRX byte equality."
}

if ($Action -ne "export" -and $UpdateUserForm.Count -gt 0) {
    throw "-UpdateUserForm is valid only with the export action."
}

if ($Action -eq "verify" -and $Prune) {
    throw "-Prune is not valid with the verify action."
}

$layout = Get-ProjectLayout $PSScriptRoot
$projectRoot = $layout.ProjectRoot
$canonicalDirectory = $layout.CanonicalDirectory
$sourceDirectory = $layout.SourceDirectory
$pptmResolved = $layout.PptmPath
$manifest = Assert-IguanaTexUserFormManifest -ProjectRoot $projectRoot
$formNames = @($manifest.Forms | ForEach-Object { [string]$_.Name })
$requestedFormNames = @(Normalize-RequestedNames $UpdateUserForm)
$canonicalNameMap = @{}
foreach ($name in $formNames) { $canonicalNameMap[$name.ToLowerInvariant()] = $name }
$resolvedNames = @()
foreach ($name in $requestedFormNames) {
    $key = $name.ToLowerInvariant()
    if ($key -eq "all") {
        throw "-UpdateUserForm requires explicit form names; 'all' is intentionally not supported."
    }
    if (-not $canonicalNameMap.ContainsKey($key)) {
        throw "Requested UserForm is not listed in the canonical manifest: $name"
    }
    $resolvedNames += $canonicalNameMap[$key]
}
$requestedFormNames = @($resolvedNames | Select-Object -Unique)

$codec = Invoke-FrxEditBuild -ProjectRoot $projectRoot -Mode Pinned
$frxEditPath = [string]$codec.ExecutablePath
[void](New-IguanaTexUserFormStaging `
    -ProjectRoot $projectRoot `
    -FrxEditPath $frxEditPath `
    -OutputDirectory $sourceDirectory)

if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
    throw "Generated VBA source directory not found: $sourceDirectory"
}

$ppt = $null
$presentations = $null
$presentation = $null
$project = $null
$stagingDirectory = $null
$importSaved = $false
$exitCode = 0

try {
    Write-Host "Action: $Action"
    Write-Host "Root:   $projectRoot"
    Write-Host "PPTM:   $pptmResolved"
    Write-Host "Canonical: $canonicalDirectory"
    Write-Host "Generated: $sourceDirectory"

    if ($DryRun) {
        Write-Host "Mode:   dry-run"
    }

    if ($requestedFormNames.Count -gt 0) {
        Write-Host ("UserForms: " + ($requestedFormNames -join ", "))
    }

    Write-Host ""

    $ppt = New-Object -ComObject PowerPoint.Application

    if ($Visible) {
        $ppt.Visible = $MSO_TRUE
    }

    $presentations = $ppt.Presentations

    $readOnly = $MSO_TRUE

    if ($Action -eq "import" -and -not $DryRun) {
        $readOnly = $MSO_FALSE
    }

    $withWindow = $MSO_FALSE

    if ($Visible) {
        $withWindow = $MSO_TRUE
    }

    $presentation = $presentations.Open(
        $pptmResolved,
        $readOnly,
        $MSO_FALSE,
        $withWindow
    )

    try {
        $project = $presentation.VBProject
    }
    catch {
        throw (
            "Cannot access VBProject. Enable 'Trust access to the VBA project object model'. " +
            $_.Exception.Message
        )
    }

    switch ($Action) {
        "import" {
            Import-VbaSourceTree `
                -Project $project `
                -SourceDirectory $sourceDirectory `
                -PruneComponents:$Prune `
                -WhatIfOnly:$DryRun

            if (-not $DryRun) {
                $presentation.Save()
                $importSaved = $true
                Write-Host ""
                Write-Host "IMPORT OK"
            }
            else {
                Write-Host ""
                Write-Host "IMPORT DRY-RUN OK"
            }
        }

        "export" {
            $stagingDirectory = New-StagingDirectory

            Write-Host "Stage:  $stagingDirectory"
            Write-Host ""

            Export-ProjectToDirectory `
                -Project $project `
                -Directory $stagingDirectory

            Assert-ExpectedFormSet `
                -Directory $stagingDirectory `
                -ExpectedNames $formNames

            Write-Host ""

            Sync-ExportedTextComponents `
                -StagingDirectory $stagingDirectory `
                -CanonicalDirectory $canonicalDirectory `
                -PruneFiles:$Prune `
                -WhatIfOnly:$DryRun

            if ($requestedFormNames.Count -gt 0) {
                [void](Export-IguanaTexCanonicalUserForms `
                    -ProjectRoot $projectRoot `
                    -FrxEditPath $frxEditPath `
                    -InputDirectory $stagingDirectory `
                    -Form $requestedFormNames `
                    -DryRun:$DryRun)
            }

            Write-Host ""

            if ($DryRun) {
                Write-Host "EXPORT DRY-RUN OK"
            }
            else {
                Write-Host "EXPORT OK"
            }
        }

        "verify" {
            $stagingDirectory = New-StagingDirectory

            Write-Host "Stage:  $stagingDirectory"
            Write-Host ""

            Export-ProjectToDirectory `
                -Project $project `
                -Directory $stagingDirectory

            Write-Host ""

            $exitCode = Verify-ExportedTree `
                -StagingDirectory $stagingDirectory `
                -GeneratedDirectory $sourceDirectory `
                -FormNames $formNames `
                -FrxEditPath $frxEditPath
        }
    }
}
catch {
    Write-Error $_
    $exitCode = 1
}
finally {
    Release-ComObject $project

    if ($null -ne $presentation) {
        if ($Action -eq "import" -and -not $importSaved) {
            try {
                $presentation.Saved = $MSO_TRUE
            }
            catch {
            }
        }

        try {
            $presentation.Close()
        }
        catch {
        }

        Release-ComObject $presentation
    }

    Release-ComObject $presentations

    if ($null -ne $ppt) {
        try {
            $ppt.Quit()
        }
        catch {
        }

        Release-ComObject $ppt
    }

    if ($null -ne $stagingDirectory) {
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            try {
                Remove-Item -LiteralPath $stagingDirectory -Recurse -Force
            }
            catch {
                Write-Warning "Could not remove staging directory: $stagingDirectory"
            }
        }
    }

    Invoke-ComGarbageCollection
}

exit $exitCode

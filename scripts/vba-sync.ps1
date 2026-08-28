#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("import", "export", "verify")]
    [string]$Action,

    [switch]$Prune,

    [string[]]$UpdateFrx = @(),

    [switch]$VerifyFrx,

    [switch]$DryRun,

    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Import-Module (
    Join-Path $PSScriptRoot "lib\IguanaTex.Office.psm1"
) -Force -DisableNameChecking -ErrorAction Stop

$VBEXT_CT_DOCUMENT = 100

$MSO_TRUE = -1
$MSO_FALSE = 0

function Get-ProjectLayout {
    param([string]$ScriptDirectory)

    $projectRoot = (Resolve-Path -LiteralPath (
        Join-Path $ScriptDirectory ".."
    )).Path
    $sourceDirectory = Join-Path $projectRoot "src"
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

function Normalize-UpdateFrxNames {
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

function Sync-ExportedTree {
    param(
        [string]$StagingDirectory,
        [string]$CanonicalDirectory,
        [switch]$PruneFiles,
        [string[]]$FrxNames,
        [switch]$WhatIfOnly
    )

    if (-not (Test-Path -LiteralPath $CanonicalDirectory -PathType Container)) {
        if ($WhatIfOnly) {
            Write-Host "Would create directory: $CanonicalDirectory"
        }
        else {
            [void](New-Item -ItemType Directory -Path $CanonicalDirectory)
            Write-Host "Created directory: $CanonicalDirectory"
        }
    }

    $stageFiles = @(Get-VbaSourceFiles $StagingDirectory)
    $stageKeys = @{}

    foreach ($file in $stageFiles) {
        $key = $file.Name.ToLowerInvariant()
        $stageKeys[$key] = $true

        $destination = Join-Path $CanonicalDirectory $file.Name

        [void](Copy-FileIfChanged `
            -Source $file.FullName `
            -Destination $destination `
            -WhatIfOnly:$WhatIfOnly)
    }

    if ($PruneFiles -and
        (Test-Path -LiteralPath $CanonicalDirectory -PathType Container)) {

        $canonicalTextFiles = @(Get-VbaSourceFiles $CanonicalDirectory)

        foreach ($file in $canonicalTextFiles) {
            $key = $file.Name.ToLowerInvariant()

            if (-not $stageKeys.ContainsKey($key)) {
                [void](Remove-FileIfPresent `
                    -Path $file.FullName `
                    -WhatIfOnly:$WhatIfOnly)

                if ($file.Extension.ToLowerInvariant() -eq ".frm") {
                    $frxPath = Join-Path $CanonicalDirectory ($file.BaseName + ".frx")

                    [void](Remove-FileIfPresent `
                        -Path $frxPath `
                        -WhatIfOnly:$WhatIfOnly)
                }
            }
        }
    }

    $forms = @(
        Get-ChildItem -LiteralPath $StagingDirectory -Filter "*.frm" -File |
        Sort-Object Name
    )

    $requested = @{}

    foreach ($name in $FrxNames) {
        $requested[$name.ToLowerInvariant()] = $true
    }

    $updateAll = $requested.ContainsKey("all")
    $stageFormNames = @{}

    foreach ($form in $forms) {
        $formName = $form.BaseName
        $stageFormNames[$formName.ToLowerInvariant()] = $true

        $stageFrx = Join-Path $StagingDirectory ($formName + ".frx")
        $canonicalFrx = Join-Path $CanonicalDirectory ($formName + ".frx")

        $hasStageFrx = Test-Path -LiteralPath $stageFrx -PathType Leaf
        $hasCanonicalFrx = Test-Path -LiteralPath $canonicalFrx -PathType Leaf

        $explicitUpdate = (
            $updateAll -or
            $requested.ContainsKey($formName.ToLowerInvariant())
        )

        if (-not $hasCanonicalFrx) {
            if ($hasStageFrx) {
                if ($WhatIfOnly) {
                    Write-Host "Would add FRX: $canonicalFrx"
                }
                else {
                    Copy-Item -LiteralPath $stageFrx -Destination $canonicalFrx -Force
                    Write-Host "Added FRX: $canonicalFrx"
                }
            }

            continue
        }

        if (-not $explicitUpdate) {
            Write-Host "Preserved FRX: $canonicalFrx"
            continue
        }

        if ($hasStageFrx) {
            if (Test-FilesEqual $stageFrx $canonicalFrx) {
                Write-Host "FRX unchanged: $canonicalFrx"
            }
            elseif ($WhatIfOnly) {
                Write-Host "Would update FRX: $canonicalFrx"
            }
            else {
                Copy-Item -LiteralPath $stageFrx -Destination $canonicalFrx -Force
                Write-Host "Updated FRX: $canonicalFrx"
            }
        }
        else {
            [void](Remove-FileIfPresent `
                -Path $canonicalFrx `
                -WhatIfOnly:$WhatIfOnly)
        }
    }

    if (-not $updateAll) {
        foreach ($name in $requested.Keys) {
            if (-not $stageFormNames.ContainsKey($name)) {
                throw "Requested FRX form not found in PPTM export: $name"
            }
        }
    }
}

function Verify-ExportedTree {
    param(
        [string]$StagingDirectory,
        [string]$CanonicalDirectory,
        [switch]$IncludeFrx
    )

    if (-not (Test-Path -LiteralPath $CanonicalDirectory -PathType Container)) {
        Write-Host "Missing source directory: $CanonicalDirectory"
        return 1
    }

    $differenceCount = 0

    $stageFiles = @(Get-VbaSourceFiles $StagingDirectory)
    $canonicalFiles = @(Get-VbaSourceFiles $CanonicalDirectory)

    $stageMap = @{}
    $canonicalMap = @{}

    foreach ($file in $stageFiles) {
        $stageMap[$file.Name.ToLowerInvariant()] = $file
    }

    foreach ($file in $canonicalFiles) {
        $canonicalMap[$file.Name.ToLowerInvariant()] = $file
    }

    $allKeys = @(
        @($stageMap.Keys) + @($canonicalMap.Keys) |
        Sort-Object -Unique
    )

    foreach ($key in $allKeys) {
        $inStage = $stageMap.ContainsKey($key)
        $inCanonical = $canonicalMap.ContainsKey($key)

        if (-not $inStage) {
            Write-Host "EXTRA in source: $($canonicalMap[$key].Name)"
            $differenceCount++
            continue
        }

        if (-not $inCanonical) {
            Write-Host "MISSING in source: $($stageMap[$key].Name)"
            $differenceCount++
            continue
        }

        if (-not (Test-FilesEqual `
            $stageMap[$key].FullName `
            $canonicalMap[$key].FullName)) {

            Write-Host "DIFF: $($stageMap[$key].Name)"
            $differenceCount++
        }
    }

    if ($IncludeFrx) {
        Write-Warning "FRX verification is byte-level only and may be nondeterministic."

        $stageFrxFiles = @(
            Get-ChildItem -LiteralPath $StagingDirectory -Filter "*.frx" -File
        )

        $canonicalFrxFiles = @(
            Get-ChildItem -LiteralPath $CanonicalDirectory -Filter "*.frx" -File
        )

        $stageFrxMap = @{}
        $canonicalFrxMap = @{}

        foreach ($file in $stageFrxFiles) {
            $stageFrxMap[$file.Name.ToLowerInvariant()] = $file
        }

        foreach ($file in $canonicalFrxFiles) {
            $canonicalFrxMap[$file.Name.ToLowerInvariant()] = $file
        }

        $allFrxKeys = @(
            @($stageFrxMap.Keys) + @($canonicalFrxMap.Keys) |
            Sort-Object -Unique
        )

        foreach ($key in $allFrxKeys) {
            $inStage = $stageFrxMap.ContainsKey($key)
            $inCanonical = $canonicalFrxMap.ContainsKey($key)

            if (-not $inStage) {
                Write-Host "EXTRA FRX in source: $($canonicalFrxMap[$key].Name)"
                $differenceCount++
                continue
            }

            if (-not $inCanonical) {
                Write-Host "MISSING FRX in source: $($stageFrxMap[$key].Name)"
                $differenceCount++
                continue
            }

            if (-not (Test-FilesEqual `
                $stageFrxMap[$key].FullName `
                $canonicalFrxMap[$key].FullName)) {

                Write-Host "DIFF FRX: $($stageFrxMap[$key].Name)"
                $differenceCount++
            }
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

if ($Action -ne "export" -and $UpdateFrx.Count -gt 0) {
    throw "-UpdateFrx is valid only with the export action."
}

if ($Action -ne "verify" -and $VerifyFrx) {
    throw "-VerifyFrx is valid only with the verify action."
}

if ($Action -eq "verify" -and $Prune) {
    throw "-Prune is not valid with the verify action."
}

$layout = Get-ProjectLayout $PSScriptRoot
$projectRoot = $layout.ProjectRoot
$sourceDirectory = $layout.SourceDirectory
$pptmResolved = $layout.PptmPath
$frxNames = @(Normalize-UpdateFrxNames $UpdateFrx)

if ($Action -in @("import", "verify")) {
    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        throw "VBA source directory not found: $sourceDirectory"
    }
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
    Write-Host "Source: $sourceDirectory"

    if ($DryRun) {
        Write-Host "Mode:   dry-run"
    }

    if ($frxNames.Count -gt 0) {
        Write-Host ("FRX:    " + ($frxNames -join ", "))
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

            Write-Host ""

            Sync-ExportedTree `
                -StagingDirectory $stagingDirectory `
                -CanonicalDirectory $sourceDirectory `
                -PruneFiles:$Prune `
                -FrxNames $frxNames `
                -WhatIfOnly:$DryRun

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
                -CanonicalDirectory $sourceDirectory `
                -IncludeFrx:$VerifyFrx
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

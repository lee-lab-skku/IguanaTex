#requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:VBEXT_CT_STDMODULE = 1
$script:VBEXT_CT_CLASSMODULE = 2
$script:VBEXT_CT_MSFORM = 3
$script:VBEXT_CT_DOCUMENT = 100

function Release-ComObject {
    [CmdletBinding()]
    param([AllowNull()][object]$Object)

    if ($null -ne $Object) {
        try {
            if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
            }
        }
        catch {
        }
    }
}

function Invoke-ComGarbageCollection {
    [CmdletBinding()]
    param()

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Get-VbaComponentExtension {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][int]$Type)

    switch ($Type) {
        $script:VBEXT_CT_STDMODULE { return ".bas" }
        $script:VBEXT_CT_CLASSMODULE { return ".cls" }
        $script:VBEXT_CT_MSFORM { return ".frm" }
        default { return $null }
    }
}

function Get-VbaSourceFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Directory
    )

    $files = @()

    foreach ($pattern in @("*.bas", "*.cls", "*.frm")) {
        $files += @(Get-ChildItem -LiteralPath $Directory -Filter $pattern -File)
    }

    $files = @($files | Sort-Object Extension, Name)

    $duplicates = @(
        $files |
        Group-Object BaseName |
        Where-Object { $_.Count -gt 1 }
    )

    if ($duplicates.Count -gt 0) {
        $names = ($duplicates.Name -join ", ")
        throw "Duplicate VBA component names found: $names"
    }

    return $files
}

function Assert-VbaSourceClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceDirectory
    )

    $sourceFiles = @(Get-VbaSourceFiles -Directory $SourceDirectory)

    if ($sourceFiles.Count -eq 0) {
        throw "No .bas, .cls, or .frm files found in: $SourceDirectory"
    }

    $formNames = @{}

    foreach ($formFile in @($sourceFiles | Where-Object Extension -ieq ".frm")) {
        $formNames[$formFile.BaseName.ToLowerInvariant()] = $true
    }

    foreach ($frxFile in @(
        Get-ChildItem -LiteralPath $SourceDirectory -Filter "*.frx" -File
    )) {
        if (-not $formNames.ContainsKey($frxFile.BaseName.ToLowerInvariant())) {
            throw "Orphan FRX has no canonical form source: $($frxFile.FullName)"
        }
    }

    foreach ($file in $sourceFiles) {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        $nameMatches = [regex]::Matches(
            $text,
            '(?im)^[\t ]*Attribute[\t ]+VB_Name[\t ]*=[\t ]*"([^"\r\n]+)"[\t ]*\r?$'
        )

        if ($nameMatches.Count -ne 1) {
            throw (
                "Expected exactly one Attribute VB_Name in {0}, found {1}." -f
                $file.FullName,
                $nameMatches.Count
            )
        }

        $declaredName = $nameMatches[0].Groups[1].Value

        if ($declaredName -cne $file.BaseName) {
            throw (
                "VBA component name mismatch: {0} declares {1}." -f
                $file.Name,
                $declaredName
            )
        }

        if ($file.Extension -ine ".frm") {
            continue
        }

        $expectedFrxName = $file.BaseName + ".frx"
        $frxPath = Join-Path $SourceDirectory $expectedFrxName

        if (-not (Test-Path -LiteralPath $frxPath -PathType Leaf)) {
            throw "Missing companion FRX for $($file.Name): $frxPath"
        }

        $blobMatches = [regex]::Matches(
            $text,
            '(?im)^[\t ]*OleObjectBlob[\t ]*=[\t ]*"([^"\r\n]+)"(?::[0-9A-F]+)?[\t ]*\r?$'
        )

        if ($blobMatches.Count -ne 1) {
            throw (
                "Expected exactly one OleObjectBlob reference in {0}, found {1}." -f
                $file.FullName,
                $blobMatches.Count
            )
        }

        $referencedFrxName = $blobMatches[0].Groups[1].Value

        if ($referencedFrxName -cne $expectedFrxName) {
            throw (
                "FRX reference mismatch in {0}: expected {1}, found {2}." -f
                $file.Name,
                $expectedFrxName,
                $referencedFrxName
            )
        }
    }
}

function Get-VbaProjectComponentNames {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Project)

    $result = @{}
    $components = $null

    try {
        $components = $Project.VBComponents

        for ($i = 1; $i -le $components.Count; $i++) {
            $component = $null

            try {
                $component = $components.Item($i)
                $type = [int]$component.Type

                if ($type -in @(
                    $script:VBEXT_CT_STDMODULE,
                    $script:VBEXT_CT_CLASSMODULE,
                    $script:VBEXT_CT_MSFORM
                )) {
                    $result[[string]$component.Name] = $type
                }
            }
            finally {
                Release-ComObject $component
            }
        }
    }
    finally {
        Release-ComObject $components
    }

    return $result
}

function Remove-VbaProjectComponent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$Name,
        [switch]$WhatIfOnly
    )

    if ($WhatIfOnly) {
        Write-Host "Would remove component: $Name"
        return
    }

    $components = $null
    $component = $null

    try {
        $components = $Project.VBComponents
        $component = $components.Item($Name)

        if ([int]$component.Type -eq $script:VBEXT_CT_DOCUMENT) {
            throw "Document component cannot be removed: $Name"
        }

        $components.Remove($component)
        Write-Host "Removed component: $Name"
    }
    finally {
        Release-ComObject $component
        Release-ComObject $components
    }
}

function Import-VbaSourceTree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Project,
        [Parameter(Mandatory = $true)][string]$SourceDirectory,
        [switch]$PruneComponents,
        [switch]$WhatIfOnly
    )

    Assert-VbaSourceClosure -SourceDirectory $SourceDirectory
    $sourceFiles = @(Get-VbaSourceFiles -Directory $SourceDirectory)
    $sourceNames = @{}

    foreach ($file in $sourceFiles) {
        $sourceNames[$file.BaseName] = $true
    }

    $projectNames = Get-VbaProjectComponentNames -Project $Project

    if ($PruneComponents) {
        foreach ($name in @($projectNames.Keys | Sort-Object)) {
            if (-not $sourceNames.ContainsKey($name)) {
                Remove-VbaProjectComponent `
                    -Project $Project `
                    -Name $name `
                    -WhatIfOnly:$WhatIfOnly
            }
        }
    }

    foreach ($file in $sourceFiles) {
        $name = $file.BaseName
        $components = $null
        $existing = $null

        try {
            $components = $Project.VBComponents

            try {
                $existing = $components.Item($name)
            }
            catch {
                $existing = $null
            }

            if ($null -ne $existing) {
                if ([int]$existing.Type -eq $script:VBEXT_CT_DOCUMENT) {
                    throw "Document component cannot be replaced: $name"
                }

                if ($WhatIfOnly) {
                    Write-Host "Would replace: $name <- $($file.Name)"
                    continue
                }

                $components.Remove($existing)
                Release-ComObject $existing
                $existing = $null

                $newComponent = $null

                try {
                    $newComponent = $components.Import($file.FullName)

                    Write-Host (
                        "Replaced: {0} <- {1}" -f
                        ([string]$newComponent.Name),
                        $file.Name
                    )

                    if ([string]$newComponent.Name -ne $name) {
                        Write-Warning (
                            "Imported component name differs from file name: {0} -> {1}" -f
                            $name,
                            ([string]$newComponent.Name)
                        )
                    }
                }
                finally {
                    Release-ComObject $newComponent
                }
            }
            else {
                if ($WhatIfOnly) {
                    Write-Host "Would add: $name <- $($file.Name)"
                    continue
                }

                $newComponent = $null

                try {
                    $newComponent = $components.Import($file.FullName)

                    Write-Host (
                        "Added: {0} <- {1}" -f
                        ([string]$newComponent.Name),
                        $file.Name
                    )

                    if ([string]$newComponent.Name -ne $name) {
                        Write-Warning (
                            "Imported component name differs from file name: {0} -> {1}" -f
                            $name,
                            ([string]$newComponent.Name)
                        )
                    }
                }
                finally {
                    Release-ComObject $newComponent
                }
            }
        }
        finally {
            Release-ComObject $existing
            Release-ComObject $components
        }
    }
}

Export-ModuleMember -Function @(
    "Release-ComObject",
    "Invoke-ComGarbageCollection",
    "Get-VbaComponentExtension",
    "Get-VbaSourceFiles",
    "Get-VbaProjectComponentNames",
    "Remove-VbaProjectComponent",
    "Import-VbaSourceTree",
    "Assert-VbaSourceClosure"
)

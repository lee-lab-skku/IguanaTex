#requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:UserFormManifestRelativePath = "office\forms\manifest.json"
$script:FrxEditRelativePath = "tools\frx-edit"
$script:FrxEditProjectRelativePath =
    "tools\frx-edit\src\FrxEdit.Cli\FrxEdit.Cli.csproj"
$script:GeneratedTreeMarker = ".iguanatex-userforms-staging.json"

function Test-IguanaTexPathWithinRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar
    )
    $prefix = $fullRoot + [IO.Path]::DirectorySeparatorChar

    return $fullPath.Equals($fullRoot, [StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)
}

function Get-IguanaTexRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar
    )

    if (-not (Test-IguanaTexPathWithinRoot -Path $fullPath -Root $fullRoot)) {
        throw "Path is outside the expected root '$fullRoot': $fullPath"
    }

    if ($fullPath.Equals($fullRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return "."
    }

    return $fullPath.Substring($fullRoot.Length + 1).Replace("\", "/")
}

function Resolve-IguanaTexRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        throw "$Description must be a nonempty relative path."
    }

    if ($RelativePath.Contains("\")) {
        throw "$Description must use forward slashes: $RelativePath"
    }

    if (
        [IO.Path]::IsPathRooted($RelativePath) -or
        $RelativePath.StartsWith("/", [StringComparison]::Ordinal) -or
        $RelativePath -match '^[A-Za-z]:'
    ) {
        throw "$Description must be relative: $RelativePath"
    }

    $segments = @($RelativePath -split "/")
    if (
        $segments.Count -eq 0 -or
        @($segments | Where-Object {
            [string]::IsNullOrWhiteSpace($_) -or $_ -eq "." -or $_ -eq ".."
        }).Count -gt 0
    ) {
        throw "$Description contains an empty or traversal segment: $RelativePath"
    }

    try {
        $nativeRelativePath = $RelativePath.Replace(
            "/",
            [string][IO.Path]::DirectorySeparatorChar
        )
        $resolved = [IO.Path]::GetFullPath((Join-Path $Root $nativeRelativePath))
    }
    catch {
        throw ("Invalid {0} '{1}': {2}" -f
            $Description,
            $RelativePath,
            $_.Exception.Message)
    }

    if (-not (Test-IguanaTexPathWithinRoot -Path $resolved -Root $Root)) {
        throw "$Description escapes its root: $RelativePath"
    }

    return $resolved
}

function Get-IguanaTexObjectPropertyValue {
    param(
        [AllowNull()][object]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-IguanaTexFileAssetReferences {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return
    }

    if ($Value -is [string]) {
        if ($Value.StartsWith("file://", [StringComparison]::OrdinalIgnoreCase)) {
            Write-Output $Value
        }
        return
    }

    if ($Value -is [Collections.IDictionary]) {
        foreach ($entryValue in $Value.Values) {
            Get-IguanaTexFileAssetReferences -Value $entryValue
        }
        return
    }

    if ($Value -is [Collections.IEnumerable]) {
        foreach ($item in $Value) {
            Get-IguanaTexFileAssetReferences -Value $item
        }
        return
    }

    if ($Value -is [ValueType]) {
        return
    }

    foreach ($property in $Value.PSObject.Properties) {
        Get-IguanaTexFileAssetReferences -Value $property.Value
    }
}

function Read-IguanaTexJsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description
    )

    try {
        return Get-Content -Raw -LiteralPath $Path -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Could not parse {0} '{1}': {2}" -f
            $Description,
            $Path,
            $_.Exception.Message)
    }
}

function Assert-IguanaTexFormSource {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [Parameter(Mandatory = $true)][string]$VbaPath,
        [Parameter(Mandatory = $true)][string]$AssetRoot
    )

    if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
        throw "UserForm template not found for '$Name': $TemplatePath"
    }
    if (-not (Test-Path -LiteralPath $VbaPath -PathType Leaf)) {
        throw "UserForm VBA not found for '$Name': $VbaPath"
    }

    $template = Read-IguanaTexJsonFile `
        -Path $TemplatePath `
        -Description "UserForm template"
    $properties = Get-IguanaTexObjectPropertyValue `
        -Object $template `
        -Name "properties"
    if ($null -eq $properties) {
        throw "UserForm template '$TemplatePath' has no 'properties' object."
    }
    if ($null -eq $properties.PSObject.Properties[$Name]) {
        throw "UserForm template '$TemplatePath' has no root '$Name' properties."
    }

    $assetPaths = [Collections.Generic.List[string]]::new()
    $assetSet = @{}
    $assetReferences = @(Get-IguanaTexFileAssetReferences -Value $template)

    foreach ($reference in $assetReferences) {
        $relativePath = $reference.Substring("file://".Length)
        $templateDirectory = Split-Path -Parent $TemplatePath
        $assetPath = Resolve-IguanaTexRelativePath `
            -Root $templateDirectory `
            -RelativePath $relativePath `
            -Description "file asset reference in $TemplatePath"

        if (-not (Test-IguanaTexPathWithinRoot -Path $assetPath -Root $AssetRoot)) {
            throw ("Asset reference for '{0}' escapes its owned directory '{1}': {2}" -f
                $Name,
                $AssetRoot,
                $reference)
        }
        if ($assetPath.Equals(
            [IO.Path]::GetFullPath($AssetRoot),
            [StringComparison]::OrdinalIgnoreCase
        )) {
            throw "Asset reference for '$Name' resolves to a directory: $reference"
        }
        if (-not (Test-Path -LiteralPath $assetPath -PathType Leaf)) {
            throw "Asset reference for '$Name' does not exist: $assetPath"
        }

        $key = $assetPath.ToLowerInvariant()
        if (-not $assetSet.ContainsKey($key)) {
            $assetSet[$key] = $true
            $assetPaths.Add($assetPath)
        }
    }

    if (Test-Path -LiteralPath $AssetRoot -PathType Container) {
        foreach ($assetFile in Get-ChildItem -LiteralPath $AssetRoot -Recurse -File) {
            if (-not $assetSet.ContainsKey($assetFile.FullName.ToLowerInvariant())) {
                throw "Unreferenced UserForm asset for '$Name': $($assetFile.FullName)"
            }
        }
    }

    return [PSCustomObject]@{
        Name = $Name
        TemplatePath = [IO.Path]::GetFullPath($TemplatePath)
        VbaPath = [IO.Path]::GetFullPath($VbaPath)
        AssetRoot = [IO.Path]::GetFullPath($AssetRoot)
        AssetPaths = @($assetPaths | Sort-Object)
    }
}

function Assert-IguanaTexUserFormManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [string]$ManifestPath
    )

    $resolvedProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
    if (-not (Test-Path -LiteralPath $resolvedProjectRoot -PathType Container)) {
        throw "Project root is not a directory: $ProjectRoot"
    }

    if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
        $ManifestPath = Join-Path `
            $resolvedProjectRoot `
            $script:UserFormManifestRelativePath
    }
    else {
        $ManifestPath = [IO.Path]::GetFullPath($ManifestPath)
    }

    if (-not (Test-IguanaTexPathWithinRoot `
        -Path $ManifestPath `
        -Root $resolvedProjectRoot
    )) {
        throw "UserForm manifest must be inside the project root: $ManifestPath"
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "UserForm manifest not found: $ManifestPath"
    }

    $manifest = Read-IguanaTexJsonFile `
        -Path $ManifestPath `
        -Description "UserForm manifest"
    $schemaVersion = Get-IguanaTexObjectPropertyValue `
        -Object $manifest `
        -Name "schemaVersion"
    if ($null -eq $schemaVersion -or [int]$schemaVersion -ne 1) {
        throw "UserForm manifest schemaVersion must be 1: $ManifestPath"
    }

    $rawForms = Get-IguanaTexObjectPropertyValue -Object $manifest -Name "forms"
    if ($null -eq $rawForms) {
        throw "UserForm manifest has no 'forms' array: $ManifestPath"
    }
    $rawForms = @($rawForms)
    if ($rawForms.Count -ne 9) {
        throw ("UserForm manifest must contain exactly 9 forms; found {0}: {1}" -f
            $rawForms.Count,
            $ManifestPath)
    }

    $sourceDirectory = Join-Path $resolvedProjectRoot "src"
    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        throw "Canonical VBA source directory not found: $sourceDirectory"
    }

    $formRecords = [Collections.Generic.List[object]]::new()
    $nameSet = @{}
    $templateSet = @{}
    $vbaSet = @{}
    $canonicalSet = @{}

    foreach ($rawForm in $rawForms) {
        $name = [string](Get-IguanaTexObjectPropertyValue `
            -Object $rawForm `
            -Name "name")
        $templateRelative = [string](Get-IguanaTexObjectPropertyValue `
            -Object $rawForm `
            -Name "template")
        $vbaRelative = [string](Get-IguanaTexObjectPropertyValue `
            -Object $rawForm `
            -Name "vba")

        if ($name -cnotmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
            throw "Invalid UserForm name in manifest: '$name'"
        }
        $nameKey = $name.ToLowerInvariant()
        if ($nameSet.ContainsKey($nameKey)) {
            throw "Duplicate UserForm name in manifest: $name"
        }
        $nameSet[$nameKey] = $true

        $templatePath = Resolve-IguanaTexRelativePath `
            -Root $resolvedProjectRoot `
            -RelativePath $templateRelative `
            -Description "template path for $name"
        $vbaPath = Resolve-IguanaTexRelativePath `
            -Root $resolvedProjectRoot `
            -RelativePath $vbaRelative `
            -Description "VBA path for $name"

        if ((Split-Path -Parent $templatePath) -cne $sourceDirectory) {
            throw "UserForm template must be a top-level file in src/: $templateRelative"
        }
        if ((Split-Path -Parent $vbaPath) -cne $sourceDirectory) {
            throw "UserForm VBA must be a top-level file in src/: $vbaRelative"
        }
        if ([IO.Path]::GetExtension($templatePath) -cne ".json") {
            throw "UserForm template must have a .json extension: $templateRelative"
        }
        if ([IO.Path]::GetExtension($vbaPath) -cne ".vba") {
            throw "UserForm VBA must have a .vba extension: $vbaRelative"
        }
        if ([IO.Path]::GetFileNameWithoutExtension($templatePath) -cne $name) {
            throw "UserForm template basename must match '$name': $templateRelative"
        }
        if ([IO.Path]::GetFileNameWithoutExtension($vbaPath) -cne $name) {
            throw "UserForm VBA basename must match '$name': $vbaRelative"
        }
        if ([IO.Path]::ChangeExtension($templatePath, ".vba") -cne $vbaPath) {
            throw "UserForm template and VBA must be same-directory sibling files: $name"
        }

        $templateKey = $templatePath.ToLowerInvariant()
        $vbaKey = $vbaPath.ToLowerInvariant()
        if ($templateSet.ContainsKey($templateKey)) {
            throw "Duplicate UserForm template path in manifest: $templateRelative"
        }
        if ($vbaSet.ContainsKey($vbaKey)) {
            throw "Duplicate UserForm VBA path in manifest: $vbaRelative"
        }
        $templateSet[$templateKey] = $true
        $vbaSet[$vbaKey] = $true

        $assetRoot = Join-Path $sourceDirectory $name
        $record = Assert-IguanaTexFormSource `
            -Name $name `
            -TemplatePath $templatePath `
            -VbaPath $vbaPath `
            -AssetRoot $assetRoot
        $formRecords.Add($record)

        $canonicalSet[$templateKey] = $true
        $canonicalSet[$vbaKey] = $true
        foreach ($assetPath in $record.AssetPaths) {
            $canonicalSet[$assetPath.ToLowerInvariant()] = $true
        }
    }

    foreach ($templateFile in Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -Filter "*.json" `
        -File
    ) {
        if (-not $templateSet.ContainsKey($templateFile.FullName.ToLowerInvariant())) {
            throw "UserForm template is not listed in the manifest: $($templateFile.FullName)"
        }
    }
    foreach ($vbaFile in Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -Filter "*.vba" `
        -File
    ) {
        if (-not $vbaSet.ContainsKey($vbaFile.FullName.ToLowerInvariant())) {
            throw "UserForm VBA file is not listed in the manifest: $($vbaFile.FullName)"
        }
    }

    $nativeSources = @(Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -Recurse `
        -File |
        Where-Object { $_.Extension -iin @(".frm", ".frx") })
    if ($nativeSources.Count -gt 0) {
        throw ("Generated .frm/.frx files are not canonical src/ inputs: {0}" -f
            (($nativeSources.FullName | Sort-Object) -join ", "))
    }

    foreach ($sourceFile in Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -Recurse `
        -File
    ) {
        $sourceKey = $sourceFile.FullName.ToLowerInvariant()
        if ($canonicalSet.ContainsKey($sourceKey)) {
            continue
        }
        if (
            $sourceFile.DirectoryName -ceq $sourceDirectory -and
            $sourceFile.Extension -iin @(".bas", ".cls")
        ) {
            continue
        }

        throw "Source file is outside the canonical manifest closure: $($sourceFile.FullName)"
    }

    $canonicalFiles = [Collections.Generic.List[string]]::new()
    $canonicalFiles.Add([IO.Path]::GetFullPath($ManifestPath))
    foreach ($record in $formRecords) {
        $canonicalFiles.Add($record.TemplatePath)
        $canonicalFiles.Add($record.VbaPath)
        foreach ($assetPath in $record.AssetPaths) {
            $canonicalFiles.Add($assetPath)
        }
    }

    return [PSCustomObject]@{
        ProjectRoot = $resolvedProjectRoot
        ManifestPath = [IO.Path]::GetFullPath($ManifestPath)
        SchemaVersion = 1
        SourceDirectory = $sourceDirectory
        Forms = @($formRecords)
        CanonicalFiles = @($canonicalFiles | Sort-Object -Unique)
    }
}

function Get-IguanaTexUserFormManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [string]$ManifestPath
    )

    return Assert-IguanaTexUserFormManifest `
        -ProjectRoot $ProjectRoot `
        -ManifestPath $ManifestPath
}

function Get-IguanaTexCanonicalInputHash {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [object]$Manifest
    )

    if ($null -eq $Manifest) {
        $Manifest = Assert-IguanaTexUserFormManifest -ProjectRoot $ProjectRoot
    }

    $files = [Collections.Generic.List[object]]::new()
    foreach ($path in @($Manifest.CanonicalFiles | Sort-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Canonical UserForm input disappeared while hashing: $path"
        }
        $relativePath = Get-IguanaTexRelativePath `
            -Path $path `
            -Root $Manifest.ProjectRoot
        $fileHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
        $files.Add([PSCustomObject]@{
            Path = $relativePath
            Hash = $fileHash
        })
    }

    $lines = @($files | Sort-Object Path | ForEach-Object {
        "{0}`t{1}" -f $_.Path, $_.Hash
    })
    $payload = ($lines -join "`n") + "`n"
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = $sha256.ComputeHash(
            ([Text.UTF8Encoding]::new($false)).GetBytes($payload)
        )
    }
    finally {
        $sha256.Dispose()
    }

    return [PSCustomObject]@{
        Algorithm = "SHA256"
        Hash = ([BitConverter]::ToString($digest) -replace "-", "").ToLowerInvariant()
        Files = @($files | Sort-Object Path)
    }
}

function Invoke-IguanaTexProcessCapture {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [string]$WorkingDirectory
    )

    $previousLocation = $null
    try {
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
            if (-not (Test-Path -LiteralPath $WorkingDirectory -PathType Container)) {
                throw "Process working directory not found: $WorkingDirectory"
            }
            $previousLocation = Get-Location
            Set-Location -LiteralPath $WorkingDirectory
        }

        $rawOutput = @(& $FilePath @Arguments 2>&1)
        $exitCode = $LASTEXITCODE
        $lines = @($rawOutput | ForEach-Object { [string]$_ })
        return [PSCustomObject]@{
            ExitCode = $exitCode
            Output = $lines
        }
    }
    finally {
        if ($null -ne $previousLocation) {
            Set-Location -LiteralPath $previousLocation.Path
        }
    }
}

function Write-IguanaTexProcessOutput {
    param([Parameter(Mandatory = $true)][object]$Result)

    foreach ($line in @($Result.Output)) {
        Write-Host $line
    }
}

function Invoke-IguanaTexGitCapture {
    param(
        [Parameter(Mandatory = $true)][string]$Repository,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )

    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($null -eq $git) {
        throw "Git is required to validate the FrxEdit submodule."
    }

    $result = Invoke-IguanaTexProcessCapture -FilePath $git.Source -Arguments (@("-C", $Repository) + $Arguments)
    if ($result.ExitCode -ne 0) {
        throw ("Git command failed (exit {0}): git -C {1} {2}{3}{4}" -f
            $result.ExitCode,
            $Repository,
            ($Arguments -join " "),
            [Environment]::NewLine,
            ($result.Output -join [Environment]::NewLine))
    }

    return @($result.Output)
}

function Assert-FrxEditCheckout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [ValidateSet("Pinned", "WorkingTree")]
        [string]$Mode = "Pinned"
    )

    $resolvedProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
    $submodulePath = Join-Path $resolvedProjectRoot $script:FrxEditRelativePath
    $projectPath = Join-Path $resolvedProjectRoot $script:FrxEditProjectRelativePath

    if (
        -not (Test-Path -LiteralPath $submodulePath -PathType Container) -or
        -not (Test-Path -LiteralPath (Join-Path $submodulePath ".git")) -or
        -not (Test-Path -LiteralPath $projectPath -PathType Leaf)
    ) {
        throw (
            "FrxEdit submodule is not initialized at '{0}'. Run " +
            "'git submodule update --init --recursive'." -f $submodulePath
        )
    }

    $gitlinkLines = @(Invoke-IguanaTexGitCapture -Repository $resolvedProjectRoot -Arguments @(
        "ls-files", "--stage", "--", "tools/frx-edit"
    ))
    if ($gitlinkLines.Count -ne 1) {
        throw ("Expected exactly one FrxEdit gitlink in the index; found {0}." -f
            $gitlinkLines.Count)
    }

    $gitlinkMatch = [regex]::Match(
        $gitlinkLines[0],
        '^160000\s+([0-9a-fA-F]{40})\s+0\s+tools/frx-edit$'
    )
    if (-not $gitlinkMatch.Success) {
        throw ("FrxEdit index entry must be a stage-0 mode-160000 gitlink: {0}" -f
            $gitlinkLines[0])
    }
    $gitlinkCommit = $gitlinkMatch.Groups[1].Value.ToLowerInvariant()

    $headLines = @(Invoke-IguanaTexGitCapture -Repository $submodulePath -Arguments @(
        "rev-parse", "HEAD"
    ))
    if ($headLines.Count -ne 1 -or $headLines[0] -cnotmatch '^[0-9a-fA-F]{40}$') {
        throw "Could not determine the FrxEdit submodule HEAD."
    }
    $headCommit = $headLines[0].ToLowerInvariant()

    $statusLines = @(Invoke-IguanaTexGitCapture -Repository $submodulePath -Arguments @(
        "status", "--porcelain=v1", "--untracked-files=all"
    ))
    $dirty = $statusLines.Count -gt 0

    if ($Mode -ceq "Pinned") {
        if ($headCommit -cne $gitlinkCommit) {
            throw (
                "FrxEdit submodule HEAD does not match the indexed gitlink. " +
                "Expected {0}, found {1}. Run " +
                "'git submodule update --init --recursive'." -f
                $gitlinkCommit,
                $headCommit
            )
        }
        if ($dirty) {
            throw (
                "FrxEdit submodule has working-tree changes. Pinned mode " +
                "accepts only a clean committed codec revision:{0}{1}" -f
                [Environment]::NewLine,
                ($statusLines -join [Environment]::NewLine)
            )
        }
    }

    return [PSCustomObject]@{
        Mode = $Mode
        SubmodulePath = $submodulePath
        ProjectPath = $projectPath
        GitlinkCommit = $gitlinkCommit
        HeadCommit = $headCommit
        Dirty = $dirty
        Status = $statusLines
    }
}

function Invoke-FrxEditCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExecutablePath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Arguments,

        [string]$WorkingDirectory
    )

    $resolvedExecutable = [IO.Path]::GetFullPath($ExecutablePath)
    if (-not (Test-Path -LiteralPath $resolvedExecutable -PathType Leaf)) {
        throw "FrxEdit executable not found: $resolvedExecutable"
    }

    $result = Invoke-IguanaTexProcessCapture -FilePath $resolvedExecutable -Arguments $Arguments -WorkingDirectory $WorkingDirectory
    Write-IguanaTexProcessOutput -Result $result

    if ($result.ExitCode -ne 0) {
        throw ("FrxEdit failed (exit {0}): {1} {2}" -f
            $result.ExitCode,
            $resolvedExecutable,
            ($Arguments -join " "))
    }
}

function Remove-IguanaTexDirectoryWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 50)][int]$Attempts = 10,
        [ValidateRange(0, 5000)][int]$DelayMilliseconds = 250,
        [switch]$WarnOnFailure
    )

    $lastError = $null
    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (-not (Test-Path -LiteralPath $Path)) {
            return $true
        }
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            return $true
        }
        catch {
            $lastError = $_
            if ($attempt -lt $Attempts -and $DelayMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
        }
    }

    $message = "Could not remove temporary directory after $Attempts attempts: $Path"
    if ($null -ne $lastError) {
        $message += " ($($lastError.Exception.Message))"
    }
    if ($WarnOnFailure) {
        Write-Warning $message
        return $false
    }
    throw $message
}

function Publish-IguanaTexDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$StagingDirectory,
        [Parameter(Mandatory = $true)][string]$DestinationDirectory
    )

    $destination = [IO.Path]::GetFullPath($DestinationDirectory)
    $parent = Split-Path -Parent $destination
    if ([string]::IsNullOrWhiteSpace($parent)) {
        throw "Destination directory has no parent: $destination"
    }
    [void](New-Item -ItemType Directory -Force -Path $parent)

    if (Test-Path -LiteralPath $destination -PathType Leaf) {
        throw "Destination is an existing file, not a directory: $destination"
    }

    $backup = Join-Path $parent (
        ([IO.Path]::GetFileName($destination)) +
        ".previous-" +
        [Guid]::NewGuid().ToString("N")
    )
    $movedExisting = $false
    try {
        if (Test-Path -LiteralPath $destination -PathType Container) {
            Move-Item -LiteralPath $destination -Destination $backup
            $movedExisting = $true
        }

        Move-Item -LiteralPath $StagingDirectory -Destination $destination
    }
    catch {
        $publishError = $_
        if (
            $movedExisting -and
            -not (Test-Path -LiteralPath $destination) -and
            (Test-Path -LiteralPath $backup -PathType Container)
        ) {
            Move-Item -LiteralPath $backup -Destination $destination
        }
        throw $publishError
    }

    if (Test-Path -LiteralPath $backup -PathType Container) {
        # The newly published destination is already authoritative. A running
        # FrxEdit process or antivirus scanner can briefly retain a handle to
        # the previous executable on Windows, so cleanup must not turn a
        # successful publish into a failure or roll back the new directory.
        [void](Remove-IguanaTexDirectoryWithRetry -Path $backup -WarnOnFailure)
    }
}

function Invoke-FrxEditBuild {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [ValidateSet("Pinned", "WorkingTree")]
        [string]$Mode = "Pinned",

        [string]$OutputDirectory
    )

    $resolvedProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = Join-Path $resolvedProjectRoot ".build\frxedit"
    }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)

    $checkoutBefore = Assert-FrxEditCheckout -ProjectRoot $resolvedProjectRoot -Mode $Mode
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnet) {
        throw ".NET SDK is required to publish FrxEdit."
    }

    $sdkResult = Invoke-IguanaTexProcessCapture -FilePath $dotnet.Source -Arguments @(
        "--version"
    )
    if ($sdkResult.ExitCode -ne 0 -or @($sdkResult.Output).Count -ne 1) {
        throw ("Could not determine the .NET SDK version:{0}{1}" -f
            [Environment]::NewLine,
            ($sdkResult.Output -join [Environment]::NewLine))
    }
    $sdkVersion = $sdkResult.Output[0].Trim()

    $outputParent = Split-Path -Parent $OutputDirectory
    [void](New-Item -ItemType Directory -Force -Path $outputParent)
    $stagingDirectory = Join-Path $outputParent (
        ".frxedit-publish-" + [Guid]::NewGuid().ToString("N")
    )
    [void](New-Item -ItemType Directory -Path $stagingDirectory)

    try {
        $publishArguments = @(
            "publish",
            $checkoutBefore.ProjectPath,
            "-c", "Release",
            "-r", "win-x64",
            "--self-contained", "true",
            "-p:PublishSingleFile=true",
            "-p:PublishTrimmed=false",
            "-o", $stagingDirectory
        )
        $publishResult = Invoke-IguanaTexProcessCapture -FilePath $dotnet.Source -Arguments $publishArguments -WorkingDirectory $resolvedProjectRoot
        Write-IguanaTexProcessOutput -Result $publishResult
        if ($publishResult.ExitCode -ne 0) {
            throw ("dotnet publish failed with exit code {0}." -f
                $publishResult.ExitCode)
        }

        $stagedExecutable = Join-Path $stagingDirectory "frxedit.exe"
        if (-not (Test-Path -LiteralPath $stagedExecutable -PathType Leaf)) {
            throw "FrxEdit publish did not produce frxedit.exe: $stagingDirectory"
        }

        $checkoutAfter = Assert-FrxEditCheckout -ProjectRoot $resolvedProjectRoot -Mode $Mode
        if ($checkoutAfter.HeadCommit -cne $checkoutBefore.HeadCommit) {
            throw "FrxEdit HEAD changed while it was being published."
        }
        if ($checkoutAfter.Dirty -ne $checkoutBefore.Dirty) {
            throw "FrxEdit working-tree state changed while it was being published."
        }

        $sha256 = (Get-FileHash -LiteralPath $stagedExecutable -Algorithm SHA256).Hash.ToLowerInvariant()
        $executablePath = Join-Path $OutputDirectory "frxedit.exe"
        $provenancePath = Join-Path $OutputDirectory "provenance.json"
        $provenance = [ordered]@{
            schemaVersion = 1
            mode = $Mode
            frxEditCommit = $checkoutAfter.HeadCommit
            gitlinkCommit = $checkoutAfter.GitlinkCommit
            workingTreeDirty = [bool]$checkoutAfter.Dirty
            dotNetSdkVersion = $sdkVersion
            configuration = "Release"
            runtimeIdentifier = "win-x64"
            selfContained = $true
            publishSingleFile = $true
            publishTrimmed = $false
            executable = Get-IguanaTexRelativePath -Path $executablePath -Root $resolvedProjectRoot
            executableSha256 = $sha256
            builtAtUtc = [DateTime]::UtcNow.ToString("o")
        }
        [IO.File]::WriteAllText(
            (Join-Path $stagingDirectory "provenance.json"),
            ($provenance | ConvertTo-Json -Depth 10),
            [Text.UTF8Encoding]::new($false)
        )

        Publish-IguanaTexDirectory -StagingDirectory $stagingDirectory -DestinationDirectory $OutputDirectory

        return [PSCustomObject]@{
            Mode = $Mode
            GitlinkCommit = $checkoutAfter.GitlinkCommit
            HeadCommit = $checkoutAfter.HeadCommit
            Dirty = [bool]$checkoutAfter.Dirty
            DotNetSdkVersion = $sdkVersion
            Sha256 = $sha256
            ExecutablePath = $executablePath
            ProvenancePath = $provenancePath
        }
    }
    finally {
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            [void](Remove-IguanaTexDirectoryWithRetry -Path $stagingDirectory -WarnOnFailure)
        }
    }
}

function Test-IguanaTexManagedOutputDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$ProjectRoot,
        [Parameter(Mandatory = $true)][string]$OutputDirectory
    )

    $buildRoot = Join-Path $ProjectRoot ".build"
    if (Test-IguanaTexPathWithinRoot -Path $OutputDirectory -Root $buildRoot) {
        return $true
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        return $true
    }

    return [bool](Test-Path -LiteralPath (
        Join-Path $OutputDirectory $script:GeneratedTreeMarker
    ) -PathType Leaf)
}

function Assert-IguanaTexGeneratedVba {
    param(
        [Parameter(Mandatory = $true)][string]$FormPath,
        [Parameter(Mandatory = $true)][string]$CanonicalVbaPath
    )

    $encoding = [Text.Encoding]::GetEncoding(1252)
    $formText = $encoding.GetString([IO.File]::ReadAllBytes($FormPath))
    $canonicalVba = $encoding.GetString(
        [IO.File]::ReadAllBytes($CanonicalVbaPath)
    )
    $attributeBlock = [regex]::Match(
        $formText,
        '^(Attribute\s+VB_[A-Za-z0-9_]+\s*=\s*.*?\r?\n)+',
        [Text.RegularExpressions.RegexOptions]::Multiline
    )
    if (-not $attributeBlock.Success) {
        throw "Generated UserForm has no VBA attribute boundary: $FormPath"
    }

    $generatedVba = $formText.Substring(
        $attributeBlock.Index + $attributeBlock.Length
    )

    # PowerPoint's VBE inserts one blank separator line after the exported
    # Attribute block. FrxEdit deliberately omits that container-only separator
    # when it writes the adjacent .vba payload. Permit exactly that one optional
    # leading separator; keep every byte of the actual VBA payload (including
    # line endings, casing, spaces, comments, and trailing blank lines)
    # significant.
    $vbaMatches = $generatedVba -ceq $canonicalVba
    if (-not $vbaMatches) {
        $separatorLength = 0
        if ($generatedVba.StartsWith("`r`n")) {
            $separatorLength = 2
        }
        elseif ($generatedVba.StartsWith("`r") -or $generatedVba.StartsWith("`n")) {
            $separatorLength = 1
        }
        if ($separatorLength -gt 0) {
            $vbaMatches = $generatedVba.Substring($separatorLength) -ceq $canonicalVba
        }
    }

    if (-not $vbaMatches) {
        throw (
            "Generated UserForm VBA does not match its canonical sibling " +
            "after allowed Attribute-boundary separator normalization: " +
            $CanonicalVbaPath
        )
    }
}

function New-IguanaTexUserFormStaging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FrxEditPath,

        [string]$OutputDirectory
    )

    $manifest = Assert-IguanaTexUserFormManifest -ProjectRoot $ProjectRoot
    $resolvedProjectRoot = $manifest.ProjectRoot
    $resolvedFrxEditPath = [IO.Path]::GetFullPath($FrxEditPath)
    if (-not (Test-Path -LiteralPath $resolvedFrxEditPath -PathType Leaf)) {
        throw "FrxEdit executable not found: $resolvedFrxEditPath"
    }

    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = Join-Path $resolvedProjectRoot ".build\vba-source"
    }
    $OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
    if ($OutputDirectory.Equals(
        $manifest.SourceDirectory,
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Generated UserForms must not be written into canonical src/: $OutputDirectory"
    }
    if (-not (Test-IguanaTexManagedOutputDirectory -ProjectRoot $resolvedProjectRoot -OutputDirectory $OutputDirectory)) {
        throw (
            "Refusing to replace an unmanaged output directory. Remove it, " +
            "choose a new path, or use a previously generated tree: $OutputDirectory"
        )
    }

    $hashBefore = Get-IguanaTexCanonicalInputHash -ProjectRoot $resolvedProjectRoot -Manifest $manifest
    $outputParent = Split-Path -Parent $OutputDirectory
    [void](New-Item -ItemType Directory -Force -Path $outputParent)
    $stagingDirectory = Join-Path $outputParent (
        ".vba-source-" + [Guid]::NewGuid().ToString("N")
    )
    [void](New-Item -ItemType Directory -Path $stagingDirectory)

    try {
        foreach ($component in Get-ChildItem -LiteralPath $manifest.SourceDirectory -File |
            Where-Object { $_.Extension -iin @(".bas", ".cls") } |
            Sort-Object Extension, Name
        ) {
            Copy-Item -LiteralPath $component.FullName -Destination $stagingDirectory
        }

        foreach ($form in $manifest.Forms) {
            $outputForm = Join-Path $stagingDirectory ($form.Name + ".frm")
            Invoke-FrxEditCommand -ExecutablePath $resolvedFrxEditPath -Arguments @(
                "create", $outputForm,
                "--name", $form.Name,
                "--patch", $form.TemplatePath
            ) -WorkingDirectory $resolvedProjectRoot
            Assert-IguanaTexGeneratedVba -FormPath $outputForm -CanonicalVbaPath $form.VbaPath
            Invoke-FrxEditCommand -ExecutablePath $resolvedFrxEditPath -Arguments @(
                "validate", $outputForm,
                "--mode", "strict"
            ) -WorkingDirectory $resolvedProjectRoot
        }

        $officeModulePath = Join-Path $resolvedProjectRoot "scripts\lib\IguanaTex.Office.psm1"
        if (-not (Test-Path -LiteralPath $officeModulePath -PathType Leaf)) {
            throw "Office source-closure module not found: $officeModulePath"
        }
        # Do not force-reload the Office module here. Callers such as
        # office-build.ps1 and vba-sync.ps1 import it into their own session
        # scope first; forcing a nested reload would remove those exported
        # commands from the caller after staging completes.
        Import-Module -Name $officeModulePath -DisableNameChecking -ErrorAction Stop
        Assert-VbaSourceClosure -SourceDirectory $stagingDirectory

        $manifestAfter = Assert-IguanaTexUserFormManifest -ProjectRoot $resolvedProjectRoot
        $hashAfter = Get-IguanaTexCanonicalInputHash -ProjectRoot $resolvedProjectRoot -Manifest $manifestAfter
        if ($hashAfter.Hash -cne $hashBefore.Hash) {
            throw "Canonical UserForm inputs changed while staging was generated."
        }

        $componentCount = @(Get-ChildItem -LiteralPath $stagingDirectory -File |
            Where-Object { $_.Extension -iin @(".bas", ".cls", ".frm") }).Count
        $marker = [ordered]@{
            schemaVersion = 1
            formCount = $manifest.Forms.Count
            componentCount = $componentCount
            canonicalInputSha256 = $hashBefore.Hash
            frxEditExecutableSha256 = (Get-FileHash -LiteralPath $resolvedFrxEditPath -Algorithm SHA256).Hash.ToLowerInvariant()
            generatedAtUtc = [DateTime]::UtcNow.ToString("o")
        }
        [IO.File]::WriteAllText(
            (Join-Path $stagingDirectory $script:GeneratedTreeMarker),
            ($marker | ConvertTo-Json -Depth 10),
            [Text.UTF8Encoding]::new($false)
        )

        Publish-IguanaTexDirectory -StagingDirectory $stagingDirectory -DestinationDirectory $OutputDirectory

        return [PSCustomObject]@{
            OutputDirectory = $OutputDirectory
            FormCount = $manifest.Forms.Count
            ComponentCount = $componentCount
            CanonicalInputSha256 = $hashBefore.Hash
        }
    }
    finally {
        if (Test-Path -LiteralPath $stagingDirectory -PathType Container) {
            [void](Remove-IguanaTexDirectoryWithRetry -Path $stagingDirectory -WarnOnFailure)
        }
    }
}

function Compare-IguanaTexUserFormSemantics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FrxEditPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$OriginalFormPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CandidateFormPath,

        [string]$ArtifactsDirectory,

        [string]$ReportPath
    )

    $resolvedProjectRoot = (Resolve-Path -LiteralPath $ProjectRoot).Path
    $resolvedExecutable = [IO.Path]::GetFullPath($FrxEditPath)
    $resolvedOriginal = [IO.Path]::GetFullPath($OriginalFormPath)
    $resolvedCandidate = [IO.Path]::GetFullPath($CandidateFormPath)
    foreach ($formPath in @($resolvedOriginal, $resolvedCandidate)) {
        if (-not (Test-Path -LiteralPath $formPath -PathType Leaf)) {
            throw "UserForm source for semantic comparison not found: $formPath"
        }
        $frxPath = [IO.Path]::ChangeExtension($formPath, ".frx")
        if (-not (Test-Path -LiteralPath $frxPath -PathType Leaf)) {
            throw "UserForm binary for semantic comparison not found: $frxPath"
        }
    }

    $comparatorPath = Join-Path $resolvedProjectRoot (
        "tools\frx-edit\scripts\compare-canonical-form.ps1"
    )
    if (-not (Test-Path -LiteralPath $comparatorPath -PathType Leaf)) {
        throw "Pinned FrxEdit semantic comparator not found: $comparatorPath"
    }

    $ownsArtifacts = [string]::IsNullOrWhiteSpace($ArtifactsDirectory)
    if ($ownsArtifacts) {
        $ArtifactsDirectory = Join-Path $resolvedProjectRoot (
            ".build\userforms-compare-" + [Guid]::NewGuid().ToString("N")
        )
    }
    $ArtifactsDirectory = [IO.Path]::GetFullPath($ArtifactsDirectory)
    [void](New-Item -ItemType Directory -Force -Path $ArtifactsDirectory)

    if ([string]::IsNullOrWhiteSpace($ReportPath)) {
        $ReportPath = Join-Path $ArtifactsDirectory "semantic-comparison.json"
    }
    $ReportPath = [IO.Path]::GetFullPath($ReportPath)
    $reportParent = Split-Path -Parent $ReportPath
    if (-not [string]::IsNullOrWhiteSpace($reportParent)) {
        [void](New-Item -ItemType Directory -Force -Path $reportParent)
    }

    try {
        $originalRaw = Join-Path $ArtifactsDirectory "original.raw.json"
        $candidateRaw = Join-Path $ArtifactsDirectory "candidate.raw.json"
        Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
            "inspect", $resolvedOriginal,
            "--mode", "strict",
            "--raw-out", $originalRaw,
            "--out", (Join-Path $ArtifactsDirectory "original.inspect.json")
        ) -WorkingDirectory $resolvedProjectRoot
        Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
            "inspect", $resolvedCandidate,
            "--mode", "strict",
            "--raw-out", $candidateRaw,
            "--out", (Join-Path $ArtifactsDirectory "candidate.inspect.json")
        ) -WorkingDirectory $resolvedProjectRoot

        & $comparatorPath -OriginalRaw $originalRaw -CandidateRaw $candidateRaw -ReportOut $ReportPath
        if (-not (Test-Path -LiteralPath $ReportPath -PathType Leaf)) {
            throw "Semantic comparator did not produce a report: $ReportPath"
        }
        $report = Read-IguanaTexJsonFile -Path $ReportPath -Description "semantic comparison report"
        if (-not [bool]$report.semanticEqual) {
            throw "UserForm semantic comparison failed: $ReportPath"
        }

        return $report
    }
    finally {
        if (
            $ownsArtifacts -and
            (Test-Path -LiteralPath $ArtifactsDirectory -PathType Container)
        ) {
            Remove-Item -LiteralPath $ArtifactsDirectory -Recurse -Force
        }
    }
}

function Resolve-IguanaTexExplicitFormSelection {
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][string[]]$Form
    )

    $requested = [Collections.Generic.List[string]]::new()
    $requestedSet = @{}
    foreach ($value in @($Form)) {
        if ($null -eq $value) {
            continue
        }
        foreach ($part in @($value -split ",")) {
            $name = $part.Trim()
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }
            if ($name -ieq "all") {
                throw (
                    "'all' is intentionally not accepted for canonical UserForm " +
                    "export. Name every intended form explicitly."
                )
            }
            $key = $name.ToLowerInvariant()
            if (-not $requestedSet.ContainsKey($key)) {
                $requestedSet[$key] = $true
                $requested.Add($name)
            }
        }
    }
    if ($requested.Count -eq 0) {
        throw "Specify at least one explicit UserForm name with -Form."
    }

    $selected = [Collections.Generic.List[object]]::new()
    foreach ($name in $requested) {
        $matches = @($Manifest.Forms | Where-Object { $_.Name -ieq $name })
        if ($matches.Count -ne 1) {
            $validNames = @($Manifest.Forms.Name | Sort-Object) -join ", "
            throw "Unknown UserForm '$name'. Manifest names: $validNames"
        }
        $selected.Add($matches[0])
    }

    return @($selected)
}

function Test-IguanaTexFilesEqual {
    param(
        [Parameter(Mandatory = $true)][string]$Left,
        [Parameter(Mandatory = $true)][string]$Right
    )

    if (
        -not (Test-Path -LiteralPath $Left -PathType Leaf) -or
        -not (Test-Path -LiteralPath $Right -PathType Leaf)
    ) {
        return $false
    }
    $leftInfo = Get-Item -LiteralPath $Left
    $rightInfo = Get-Item -LiteralPath $Right
    if ($leftInfo.Length -ne $rightInfo.Length) {
        return $false
    }
    return (Get-FileHash -LiteralPath $Left -Algorithm SHA256).Hash -ceq
        (Get-FileHash -LiteralPath $Right -Algorithm SHA256).Hash
}

function Get-IguanaTexDirectoryHashManifest {
    param([Parameter(Mandatory = $true)][string]$Directory)

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $Directory -Recurse -File |
        ForEach-Object {
            [PSCustomObject]@{
                Path = Get-IguanaTexRelativePath -Path $_.FullName -Root $Directory
                Hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        } |
        Sort-Object Path)
}

function Test-IguanaTexDirectoriesEqual {
    param(
        [Parameter(Mandatory = $true)][string]$Left,
        [Parameter(Mandatory = $true)][string]$Right
    )

    $leftManifest = @(Get-IguanaTexDirectoryHashManifest -Directory $Left)
    $rightManifest = @(Get-IguanaTexDirectoryHashManifest -Directory $Right)
    if ($leftManifest.Count -ne $rightManifest.Count) {
        return $false
    }
    for ($i = 0; $i -lt $leftManifest.Count; $i++) {
        if (
            $leftManifest[$i].Path -cne $rightManifest[$i].Path -or
            $leftManifest[$i].Hash -cne $rightManifest[$i].Hash
        ) {
            return $false
        }
    }
    return $true
}

function Export-IguanaTexCanonicalUserForms {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FrxEditPath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InputDirectory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Form,

        [switch]$DryRun
    )

    $manifest = Assert-IguanaTexUserFormManifest -ProjectRoot $ProjectRoot
    $resolvedProjectRoot = $manifest.ProjectRoot
    $resolvedExecutable = [IO.Path]::GetFullPath($FrxEditPath)
    if (-not (Test-Path -LiteralPath $resolvedExecutable -PathType Leaf)) {
        throw "FrxEdit executable not found: $resolvedExecutable"
    }
    $resolvedInput = (Resolve-Path -LiteralPath $InputDirectory).Path
    if (-not (Test-Path -LiteralPath $resolvedInput -PathType Container)) {
        throw "UserForm export input is not a directory: $InputDirectory"
    }

    $selectedForms = @(Resolve-IguanaTexExplicitFormSelection -Manifest $manifest -Form $Form)
    $canonicalHashBefore = Get-IguanaTexCanonicalInputHash -ProjectRoot $resolvedProjectRoot -Manifest $manifest
    $scratchRoot = Join-Path $resolvedProjectRoot (
        ".build\userforms-export-" + [Guid]::NewGuid().ToString("N")
    )
    [void](New-Item -ItemType Directory -Force -Path $scratchRoot)
    $candidates = [Collections.Generic.List[object]]::new()

    try {
        foreach ($formRecord in $selectedForms) {
            $inputForm = Join-Path $resolvedInput ($formRecord.Name + ".frm")
            $inputFrx = Join-Path $resolvedInput ($formRecord.Name + ".frx")
            if (-not (Test-Path -LiteralPath $inputForm -PathType Leaf)) {
                throw "Selected UserForm source not found: $inputForm"
            }
            if (-not (Test-Path -LiteralPath $inputFrx -PathType Leaf)) {
                throw "Selected UserForm binary not found: $inputFrx"
            }

            $caseRoot = Join-Path $scratchRoot $formRecord.Name
            $extractedRoot = Join-Path $caseRoot "extracted"
            $recreatedRoot = Join-Path $caseRoot "recreated"
            $reexportedRoot = Join-Path $caseRoot "reexported"
            [void](New-Item -ItemType Directory -Force -Path $extractedRoot)
            [void](New-Item -ItemType Directory -Force -Path $recreatedRoot)
            [void](New-Item -ItemType Directory -Force -Path $reexportedRoot)

            Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
                "validate", $inputForm,
                "--mode", "strict"
            ) -WorkingDirectory $resolvedProjectRoot

            $extractedTemplate = Join-Path $extractedRoot ($formRecord.Name + ".json")
            $extractedVba = Join-Path $extractedRoot ($formRecord.Name + ".vba")
            $extractedAssetRoot = Join-Path $extractedRoot $formRecord.Name
            Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
                "inspect", $inputForm,
                "--mode", "strict",
                "--as-template",
                "--extract-images",
                "--out", $extractedTemplate
            ) -WorkingDirectory $resolvedProjectRoot
            $extractedSource = Assert-IguanaTexFormSource -Name $formRecord.Name -TemplatePath $extractedTemplate -VbaPath $extractedVba -AssetRoot $extractedAssetRoot
            Assert-IguanaTexGeneratedVba `
                -FormPath $inputForm `
                -CanonicalVbaPath $extractedVba

            $recreatedForm = Join-Path $recreatedRoot ($formRecord.Name + ".frm")
            Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
                "create", $recreatedForm,
                "--name", $formRecord.Name,
                "--patch", $extractedTemplate
            ) -WorkingDirectory $resolvedProjectRoot
            Assert-IguanaTexGeneratedVba -FormPath $recreatedForm -CanonicalVbaPath $extractedVba
            Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
                "validate", $recreatedForm,
                "--mode", "strict"
            ) -WorkingDirectory $resolvedProjectRoot
            [void](Compare-IguanaTexUserFormSemantics -ProjectRoot $resolvedProjectRoot -FrxEditPath $resolvedExecutable -OriginalFormPath $inputForm -CandidateFormPath $recreatedForm -ArtifactsDirectory (Join-Path $caseRoot "comparison") -ReportPath (Join-Path $caseRoot "semantic-comparison.json"))

            $reexportedTemplate = Join-Path $reexportedRoot ($formRecord.Name + ".json")
            $reexportedVba = Join-Path $reexportedRoot ($formRecord.Name + ".vba")
            Invoke-FrxEditCommand -ExecutablePath $resolvedExecutable -Arguments @(
                "inspect", $recreatedForm,
                "--mode", "strict",
                "--as-template",
                "--extract-images",
                "--out", $reexportedTemplate
            ) -WorkingDirectory $resolvedProjectRoot
            if (-not (Test-IguanaTexFilesEqual -Left $extractedVba -Right $reexportedVba)) {
                throw "Exported VBA did not survive template recreation: $($formRecord.Name)"
            }

            $candidates.Add([PSCustomObject]@{
                Form = $formRecord
                TemplatePath = $extractedSource.TemplatePath
                VbaPath = $extractedSource.VbaPath
                AssetRoot = $extractedSource.AssetRoot
                TemplateChanged = -not (Test-IguanaTexFilesEqual -Left $extractedSource.TemplatePath -Right $formRecord.TemplatePath)
                VbaChanged = -not (Test-IguanaTexFilesEqual -Left $extractedSource.VbaPath -Right $formRecord.VbaPath)
                AssetsChanged = -not (Test-IguanaTexDirectoriesEqual -Left $extractedSource.AssetRoot -Right $formRecord.AssetRoot)
            })
        }

        $currentManifest = Assert-IguanaTexUserFormManifest -ProjectRoot $resolvedProjectRoot
        $currentHash = Get-IguanaTexCanonicalInputHash -ProjectRoot $resolvedProjectRoot -Manifest $currentManifest
        if ($currentHash.Hash -cne $canonicalHashBefore.Hash) {
            throw "Canonical UserForm inputs changed during export validation."
        }

        foreach ($candidate in $candidates) {
            $changes = [Collections.Generic.List[string]]::new()
            if ($candidate.TemplateChanged) { $changes.Add("JSON") }
            if ($candidate.VbaChanged) { $changes.Add("VBA") }
            if ($candidate.AssetsChanged) { $changes.Add("assets") }
            if ($changes.Count -eq 0) { $changes.Add("no byte changes") }
            $prefix = if ($DryRun) { "Would update" } else { "Validated update for" }
            Write-Host ("{0} {1}: {2}" -f
                $prefix,
                $candidate.Form.Name,
                ($changes -join ", "))
        }

        if (-not $DryRun) {
            $backupRoot = Join-Path $scratchRoot "canonical-backup"
            [void](New-Item -ItemType Directory -Force -Path $backupRoot)
            $operations = [Collections.Generic.List[object]]::new()
            foreach ($candidate in $candidates) {
                $formBackup = Join-Path $backupRoot $candidate.Form.Name
                [void](New-Item -ItemType Directory -Force -Path $formBackup)
                $operations.Add([PSCustomObject]@{
                    Source = $candidate.TemplatePath
                    Target = $candidate.Form.TemplatePath
                    Backup = Join-Path $formBackup "template.json"
                    OldMoved = $false
                    NewMoved = $false
                })
                $operations.Add([PSCustomObject]@{
                    Source = $candidate.VbaPath
                    Target = $candidate.Form.VbaPath
                    Backup = Join-Path $formBackup "code.vba"
                    OldMoved = $false
                    NewMoved = $false
                })
                $operations.Add([PSCustomObject]@{
                    Source = $candidate.AssetRoot
                    Target = $candidate.Form.AssetRoot
                    Backup = Join-Path $formBackup "assets"
                    OldMoved = $false
                    NewMoved = $false
                })
            }

            try {
                foreach ($operation in $operations) {
                    if (Test-Path -LiteralPath $operation.Target) {
                        Move-Item -LiteralPath $operation.Target -Destination $operation.Backup
                        $operation.OldMoved = $true
                    }
                    if (Test-Path -LiteralPath $operation.Source) {
                        $targetParent = Split-Path -Parent $operation.Target
                        [void](New-Item -ItemType Directory -Force -Path $targetParent)
                        Move-Item -LiteralPath $operation.Source -Destination $operation.Target
                        $operation.NewMoved = $true
                    }
                }

                [void](Assert-IguanaTexUserFormManifest -ProjectRoot $resolvedProjectRoot)
            }
            catch {
                $updateError = $_
                $rollbackErrors = [Collections.Generic.List[string]]::new()
                for ($i = $operations.Count - 1; $i -ge 0; $i--) {
                    $operation = $operations[$i]
                    try {
                        if ($operation.NewMoved -and (Test-Path -LiteralPath $operation.Target)) {
                            Remove-Item -LiteralPath $operation.Target -Recurse -Force
                        }
                        if ($operation.OldMoved -and (Test-Path -LiteralPath $operation.Backup)) {
                            Move-Item -LiteralPath $operation.Backup -Destination $operation.Target
                        }
                    }
                    catch {
                        $rollbackErrors.Add($_.Exception.Message)
                    }
                }
                if ($rollbackErrors.Count -gt 0) {
                    throw (
                        "Canonical UserForm update failed and rollback was incomplete. " +
                        "Update error: {0} Rollback errors: {1}" -f
                        $updateError.Exception.Message,
                        ($rollbackErrors -join " | ")
                    )
                }
                throw $updateError
            }
        }

        return [PSCustomObject]@{
            DryRun = [bool]$DryRun
            FormCount = $candidates.Count
            Forms = @($candidates | ForEach-Object { $_.Form.Name })
        }
    }
    finally {
        if (Test-Path -LiteralPath $scratchRoot -PathType Container) {
            Remove-Item -LiteralPath $scratchRoot -Recurse -Force
        }
    }
}

Export-ModuleMember -Function @(
    "Assert-FrxEditCheckout",
    "Assert-IguanaTexGeneratedVba",
    "Assert-IguanaTexUserFormManifest",
    "Compare-IguanaTexUserFormSemantics",
    "Export-IguanaTexCanonicalUserForms",
    "Get-IguanaTexCanonicalInputHash",
    "Get-IguanaTexUserFormManifest",
    "Invoke-FrxEditBuild",
    "Invoke-FrxEditCommand",
    "New-IguanaTexUserFormStaging"
)

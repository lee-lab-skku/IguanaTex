#requires -Version 5.1

Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$script:ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
$script:RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
$script:VbaRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
$script:VbaContentType = "application/vnd.ms-office.vbaProject"
$script:XmlContentType = "application/xml"
$script:RibbonDefinitions = @(
    [PSCustomObject]@{
        FileName = "customUI.xml"
        PartName = "customUI/customUI.xml"
        RelationshipType = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
        XmlNamespace = "http://schemas.microsoft.com/office/2006/01/customui"
    },
    [PSCustomObject]@{
        FileName = "customUI14.xml"
        PartName = "customUI/customUI14.xml"
        RelationshipType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
        XmlNamespace = "http://schemas.microsoft.com/office/2009/07/customui"
    }
)

function Get-IguanaTexResolvedFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Description not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).ProviderPath
}

function Get-IguanaTexResolvedDirectoryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "$Description not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).ProviderPath
}

function Invoke-IguanaTexZipArchive {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchiveMode]$Mode,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    $fileAccess = [System.IO.FileAccess]::Read

    if ($Mode -eq [System.IO.Compression.ZipArchiveMode]::Update) {
        $fileAccess = [System.IO.FileAccess]::ReadWrite
    }

    $fileStream = $null
    $archive = $null

    try {
        $fileStream = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            $fileAccess,
            [System.IO.FileShare]::Read
        )
        $archive = [System.IO.Compression.ZipArchive]::new(
            $fileStream,
            $Mode,
            $false
        )

        return (& $Action $archive)
    }
    finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }

        if ($null -ne $fileStream) {
            $fileStream.Dispose()
        }
    }
}

function Get-IguanaTexZipEntryMap {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive
    )

    $map = [System.Collections.Generic.Dictionary[
        string,
        System.IO.Compression.ZipArchiveEntry
    ]]::new([System.StringComparer]::Ordinal)

    foreach ($entry in $Archive.Entries) {
        $name = [string]$entry.FullName

        if ($name.EndsWith("/", [System.StringComparison]::Ordinal)) {
            continue
        }

        if ($name.StartsWith("/", [System.StringComparison]::Ordinal) -or
            $name.Contains("\")) {
            throw "Invalid OPC ZIP entry name: $name"
        }

        if ($map.ContainsKey($name)) {
            throw "Duplicate OPC ZIP entry: $name"
        }

        $map.Add($name, $entry)
    }

    return $map
}

function Get-IguanaTexRequiredZipEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.Dictionary[
            string,
            System.IO.Compression.ZipArchiveEntry
        ]]$EntryMap,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not $EntryMap.ContainsKey($Name)) {
        throw "Required OPC part is missing: $Name"
    }

    return $EntryMap[$Name]
}

function Get-IguanaTexZipEntryBytes {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchiveEntry]$Entry
    )

    $stream = $null
    $memory = $null

    try {
        $stream = $Entry.Open()
        $memory = [System.IO.MemoryStream]::new()
        $stream.CopyTo($memory)
        return ,$memory.ToArray()
    }
    finally {
        if ($null -ne $memory) {
            $memory.Dispose()
        }

        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Set-IguanaTexZipEntryBytes {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $existing = @(
        $Archive.Entries |
        Where-Object { [string]$_.FullName -ceq $Name }
    )

    if ($existing.Count -gt 1) {
        throw "Duplicate OPC ZIP entry: $Name"
    }

    $entry = $null
    $stream = $null

    try {
        if ($existing.Count -eq 1) {
            # Updating an existing entry stream preserves its central-directory
            # position. Deleting and recreating [Content_Types].xml or
            # _rels/.rels can produce a semantically valid PPAM that PowerPoint's
            # AddIns loader nevertheless refuses to load.
            $entry = $existing[0]
        }
        else {
            $entry = $Archive.CreateEntry(
                $Name,
                [System.IO.Compression.CompressionLevel]::Optimal
            )
        }

        $stream = $entry.Open()
        $stream.SetLength(0)
        $stream.Position = 0
        $stream.Write($Bytes, 0, $Bytes.Length)
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Assert-IguanaTexRequiredEntryOrder {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive
    )

    $entries = @($Archive.Entries)

    if ($entries.Count -lt 2 -or
        [string]$entries[0].FullName -cne "[Content_Types].xml" -or
        [string]$entries[1].FullName -cne "_rels/.rels") {
        throw (
            "Unexpected OPC ZIP entry order; [Content_Types].xml and " +
            "_rels/.rels must be the first two entries."
        )
    }
}

function Repair-IguanaTexRequiredEntryOrder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath
    )

    $requiresRepair = Invoke-IguanaTexZipArchive `
        -Path $PackagePath `
        -Mode ([System.IO.Compression.ZipArchiveMode]::Read) `
        -Action {
            param($archive)

            $entries = @($archive.Entries)

            return (
                $entries.Count -lt 2 -or
                [string]$entries[0].FullName -cne "[Content_Types].xml" -or
                [string]$entries[1].FullName -cne "_rels/.rels"
            )
        }

    if (-not $requiresRepair) {
        return $false
    }

    $directory = [System.IO.Path]::GetDirectoryName($PackagePath)
    $reorderedPath = Join-Path $directory (
        ".{0}.{1}.entry-order.tmp" -f
        [System.IO.Path]::GetFileName($PackagePath),
        [System.Guid]::NewGuid().ToString("N")
    )
    $sourceFile = $null
    $sourceArchive = $null
    $destinationFile = $null
    $destinationArchive = $null
    $reorderComplete = $false

    try {
        $sourceFile = [System.IO.File]::Open(
            $PackagePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read
        )
        $sourceArchive = [System.IO.Compression.ZipArchive]::new(
            $sourceFile,
            [System.IO.Compression.ZipArchiveMode]::Read,
            $false
        )
        $sourceEntries = @($sourceArchive.Entries)
        $contentTypesEntries = @(
            $sourceEntries |
            Where-Object { [string]$_.FullName -ceq "[Content_Types].xml" }
        )
        $rootRelationshipEntries = @(
            $sourceEntries |
            Where-Object { [string]$_.FullName -ceq "_rels/.rels" }
        )

        if ($contentTypesEntries.Count -ne 1 -or
            $rootRelationshipEntries.Count -ne 1) {
            throw "Cannot repair OPC entry order with missing or duplicate root entries."
        }

        $orderedEntries = @(
            $contentTypesEntries[0]
            $rootRelationshipEntries[0]
            $sourceEntries |
                Where-Object {
                    [string]$_.FullName -cne "[Content_Types].xml" -and
                    [string]$_.FullName -cne "_rels/.rels"
                }
        )

        $destinationFile = [System.IO.File]::Open(
            $reorderedPath,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $destinationArchive = [System.IO.Compression.ZipArchive]::new(
            $destinationFile,
            [System.IO.Compression.ZipArchiveMode]::Create,
            $false
        )

        foreach ($sourceEntry in $orderedEntries) {
            $destinationEntry = $destinationArchive.CreateEntry(
                [string]$sourceEntry.FullName,
                [System.IO.Compression.CompressionLevel]::Optimal
            )

            try {
                $destinationEntry.LastWriteTime = $sourceEntry.LastWriteTime
            }
            catch {
            }

            try {
                $destinationEntry.ExternalAttributes = $sourceEntry.ExternalAttributes
            }
            catch {
            }

            $inputStream = $null
            $outputStream = $null

            try {
                $inputStream = $sourceEntry.Open()
                $outputStream = $destinationEntry.Open()
                $inputStream.CopyTo($outputStream)
            }
            finally {
                if ($null -ne $outputStream) {
                    $outputStream.Dispose()
                }

                if ($null -ne $inputStream) {
                    $inputStream.Dispose()
                }
            }
        }

        $reorderComplete = $true
    }
    finally {
        if ($null -ne $destinationArchive) {
            $destinationArchive.Dispose()
        }

        if ($null -ne $destinationFile) {
            $destinationFile.Dispose()
        }

        if ($null -ne $sourceArchive) {
            $sourceArchive.Dispose()
        }

        if ($null -ne $sourceFile) {
            $sourceFile.Dispose()
        }

        if (-not $reorderComplete -and
            [System.IO.File]::Exists($reorderedPath)) {
            [System.IO.File]::Delete($reorderedPath)
        }
    }

    try {
        [System.IO.File]::Copy($reorderedPath, $PackagePath, $true)
    }
    finally {
        if ([System.IO.File]::Exists($reorderedPath)) {
            [System.IO.File]::Delete($reorderedPath)
        }
    }

    return $true
}

function ConvertFrom-IguanaTexXmlBytes {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $document = [System.Xml.XmlDocument]::new()
    $document.PreserveWhitespace = $true
    $document.XmlResolver = $null

    $settings = [System.Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.CloseInput = $false

    $memory = $null
    $reader = $null

    try {
        $memory = [System.IO.MemoryStream]::new($Bytes, $false)
        $reader = [System.Xml.XmlReader]::Create($memory, $settings)
        $document.Load($reader)
    }
    catch {
        throw "Invalid XML in $Description`: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }

        if ($null -ne $memory) {
            $memory.Dispose()
        }
    }

    return $document
}

function ConvertTo-IguanaTexXmlBytes {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document
    )

    $memory = $null

    try {
        $memory = [System.IO.MemoryStream]::new()
        $Document.Save($memory)
        return ,$memory.ToArray()
    }
    finally {
        if ($null -ne $memory) {
            $memory.Dispose()
        }
    }
}

function Get-IguanaTexXmlDocumentFromEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchiveEntry]$Entry
    )

    $bytes = Get-IguanaTexZipEntryBytes -Entry $Entry
    return ConvertFrom-IguanaTexXmlBytes `
        -Bytes $bytes `
        -Description ([string]$Entry.FullName)
}

function Assert-IguanaTexXmlRoot {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$LocalName,

        [Parameter(Mandatory = $true)]
        [string]$NamespaceUri,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    if ($null -eq $Document.DocumentElement -or
        $Document.DocumentElement.LocalName -cne $LocalName -or
        $Document.DocumentElement.NamespaceURI -cne $NamespaceUri) {
        throw (
            "Unexpected XML root in {0}; expected {{{1}}}{2}." -f
            $Description,
            $NamespaceUri,
            $LocalName
        )
    }
}

function Get-IguanaTexChildElements {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Parent,

        [Parameter(Mandatory = $true)]
        [string]$LocalName,

        [Parameter(Mandatory = $true)]
        [string]$NamespaceUri
    )

    return @(
        $Parent.ChildNodes |
        Where-Object {
            $_ -is [System.Xml.XmlElement] -and
            $_.LocalName -ceq $LocalName -and
            $_.NamespaceURI -ceq $NamespaceUri
        }
    )
}

function Resolve-IguanaTexInternalTarget {
    param(
        [AllowNull()]
        [string]$SourcePart,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    if ([string]::IsNullOrWhiteSpace($Target)) {
        throw "Internal OPC relationship has an empty target."
    }

    if ($Target.Contains("\")) {
        throw "Internal OPC relationship target contains a backslash: $Target"
    }

    $baseText = "http://iguana.invalid/"

    if (-not [string]::IsNullOrEmpty($SourcePart)) {
        $baseText += $SourcePart
    }

    try {
        $baseUri = [System.Uri]::new($baseText, [System.UriKind]::Absolute)
        $resolvedUri = [System.Uri]::new($baseUri, $Target)
    }
    catch {
        throw "Invalid internal OPC relationship target '$Target': $($_.Exception.Message)"
    }

    if ($resolvedUri.Scheme -cne $baseUri.Scheme -or
        $resolvedUri.Host -cne $baseUri.Host) {
        throw "Internal OPC relationship target is not package-relative: $Target"
    }

    $partName = [System.Uri]::UnescapeDataString(
        $resolvedUri.AbsolutePath.TrimStart("/")
    )

    if ([string]::IsNullOrEmpty($partName)) {
        throw "Internal OPC relationship resolves to the package root: $Target"
    }

    return $partName
}

function Get-IguanaTexRelationshipSourcePart {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelationshipPartName
    )

    if ($RelationshipPartName -ceq "_rels/.rels") {
        return $null
    }

    $marker = "/_rels/"
    $markerIndex = $RelationshipPartName.LastIndexOf(
        $marker,
        [System.StringComparison]::Ordinal
    )

    if ($markerIndex -lt 0 -or
        -not $RelationshipPartName.EndsWith(
            ".rels",
            [System.StringComparison]::Ordinal
        )) {
        throw "Invalid OPC relationship part name: $RelationshipPartName"
    }

    $directory = $RelationshipPartName.Substring(0, $markerIndex)
    $relatedName = $RelationshipPartName.Substring(
        $markerIndex + $marker.Length
    )
    $relatedName = $relatedName.Substring(0, $relatedName.Length - 5)

    if ([string]::IsNullOrEmpty($directory)) {
        return $relatedName
    }

    return "$directory/$relatedName"
}

function Get-IguanaTexRelationshipElements {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    Assert-IguanaTexXmlRoot `
        -Document $Document `
        -LocalName "Relationships" `
        -NamespaceUri $script:RelationshipsNamespace `
        -Description $Description

    return @(Get-IguanaTexChildElements `
        -Parent $Document.DocumentElement `
        -LocalName "Relationship" `
        -NamespaceUri $script:RelationshipsNamespace)
}

function Assert-IguanaTexRelationshipIds {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement[]]$Relationships,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $ids = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::Ordinal
    )

    foreach ($relationship in $Relationships) {
        $id = $relationship.GetAttribute("Id")

        if ([string]::IsNullOrWhiteSpace($id)) {
            throw "Relationship without Id in $Description."
        }

        if (-not $ids.Add($id)) {
            throw "Duplicate relationship Id '$id' in $Description."
        }

        if ([string]::IsNullOrWhiteSpace($relationship.GetAttribute("Type"))) {
            throw "Relationship '$id' has no Type in $Description."
        }

        if ([string]::IsNullOrWhiteSpace($relationship.GetAttribute("Target"))) {
            throw "Relationship '$id' has no Target in $Description."
        }
    }
}

function New-IguanaTexRelationshipId {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement[]]$Relationships
    )

    $ids = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::Ordinal
    )

    foreach ($relationship in $Relationships) {
        [void]$ids.Add($relationship.GetAttribute("Id"))
    }

    $number = 1

    while ($ids.Contains("rId$number")) {
        $number++
    }

    return "rId$number"
}

function Get-IguanaTexContentTypeElements {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$LocalName
    )

    Assert-IguanaTexXmlRoot `
        -Document $Document `
        -LocalName "Types" `
        -NamespaceUri $script:ContentTypesNamespace `
        -Description "[Content_Types].xml"

    return @(Get-IguanaTexChildElements `
        -Parent $Document.DocumentElement `
        -LocalName $LocalName `
        -NamespaceUri $script:ContentTypesNamespace)
}

function Get-IguanaTexEffectiveContentType {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$Document,

        [Parameter(Mandatory = $true)]
        [string]$PartName
    )

    $opcPartName = "/" + $PartName.TrimStart("/")
    $overrides = @(
        Get-IguanaTexContentTypeElements -Document $Document -LocalName "Override" |
        Where-Object { $_.GetAttribute("PartName") -ceq $opcPartName }
    )

    if ($overrides.Count -gt 1) {
        throw "Duplicate content-type Override for $opcPartName."
    }

    if ($overrides.Count -eq 1) {
        return $overrides[0].GetAttribute("ContentType")
    }

    $extension = [System.IO.Path]::GetExtension($PartName).TrimStart(".")
    $defaults = @(
        Get-IguanaTexContentTypeElements -Document $Document -LocalName "Default" |
        Where-Object {
            $_.GetAttribute("Extension").Equals(
                $extension,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        }
    )

    if ($defaults.Count -gt 1) {
        throw "Duplicate content-type Default for .$extension."
    }

    if ($defaults.Count -eq 0) {
        return $null
    }

    return $defaults[0].GetAttribute("ContentType")
}

function Get-IguanaTexExpectedMainContentType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath
    )

    $extension = [System.IO.Path]::GetExtension($PackagePath).ToLowerInvariant()

    switch ($extension) {
        ".pptm" {
            return "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
        }
        ".ppam" {
            return "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
        }
        default {
            throw (
                "Unsupported Office package extension '$extension'; " +
                "expected .pptm or .ppam: $PackagePath"
            )
        }
    }
}

function Test-IguanaTexByteArraysEqual {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Left,

        [Parameter(Mandatory = $true)]
        [byte[]]$Right
    )

    if ($Left.Length -ne $Right.Length) {
        return $false
    }

    for ($index = 0; $index -lt $Left.Length; $index++) {
        if ($Left[$index] -ne $Right[$index]) {
            return $false
        }
    }

    return $true
}

function Get-IguanaTexCanonicalRibbonInputs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RibbonDirectory
    )

    $resolvedDirectory = Get-IguanaTexResolvedDirectoryPath `
        -Path $RibbonDirectory `
        -Description "Ribbon source directory"
    $result = @()

    foreach ($definition in $script:RibbonDefinitions) {
        $path = Join-Path $resolvedDirectory $definition.FileName
        $resolvedPath = Get-IguanaTexResolvedFilePath `
            -Path $path `
            -Description "Canonical Ribbon XML"
        $bytes = [System.IO.File]::ReadAllBytes($resolvedPath)
        $document = ConvertFrom-IguanaTexXmlBytes `
            -Bytes $bytes `
            -Description $resolvedPath

        Assert-IguanaTexXmlRoot `
            -Document $document `
            -LocalName "customUI" `
            -NamespaceUri $definition.XmlNamespace `
            -Description $resolvedPath

        $result += [PSCustomObject]@{
            Definition = $definition
            Path = $resolvedPath
            Bytes = $bytes
            Document = $document
        }
    }

    return $result
}

function Get-IguanaTexRibbonCallbacks {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$RibbonInputs
    )

    $callbacks = [System.Collections.Generic.Dictionary[string,string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($input in $RibbonInputs) {
        $nodes = @($input.Document.SelectNodes("//*[@onAction]"))

        foreach ($node in $nodes) {
            $name = $node.GetAttribute("onAction")

            if ([string]::IsNullOrWhiteSpace($name)) {
                throw "Empty Ribbon onAction in $($input.Path)."
            }

            if ($name -cnotmatch "^[A-Za-z_][A-Za-z0-9_]*$") {
                throw "Unsupported Ribbon onAction callback name '$name' in $($input.Path)."
            }

            if (-not $callbacks.ContainsKey($name)) {
                $callbacks.Add($name, $name)
            }
        }
    }

    return @($callbacks.Values | Sort-Object)
}

function Assert-IguanaTexRibbonCallbacks {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Callbacks,

        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory
    )

    $resolvedSource = Get-IguanaTexResolvedDirectoryPath `
        -Path $SourceDirectory `
        -Description "Canonical VBA source directory"
    $modules = @(Get-ChildItem -LiteralPath $resolvedSource -Filter "*.bas" -File)

    if ($modules.Count -eq 0) {
        throw "No standard VBA modules found in: $resolvedSource"
    }

    $publicProcedures = [System.Collections.Generic.Dictionary[
        string,
        System.Collections.Generic.List[string]
    ]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $declarationPattern = [regex]::new(
        "(?im)^[\t ]*(?:(Public|Private|Friend)[\t ]+)?(?:(?:Static)[\t ]+)?(Sub|Function)[\t ]+([A-Za-z_][A-Za-z0-9_]*)[\t ]*\(",
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    )
    $optionPrivatePattern = [regex]::new(
        "(?im)^[\t ]*Option[\t ]+Private[\t ]+Module(?:[\t ]*(?:'.*)?)?$",
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
    )

    foreach ($module in $modules) {
        [string]$text = Get-Content -LiteralPath $module.FullName -Raw

        if ($optionPrivatePattern.IsMatch($text)) {
            continue
        }

        foreach ($match in $declarationPattern.Matches($text)) {
            $access = $match.Groups[1].Value

            if ($access.Equals("Private", [System.StringComparison]::OrdinalIgnoreCase) -or
                $access.Equals("Friend", [System.StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $name = $match.Groups[3].Value

            if (-not $publicProcedures.ContainsKey($name)) {
                $publicProcedures.Add(
                    $name,
                    [System.Collections.Generic.List[string]]::new()
                )
            }

            $publicProcedures[$name].Add($module.Name)
        }
    }

    foreach ($callback in $Callbacks) {
        if (-not $publicProcedures.ContainsKey($callback)) {
            throw (
                "Ribbon callback '$callback' is not a public/default-public " +
                "Sub or Function in a standard .bas module."
            )
        }

        $locations = @($publicProcedures[$callback])

        if ($locations.Count -ne 1) {
            throw (
                "Ribbon callback '$callback' is ambiguous across standard modules: " +
                ($locations -join ", ")
            )
        }
    }
}

function Assert-IguanaTexRelationshipClosureInArchive {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.Dictionary[
            string,
            System.IO.Compression.ZipArchiveEntry
        ]]$EntryMap
    )

    $internalCount = 0
    $externalCount = 0
    $relationshipPartCount = 0

    foreach ($entry in $Archive.Entries) {
        $relationshipPartName = [string]$entry.FullName

        if (-not $relationshipPartName.EndsWith(
            ".rels",
            [System.StringComparison]::Ordinal
        )) {
            continue
        }

        $relationshipPartCount++
        $sourcePart = Get-IguanaTexRelationshipSourcePart `
            -RelationshipPartName $relationshipPartName

        if ($null -ne $sourcePart -and -not $EntryMap.ContainsKey($sourcePart)) {
            throw (
                "Relationship part '$relationshipPartName' has a missing source part: " +
                $sourcePart
            )
        }

        $document = Get-IguanaTexXmlDocumentFromEntry -Entry $entry
        $relationships = @(Get-IguanaTexRelationshipElements `
            -Document $document `
            -Description $relationshipPartName)
        Assert-IguanaTexRelationshipIds `
            -Relationships $relationships `
            -Description $relationshipPartName
        $logicalRelationships = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::Ordinal
        )
        $logicalSeparator = [char]31

        foreach ($relationship in $relationships) {
            $mode = $relationship.GetAttribute("TargetMode")
            $type = $relationship.GetAttribute("Type")
            $logicalTarget = $null
            $logicalMode = $null

            if ([string]::IsNullOrEmpty($mode) -or $mode -ceq "Internal") {
                $resolvedTarget = Resolve-IguanaTexInternalTarget `
                    -SourcePart $sourcePart `
                    -Target $relationship.GetAttribute("Target")
                $logicalMode = "Internal"
                $logicalTarget = $resolvedTarget

                if (-not $EntryMap.ContainsKey($resolvedTarget)) {
                    throw (
                        "Missing internal relationship target '$resolvedTarget' " +
                        "from '$relationshipPartName' (Id " +
                        $relationship.GetAttribute("Id") + ")."
                    )
                }

                $internalCount++
            }
            elseif ($mode -ceq "External") {
                $logicalMode = "External"
                $logicalTarget = $relationship.GetAttribute("Target")
                $externalCount++
            }
            else {
                throw (
                    "Invalid TargetMode '$mode' in '$relationshipPartName' (Id " +
                    $relationship.GetAttribute("Id") + ")."
                )
            }

            $logicalKey = (
                $logicalMode + $logicalSeparator +
                $type + $logicalSeparator +
                $logicalTarget
            )

            if (-not $logicalRelationships.Add($logicalKey)) {
                throw (
                    "Duplicate logical relationship in '$relationshipPartName': " +
                    "Type '$type', Target '$logicalTarget', TargetMode '$logicalMode'."
                )
            }
        }
    }

    return [PSCustomObject]@{
        RelationshipParts = $relationshipPartCount
        InternalRelationships = $internalCount
        ExternalRelationships = $externalCount
    }
}

function Assert-IguanaTexRootRibbonRelationships {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlDocument]$RootRelationshipsDocument
    )

    $relationships = @(Get-IguanaTexRelationshipElements `
        -Document $RootRelationshipsDocument `
        -Description "_rels/.rels")
    Assert-IguanaTexRelationshipIds `
        -Relationships $relationships `
        -Description "_rels/.rels"

    foreach ($definition in $script:RibbonDefinitions) {
        $typed = @(
            $relationships |
            Where-Object {
                $_.GetAttribute("Type") -ceq $definition.RelationshipType
            }
        )

        if ($typed.Count -ne 1) {
            throw (
                "Expected exactly one root Ribbon relationship of type " +
                "'$($definition.RelationshipType)', found $($typed.Count)."
            )
        }

        $mode = $typed[0].GetAttribute("TargetMode")

        if (-not [string]::IsNullOrEmpty($mode) -and $mode -cne "Internal") {
            throw "Ribbon relationship must be internal: $($definition.RelationshipType)"
        }

        $resolvedTarget = Resolve-IguanaTexInternalTarget `
            -SourcePart $null `
            -Target $typed[0].GetAttribute("Target")

        if ($resolvedTarget -cne $definition.PartName) {
            throw (
                "Ribbon relationship '$($definition.RelationshipType)' resolves to " +
                "'$resolvedTarget', expected '$($definition.PartName)'."
            )
        }
    }
}

function Assert-IguanaTexRibbonState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath,

        [Parameter(Mandatory = $true)]
        [object[]]$RibbonInputs
    )

    $expectedMainContentType = Get-IguanaTexExpectedMainContentType `
        -PackagePath $PackagePath

    return Invoke-IguanaTexZipArchive `
        -Path $PackagePath `
        -Mode ([System.IO.Compression.ZipArchiveMode]::Read) `
        -Action {
            param($archive)

            Assert-IguanaTexRequiredEntryOrder -Archive $archive
            $entryMap = Get-IguanaTexZipEntryMap -Archive $archive
            $contentTypesEntry = Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "[Content_Types].xml"
            $contentTypes = Get-IguanaTexXmlDocumentFromEntry `
                -Entry $contentTypesEntry
            $xmlDefaults = @(
                Get-IguanaTexContentTypeElements `
                    -Document $contentTypes `
                    -LocalName "Default" |
                Where-Object {
                    $_.GetAttribute("Extension").Equals(
                        "xml",
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
            )

            if ($xmlDefaults.Count -ne 1 -or
                $xmlDefaults[0].GetAttribute("ContentType") -cne
                    $script:XmlContentType) {
                throw (
                    "Expected exactly one .xml Default content type of '" +
                    $script:XmlContentType + "'."
                )
            }

            [void](Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "ppt/presentation.xml")
            $mainContentType = Get-IguanaTexEffectiveContentType `
                -Document $contentTypes `
                -PartName "ppt/presentation.xml"

            if ($mainContentType -cne $expectedMainContentType) {
                throw (
                    "PowerPoint main part content type is '$mainContentType', " +
                    "expected '$expectedMainContentType'."
                )
            }

            $rootRelationshipsEntry = Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "_rels/.rels"
            $rootRelationships = Get-IguanaTexXmlDocumentFromEntry `
                -Entry $rootRelationshipsEntry

            foreach ($input in $RibbonInputs) {
                $partName = $input.Definition.PartName
                $entry = Get-IguanaTexRequiredZipEntry `
                    -EntryMap $entryMap `
                    -Name $partName
                $actualBytes = Get-IguanaTexZipEntryBytes -Entry $entry

                if (-not (Test-IguanaTexByteArraysEqual `
                    -Left $actualBytes `
                    -Right $input.Bytes)) {
                    throw "Ribbon part does not match canonical bytes: $partName"
                }

                $actualDocument = ConvertFrom-IguanaTexXmlBytes `
                    -Bytes $actualBytes `
                    -Description $partName
                Assert-IguanaTexXmlRoot `
                    -Document $actualDocument `
                    -LocalName "customUI" `
                    -NamespaceUri $input.Definition.XmlNamespace `
                    -Description $partName

                $contentType = Get-IguanaTexEffectiveContentType `
                    -Document $contentTypes `
                    -PartName $partName

                if ($contentType -cne $script:XmlContentType) {
                    throw (
                        "Ribbon part '$partName' has content type '$contentType'; " +
                        "expected '$($script:XmlContentType)'."
                    )
                }
            }

            Assert-IguanaTexRootRibbonRelationships `
                -RootRelationshipsDocument $rootRelationships

            $closure = Assert-IguanaTexRelationshipClosureInArchive `
                -Archive $archive `
                -EntryMap $entryMap

            return [PSCustomObject]@{
                MainContentType = $mainContentType
                RelationshipParts = $closure.RelationshipParts
                InternalRelationships = $closure.InternalRelationships
                ExternalRelationships = $closure.ExternalRelationships
            }
        }
}

function Add-IguanaTexRibbon {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath,

        [Parameter(Mandatory = $true)]
        [string]$RibbonDirectory
    )

    $resolvedPackage = Get-IguanaTexResolvedFilePath `
        -Path $PackagePath `
        -Description "Office package"
    [void](Get-IguanaTexExpectedMainContentType -PackagePath $resolvedPackage)
    $ribbonInputs = @(Get-IguanaTexCanonicalRibbonInputs `
        -RibbonDirectory $RibbonDirectory)
    $existingState = $null

    try {
        $existingState = Assert-IguanaTexRibbonState `
            -PackagePath $resolvedPackage `
            -RibbonInputs $ribbonInputs
    }
    catch {
        # A non-canonical Ribbon state proceeds through the transactional
        # mutation path below. That path performs the same full assertion
        # before publishing its temporary package over the caller's artifact.
        $existingState = $null
    }

    if ($null -ne $existingState) {
        return [PSCustomObject]@{
            PSTypeName = "IguanaTex.RibbonInjectionResult"
            PackagePath = $resolvedPackage
            RibbonPartsWritten = 0
            RibbonPartsReused = $ribbonInputs.Count
            RelationshipsAdded = 0
            RelationshipsReused = $script:RibbonDefinitions.Count
            XmlDefaultAdded = $false
            EntryOrderRepaired = $false
            PackageChanged = $false
            InternalRelationships = $existingState.InternalRelationships
        }
    }

    $packageDirectory = [System.IO.Path]::GetDirectoryName($resolvedPackage)
    $temporaryPath = Join-Path $packageDirectory (
        ".{0}.{1}.ribbon.tmp{2}" -f
        [System.IO.Path]::GetFileNameWithoutExtension($resolvedPackage),
        [System.Guid]::NewGuid().ToString("N"),
        [System.IO.Path]::GetExtension($resolvedPackage)
    )
    [System.IO.File]::Copy($resolvedPackage, $temporaryPath, $false)

    try {
        $editSummary = Invoke-IguanaTexZipArchive `
            -Path $temporaryPath `
            -Mode ([System.IO.Compression.ZipArchiveMode]::Update) `
            -Action {
                param($archive)

                $localRelationshipAddedCount = 0
                $localRelationshipReusedCount = 0
                $localXmlDefaultAdded = $false
                $entryMap = Get-IguanaTexZipEntryMap -Archive $archive
                $contentTypesEntry = Get-IguanaTexRequiredZipEntry `
                    -EntryMap $entryMap `
                    -Name "[Content_Types].xml"
                $contentTypes = Get-IguanaTexXmlDocumentFromEntry `
                    -Entry $contentTypesEntry
                $rootRelationshipsEntry = Get-IguanaTexRequiredZipEntry `
                    -EntryMap $entryMap `
                    -Name "_rels/.rels"
                $rootRelationships = Get-IguanaTexXmlDocumentFromEntry `
                    -Entry $rootRelationshipsEntry
                $relationships = @(Get-IguanaTexRelationshipElements `
                    -Document $rootRelationships `
                    -Description "_rels/.rels")
                Assert-IguanaTexRelationshipIds `
                    -Relationships $relationships `
                    -Description "_rels/.rels"

                $xmlDefaults = @(
                    Get-IguanaTexContentTypeElements `
                        -Document $contentTypes `
                        -LocalName "Default" |
                    Where-Object {
                        $_.GetAttribute("Extension").Equals(
                            "xml",
                            [System.StringComparison]::OrdinalIgnoreCase
                        )
                    }
                )

                if ($xmlDefaults.Count -gt 1) {
                    throw "Duplicate content-type Default for .xml."
                }

                if ($xmlDefaults.Count -eq 1) {
                    if ($xmlDefaults[0].GetAttribute("ContentType") -cne
                        $script:XmlContentType) {
                        throw (
                            "Existing .xml content type is '" +
                            $xmlDefaults[0].GetAttribute("ContentType") +
                            "', expected '$($script:XmlContentType)'."
                        )
                    }
                }
                else {
                    $newDefault = $contentTypes.CreateElement(
                        "Default",
                        $script:ContentTypesNamespace
                    )
                    $newDefault.SetAttribute("Extension", "xml")
                    $newDefault.SetAttribute("ContentType", $script:XmlContentType)
                    [void]$contentTypes.DocumentElement.AppendChild($newDefault)
                    $localXmlDefaultAdded = $true
                }

                foreach ($input in $ribbonInputs) {
                    $definition = $input.Definition
                    $typed = @(
                        $relationships |
                        Where-Object {
                            $_.GetAttribute("Type") -ceq
                            $definition.RelationshipType
                        }
                    )

                    if ($typed.Count -gt 0) {
                        foreach ($relationship in $typed) {
                            $mode = $relationship.GetAttribute("TargetMode")

                            if (-not [string]::IsNullOrEmpty($mode) -and
                                $mode -cne "Internal") {
                                throw (
                                    "Conflicting external Ribbon relationship type: " +
                                    $definition.RelationshipType
                                )
                            }

                            $resolvedTarget = Resolve-IguanaTexInternalTarget `
                                -SourcePart $null `
                                -Target $relationship.GetAttribute("Target")

                            if ($resolvedTarget -cne $definition.PartName) {
                                throw (
                                    "Conflicting Ribbon relationship type " +
                                    "'$($definition.RelationshipType)' targets " +
                                    "'$resolvedTarget'."
                                )
                            }
                        }

                        if ($typed.Count -gt 1) {
                            for ($index = 1; $index -lt $typed.Count; $index++) {
                                [void]$typed[$index].ParentNode.RemoveChild($typed[$index])
                            }
                        }

                        $localRelationshipReusedCount++
                    }
                    else {
                        $relationshipId = New-IguanaTexRelationshipId `
                            -Relationships $relationships
                        $newRelationship = $rootRelationships.CreateElement(
                            "Relationship",
                            $script:RelationshipsNamespace
                        )
                        $newRelationship.SetAttribute("Id", $relationshipId)
                        $newRelationship.SetAttribute(
                            "Type",
                            $definition.RelationshipType
                        )
                        $newRelationship.SetAttribute("Target", $definition.PartName)
                        [void]$rootRelationships.DocumentElement.AppendChild(
                            $newRelationship
                        )
                        $relationships += $newRelationship
                        $localRelationshipAddedCount++
                    }

                    Set-IguanaTexZipEntryBytes `
                        -Archive $archive `
                        -Name $definition.PartName `
                        -Bytes $input.Bytes
                }

                Set-IguanaTexZipEntryBytes `
                    -Archive $archive `
                    -Name "_rels/.rels" `
                    -Bytes (ConvertTo-IguanaTexXmlBytes `
                        -Document $rootRelationships)
                Set-IguanaTexZipEntryBytes `
                    -Archive $archive `
                    -Name "[Content_Types].xml" `
                    -Bytes (ConvertTo-IguanaTexXmlBytes `
                        -Document $contentTypes)

                return [PSCustomObject]@{
                    RelationshipsAdded = $localRelationshipAddedCount
                    RelationshipsReused = $localRelationshipReusedCount
                    XmlDefaultAdded = $localXmlDefaultAdded
                }
            }

        $entryOrderRepaired = Repair-IguanaTexRequiredEntryOrder `
            -PackagePath $temporaryPath
        $relationshipSummary = Assert-IguanaTexRibbonState `
            -PackagePath $temporaryPath `
            -RibbonInputs $ribbonInputs

        [System.IO.File]::Copy($temporaryPath, $resolvedPackage, $true)

        return [PSCustomObject]@{
            PSTypeName = "IguanaTex.RibbonInjectionResult"
            PackagePath = $resolvedPackage
            RibbonPartsWritten = $ribbonInputs.Count
            RibbonPartsReused = 0
            RelationshipsAdded = $editSummary.RelationshipsAdded
            RelationshipsReused = $editSummary.RelationshipsReused
            XmlDefaultAdded = $editSummary.XmlDefaultAdded
            EntryOrderRepaired = $entryOrderRepaired
            PackageChanged = $true
            InternalRelationships = $relationshipSummary.InternalRelationships
        }
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) {
            [System.IO.File]::Delete($temporaryPath)
        }
    }
}

function Assert-VbaPackageClosure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath
    )

    $resolvedPackage = Get-IguanaTexResolvedFilePath `
        -Path $PackagePath `
        -Description "Office package"
    $expectedMainContentType = Get-IguanaTexExpectedMainContentType `
        -PackagePath $resolvedPackage

    return Invoke-IguanaTexZipArchive `
        -Path $resolvedPackage `
        -Mode ([System.IO.Compression.ZipArchiveMode]::Read) `
        -Action {
            param($archive)

            $entryMap = Get-IguanaTexZipEntryMap -Archive $archive
            $vbaEntry = Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "ppt/vbaProject.bin"

            if ($vbaEntry.Length -le 0) {
                throw "VBA project part is empty: ppt/vbaProject.bin"
            }

            $presentationRelationshipsEntry = Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "ppt/_rels/presentation.xml.rels"
            $presentationRelationships = Get-IguanaTexXmlDocumentFromEntry `
                -Entry $presentationRelationshipsEntry
            $relationships = @(Get-IguanaTexRelationshipElements `
                -Document $presentationRelationships `
                -Description "ppt/_rels/presentation.xml.rels")
            Assert-IguanaTexRelationshipIds `
                -Relationships $relationships `
                -Description "ppt/_rels/presentation.xml.rels"
            $vbaRelationships = @(
                $relationships |
                Where-Object {
                    $_.GetAttribute("Type") -ceq $script:VbaRelationshipType
                }
            )

            if ($vbaRelationships.Count -ne 1) {
                throw (
                    "Expected exactly one VBA project relationship, found " +
                    $vbaRelationships.Count + "."
                )
            }

            $mode = $vbaRelationships[0].GetAttribute("TargetMode")

            if (-not [string]::IsNullOrEmpty($mode) -and $mode -cne "Internal") {
                throw "VBA project relationship must be internal."
            }

            $resolvedTarget = Resolve-IguanaTexInternalTarget `
                -SourcePart "ppt/presentation.xml" `
                -Target $vbaRelationships[0].GetAttribute("Target")

            if ($resolvedTarget -cne "ppt/vbaProject.bin") {
                throw (
                    "VBA project relationship resolves to '$resolvedTarget', " +
                    "expected 'ppt/vbaProject.bin'."
                )
            }

            $contentTypesEntry = Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "[Content_Types].xml"
            $contentTypes = Get-IguanaTexXmlDocumentFromEntry `
                -Entry $contentTypesEntry
            [void](Get-IguanaTexRequiredZipEntry `
                -EntryMap $entryMap `
                -Name "ppt/presentation.xml")
            $mainContentType = Get-IguanaTexEffectiveContentType `
                -Document $contentTypes `
                -PartName "ppt/presentation.xml"

            if ($mainContentType -cne $expectedMainContentType) {
                throw (
                    "PowerPoint main part content type is '$mainContentType', " +
                    "expected '$expectedMainContentType'."
                )
            }

            $binDefaults = @(
                Get-IguanaTexContentTypeElements `
                    -Document $contentTypes `
                    -LocalName "Default" |
                Where-Object {
                    $_.GetAttribute("Extension").Equals(
                        "bin",
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
            )

            if ($binDefaults.Count -ne 1) {
                throw (
                    "Expected exactly one .bin content-type Default, found " +
                    $binDefaults.Count + "."
                )
            }

            $binDefaultContentType = $binDefaults[0].GetAttribute("ContentType")

            if ($binDefaultContentType -cne $script:VbaContentType) {
                throw (
                    ".bin content-type Default is '$binDefaultContentType', " +
                    "expected '$($script:VbaContentType)'."
                )
            }

            $vbaContentType = Get-IguanaTexEffectiveContentType `
                -Document $contentTypes `
                -PartName "ppt/vbaProject.bin"

            if ($vbaContentType -cne $script:VbaContentType) {
                throw (
                    "VBA project content type is '$vbaContentType', expected " +
                    "'$($script:VbaContentType)'."
                )
            }

            return [PSCustomObject]@{
                PSTypeName = "IguanaTex.VbaPackageClosureResult"
                PackagePath = $resolvedPackage
                VbaProjectPart = "ppt/vbaProject.bin"
                VbaProjectBytes = $vbaEntry.Length
                RelationshipId = $vbaRelationships[0].GetAttribute("Id")
                MainContentType = $mainContentType
                ContentType = $vbaContentType
                BinDefaultContentType = $binDefaultContentType
            }
        }
}

function Assert-IguanaTexOfficePackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackagePath,

        [Parameter(Mandatory = $true)]
        [string]$RibbonDirectory,

        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory
    )

    $resolvedPackage = Get-IguanaTexResolvedFilePath `
        -Path $PackagePath `
        -Description "Office package"
    $ribbonInputs = @(Get-IguanaTexCanonicalRibbonInputs `
        -RibbonDirectory $RibbonDirectory)
    $callbacks = @(Get-IguanaTexRibbonCallbacks -RibbonInputs $ribbonInputs)

    Assert-IguanaTexRibbonCallbacks `
        -Callbacks $callbacks `
        -SourceDirectory $SourceDirectory

    $vbaSummary = Assert-VbaPackageClosure -PackagePath $resolvedPackage
    $relationshipSummary = Assert-IguanaTexRibbonState `
        -PackagePath $resolvedPackage `
        -RibbonInputs $ribbonInputs

    return [PSCustomObject]@{
        PSTypeName = "IguanaTex.OfficePackageValidationResult"
        PackagePath = $resolvedPackage
        RibbonParts = $ribbonInputs.Count
        RibbonRelationships = $script:RibbonDefinitions.Count
        RibbonCallbacks = $callbacks.Count
        CallbackNames = $callbacks
        MainContentType = $vbaSummary.MainContentType
        VbaProjectBytes = $vbaSummary.VbaProjectBytes
        VbaContentType = $vbaSummary.ContentType
        BinDefaultContentType = $vbaSummary.BinDefaultContentType
        RelationshipParts = $relationshipSummary.RelationshipParts
        InternalRelationships = $relationshipSummary.InternalRelationships
        ExternalRelationships = $relationshipSummary.ExternalRelationships
    }
}

Export-ModuleMember -Function @(
    "Add-IguanaTexRibbon",
    "Assert-VbaPackageClosure",
    "Assert-IguanaTexOfficePackage"
)

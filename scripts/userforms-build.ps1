#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Pinned", "WorkingTree")]
    [string]$Mode = "Pinned",

    [string]$OutputDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $PSScriptRoot "lib\IguanaTex.UserForms.psm1"

Import-Module -Name $modulePath -Force -DisableNameChecking -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $projectRoot ".build\vba-source"
}

$codec = Invoke-FrxEditBuild -ProjectRoot $projectRoot -Mode $Mode
$result = New-IguanaTexUserFormStaging `
    -ProjectRoot $projectRoot `
    -FrxEditPath $codec.ExecutablePath `
    -OutputDirectory $OutputDirectory

Write-Host ("Generated {0} UserForms and {1} VBA components: {2}" -f
    $result.FormCount,
    $result.ComponentCount,
    $result.OutputDirectory)

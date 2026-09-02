#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputDirectory,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Form,

    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $PSScriptRoot "lib\IguanaTex.UserForms.psm1"

Import-Module -Name $modulePath -Force -DisableNameChecking -ErrorAction Stop

$codec = Invoke-FrxEditBuild -ProjectRoot $projectRoot -Mode Pinned
$result = Export-IguanaTexCanonicalUserForms `
    -ProjectRoot $projectRoot `
    -FrxEditPath $codec.ExecutablePath `
    -InputDirectory $InputDirectory `
    -Form $Form `
    -DryRun:$DryRun

$verb = if ($DryRun) { "Validated" } else { "Updated" }
Write-Host ("{0} canonical source for {1} UserForm(s): {2}" -f
    $verb,
    $result.FormCount,
    ($result.Forms -join ", "))

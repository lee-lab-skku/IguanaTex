#requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet("Pinned", "WorkingTree")]
    [string]$Mode = "Pinned"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$modulePath = Join-Path $PSScriptRoot "lib\IguanaTex.UserForms.psm1"

Import-Module -Name $modulePath -Force -DisableNameChecking -ErrorAction Stop

$result = Invoke-FrxEditBuild -ProjectRoot $projectRoot -Mode $Mode

Write-Host ("FrxEdit commit: {0}" -f $result.HeadCommit)
Write-Host (".NET SDK: {0}" -f $result.DotNetSdkVersion)
Write-Host ("Executable SHA-256: {0}" -f $result.Sha256)
Write-Host ("Executable: {0}" -f $result.ExecutablePath)
Write-Host ("Provenance: {0}" -f $result.ProvenancePath)

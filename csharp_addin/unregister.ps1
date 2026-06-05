<#
.SYNOPSIS
    Reverses register.ps1 — removes the per-user (HKCU) registration of the
    PptEditor in-process PowerPoint COM add-in.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

$ProgId = 'PptEditor.AddIn'
$Clsid  = '{89E53E12-1EB0-4DDF-8017-16178D7DE66D}'

$paths = @(
    "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\$ProgId",
    "HKCU:\Software\Classes\$ProgId",
    "HKCU:\Software\Classes\CLSID\$Clsid"
)

foreach ($p in $paths) {
    if (Test-Path $p) {
        Remove-Item -Path $p -Recurse -Force
        Write-Host "Removed $p"
    }
}

Write-Host "Unregistered '$ProgId'." -ForegroundColor Green

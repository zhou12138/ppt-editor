param(
    [string]$SourcePath = (Join-Path $PSScriptRoot "skills\pptx-local-offline"),
    [string]$TargetPath = (Join-Path $HOME ".copilot\skills\pptx-local-offline")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $SourcePath)) {
    throw "Source skill directory not found: $SourcePath"
}

$sourceSkill = Get-Item -LiteralPath $SourcePath
if (-not $sourceSkill.PSIsContainer) {
    throw "Source path is not a directory: $SourcePath"
}

$requiredItems = @(
    "SKILL.md",
    "scripts",
    "references"
)

foreach ($item in $requiredItems) {
    $itemPath = Join-Path $SourcePath $item
    if (-not (Test-Path $itemPath)) {
        throw "Missing required skill item: $itemPath"
    }
}

New-Item -ItemType Directory -Force -Path $TargetPath | Out-Null

Get-ChildItem -LiteralPath $TargetPath -Force | Remove-Item -Recurse -Force
Copy-Item -Path (Join-Path $SourcePath "*") -Destination $TargetPath -Recurse -Force

Write-Host "Installed skill to: $TargetPath"
Write-Host "Contents:"
Get-ChildItem -LiteralPath $TargetPath -Recurse | ForEach-Object {
    Write-Host " - $($_.FullName)"
}
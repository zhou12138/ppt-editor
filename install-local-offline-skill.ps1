param(
    [string]$SourcePath = (Join-Path $PSScriptRoot "skills\pptx-local-offline"),
    [string]$TargetPath = (Join-Path $HOME ".copilot\skills\pptx-local-offline"),
    [switch]$SkipCSharp
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

# Publish the C# Interop host into the installed skill so `--backend csharp`
# works standalone (the installed skill has no repo csharp_interop/ above it).
$csharpProject = Join-Path $PSScriptRoot "csharp_interop\PptInteropHost"
if (-not $SkipCSharp -and (Test-Path $csharpProject)) {
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnet) {
        Write-Warning "dotnet SDK not found; skipping C# host. Install .NET SDK and re-run, or set PPTX_EDITOR_CSHARP_HOST to an existing exe."
    }
    else {
        $csharpOut = Join-Path $TargetPath "csharp_host"
        Write-Host "Publishing C# Interop host -> $csharpOut"
        try {
            & dotnet publish $csharpProject -c Release -o $csharpOut --nologo | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "C# host publish failed (exit $LASTEXITCODE); --backend csharp will be unavailable until built."
            }
        }
        catch {
            Write-Warning "C# host publish errored: $($_.Exception.Message)"
        }
    }
}

Write-Host "Installed skill to: $TargetPath"
Write-Host "Contents:"
Get-ChildItem -LiteralPath $TargetPath -Recurse | ForEach-Object {
    Write-Host " - $($_.FullName)"
}
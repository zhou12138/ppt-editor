param(
    [string]$SourcePath = (Join-Path $PSScriptRoot "skills\pptx-local-offline"),
    [string]$TargetPath = (Join-Path $HOME ".copilot\skills\pptx-local-offline"),
    [switch]$SkipCSharp,
    [switch]$SkipAddinRegister
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

# Build + deploy + register the IN-PROCESS C# COM add-in so `--backend csharp-addin`
# works standalone. Unlike the exe host above (out-of-process), this add-in runs
# inside POWERPNT.EXE (like VBA) and must be registered per-user so PowerPoint
# loads it at startup. Registration is HKCU-only (no admin/regasm needed).
$addinProject  = Join-Path $PSScriptRoot "csharp_addin\PptEditorAddin"
$addinRegister = Join-Path $PSScriptRoot "csharp_addin\register.ps1"
if (-not $SkipCSharp -and (Test-Path $addinProject)) {
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($null -eq $dotnet) {
        Write-Warning "dotnet SDK not found; skipping C# in-process add-in (--backend csharp-addin)."
    }
    else {
        $addinOut = Join-Path $TargetPath "csharp_addin"
        Write-Host "Building C# in-process add-in -> $addinOut"
        try {
            & dotnet build $addinProject -c Release --nologo | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "C# add-in build failed (exit $LASTEXITCODE); --backend csharp-addin will be unavailable until built."
            }
            else {
                $addinBuilt = Join-Path $addinProject "bin\Release\net48"
                New-Item -ItemType Directory -Force -Path $addinOut | Out-Null
                Copy-Item -Path (Join-Path $addinBuilt "*") -Destination $addinOut -Recurse -Force
                # bundle the (un)register scripts next to the deployed DLL for manual re-runs
                Copy-Item -Path $addinRegister -Destination $addinOut -Force
                Copy-Item -Path (Join-Path $PSScriptRoot "csharp_addin\unregister.ps1") -Destination $addinOut -Force
                $deployedDll = Join-Path $addinOut "PptEditorAddin.dll"
                if (-not $SkipAddinRegister) {
                    Write-Host "Registering in-process add-in (per-user HKCU) -> $deployedDll"
                    try {
                        & $addinRegister -DllPath $deployedDll
                    }
                    catch {
                        Write-Warning "Add-in registration errored: $($_.Exception.Message). Register manually: powershell -File `"$addinOut\register.ps1`" -DllPath `"$deployedDll`""
                    }
                }
                else {
                    Write-Host "Skipped add-in registration (-SkipAddinRegister). Register later: powershell -File `"$addinOut\register.ps1`" -DllPath `"$deployedDll`""
                }
            }
        }
        catch {
            Write-Warning "C# add-in build/deploy errored: $($_.Exception.Message)"
        }
    }
}

# Register the IN-PROCESS Python (pywin32) COM add-in so `--backend pywin32-addin`
# works. The scripts are already deployed under $TargetPath\scripts (copied above),
# so registration just runs the deployed module. Like the C# add-in it loads inside
# POWERPNT.EXE and is registered per-user (HKCU LoadBehavior=3, no admin).
$pyAddin = Join-Path $TargetPath "scripts\pptx_pyaddin.py"
if (-not $SkipAddinRegister -and (Test-Path $pyAddin)) {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($null -eq $python) { $python = Get-Command py -ErrorAction SilentlyContinue }
    if ($null -eq $python) {
        Write-Warning "Python not found; skipping Python in-process add-in (--backend pywin32-addin). Register later: python `"$pyAddin`""
    }
    else {
        Write-Host "Registering Python in-process add-in (per-user HKCU) -> $pyAddin"
        try {
            & $python.Source $pyAddin
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Python add-in registration exited $LASTEXITCODE (pywin32 required). Register manually: python `"$pyAddin`""
            }
        }
        catch {
            Write-Warning "Python add-in registration errored: $($_.Exception.Message). Register manually: python `"$pyAddin`""
        }
    }
}
elseif ($SkipAddinRegister -and (Test-Path $pyAddin)) {
    Write-Host "Skipped Python add-in registration (-SkipAddinRegister). Register later: python `"$pyAddin`""
}

Write-Host "Installed skill to: $TargetPath"
Write-Host "Contents:"
Get-ChildItem -LiteralPath $TargetPath -Recurse | ForEach-Object {
    Write-Host " - $($_.FullName)"
}
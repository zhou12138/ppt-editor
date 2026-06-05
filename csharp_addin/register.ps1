<#
.SYNOPSIS
    Per-user (HKCU) registration for the PptEditor in-process PowerPoint COM add-in.
    No administrator rights required: the .NET class is registered under
    HKCU\Software\Classes (per-user COM) and the add-in is listed under the
    PowerPoint per-user AddIns key with LoadBehavior=3 (load at startup).

.DESCRIPTION
    Mirrors what `regasm /codebase` would do, but entirely per-user so it can run
    unelevated. Assembly identity (name/version/token) is read from the built DLL
    so it always matches the binary.

.PARAMETER DllPath
    Optional path to PptEditorAddin.dll. Defaults to the Release build output.
#>
[CmdletBinding()]
param(
    [string]$DllPath
)

$ErrorActionPreference = 'Stop'

$ProgId   = 'PptEditor.AddIn'
$Clsid    = '{89E53E12-1EB0-4DDF-8017-16178D7DE66D}'
$ClassFqn = 'PptEditorAddin.Connect'

if (-not $DllPath) {
    $DllPath = Join-Path $PSScriptRoot 'PptEditorAddin\bin\Release\net48\PptEditorAddin.dll'
}
$DllPath = (Resolve-Path -LiteralPath $DllPath).Path
Write-Host "DLL: $DllPath"

# --- read assembly identity from the built binary so values always match ---
$asmName     = [System.Reflection.AssemblyName]::GetAssemblyName($DllPath)
$asmFullName = $asmName.FullName              # e.g. PptEditorAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
$asmVersion  = $asmName.Version.ToString()    # e.g. 1.0.0.0
$runtimeVer  = 'v4.0.30319'
$codeBase    = ([System.Uri]$DllPath).AbsoluteUri   # file:///C:/...
Write-Host "Assembly: $asmFullName"

function Set-Key([string]$path) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

# ---- 1) CLSID registration (per-user COM, mscoree shim) ----
$clsidRoot = "HKCU:\Software\Classes\CLSID\$Clsid"
Set-Key $clsidRoot
New-ItemProperty -Path $clsidRoot -Name '(default)' -Value $ClassFqn -PropertyType String -Force | Out-Null

$inproc = "$clsidRoot\InprocServer32"
Set-Key $inproc
# default = mscoree.dll, plus the managed registration values + a version subkey
New-ItemProperty -Path $inproc -Name '(default)'      -Value 'mscoree.dll'  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'ThreadingModel' -Value 'Both'         -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Class'          -Value $ClassFqn      -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'Assembly'       -Value $asmFullName   -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'RuntimeVersion' -Value $runtimeVer    -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inproc -Name 'CodeBase'       -Value $codeBase      -PropertyType String -Force | Out-Null

$inprocVer = "$inproc\$asmVersion"
Set-Key $inprocVer
New-ItemProperty -Path $inprocVer -Name 'Class'          -Value $ClassFqn    -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'Assembly'       -Value $asmFullName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'RuntimeVersion' -Value $runtimeVer  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $inprocVer -Name 'CodeBase'       -Value $codeBase    -PropertyType String -Force | Out-Null

$clsidProgId = "$clsidRoot\ProgId"
Set-Key $clsidProgId
New-ItemProperty -Path $clsidProgId -Name '(default)' -Value $ProgId -PropertyType String -Force | Out-Null

# ---- 2) ProgId -> CLSID ----
$progRoot = "HKCU:\Software\Classes\$ProgId"
Set-Key $progRoot
New-ItemProperty -Path $progRoot -Name '(default)' -Value 'PptEditor In-Process Add-in' -PropertyType String -Force | Out-Null
$progClsid = "$progRoot\CLSID"
Set-Key $progClsid
New-ItemProperty -Path $progClsid -Name '(default)' -Value $Clsid -PropertyType String -Force | Out-Null

# ---- 3) PowerPoint per-user AddIns entry (LoadBehavior=3 => load at startup) ----
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\$ProgId"
Set-Key $addinKey
New-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'PptEditor In-Process Add-in' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'Description'  -Value 'In-process bridge for fair benchmarking' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host "Registered '$ProgId' (CLSID $Clsid) per-user. LoadBehavior=3." -ForegroundColor Green

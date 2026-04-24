param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [switch]$StartAtSolidWorksStartup
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$framework = 'net8.0-windows'
$guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'
$title = 'AFCA Piping BOM Generator'
$description = 'SolidWorks piping BOM add-in for profile-driven grouped CSV/XLSX exports.'
$comHostPath = Join-Path $repoRoot "src\SolidWorksBOMAddin\bin\$Configuration\$framework\AFCA.PipingBom.Generator.comhost.dll"

if (-not (Test-Path $comHostPath)) {
    throw "Build output not found at '$comHostPath'. Build the solution first."
}

& regsvr32.exe /s $comHostPath

$addinsKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\SolidWorks\AddIns\$guid")
$addinsKey.SetValue($null, 0, [Microsoft.Win32.RegistryValueKind]::DWord)
$addinsKey.SetValue('Title', $title, [Microsoft.Win32.RegistryValueKind]::String)
$addinsKey.SetValue('Description', $description, [Microsoft.Win32.RegistryValueKind]::String)
$addinsKey.Dispose()

$startupValue = if ($StartAtSolidWorksStartup) { 1 } else { 0 }
$startupKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\SolidWorks\AddInsStartup")
$startupKey.SetValue($guid, $startupValue, [Microsoft.Win32.RegistryValueKind]::DWord)
$startupKey.Dispose()

Write-Host "Registered $title from $comHostPath"

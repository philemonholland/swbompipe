param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$framework = 'net8.0-windows'
$guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'
$comHostPath = Join-Path $repoRoot "src\SolidWorksBOMAddin\bin\$Configuration\$framework\AFCA.PipingBom.Generator.comhost.dll"

if (Test-Path $comHostPath) {
    & regsvr32.exe /u /s $comHostPath
}

[Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\SolidWorks\AddIns\$guid", $false)

$startupKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("Software\SolidWorks\AddInsStartup", $true)
if ($startupKey -and $startupKey.GetValue($guid) -ne $null) {
    $startupKey.DeleteValue($guid, $false)
}

if ($startupKey) {
    $startupKey.Dispose()
}

Write-Host "Unregistered AFCA Piping BOM Generator"

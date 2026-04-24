param(
    [string]$InstallRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe')
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

Unregister-BomPipeShellVerb

if (Test-Path $InstallRoot) {
    Remove-Item -Path $InstallRoot -Recurse -Force
}

Write-Host "Uninstalled BOMPipe from $InstallRoot"

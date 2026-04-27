param(
    [string]$InstallRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe'),
    [string[]]$PdmVaultName = @(),
    [switch]$SkipSolidWorksAddinUnregistration,
    [switch]$SkipPdmUnregistration
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

if (-not $SkipPdmUnregistration) {
    Unregister-BomPipePdmAddin -InstallRoot $InstallRoot -VaultNames $PdmVaultName
}

if (-not $SkipSolidWorksAddinUnregistration) {
    Stop-BomPipeSolidWorksLoader -InstallRoot $InstallRoot
    Unregister-BomPipeSolidWorksLoader
    Unregister-BomPipeAddin -ComHostPath (Get-BomPipeInstalledComHostPath -InstallRoot $InstallRoot)
}

Unregister-BomPipeShellVerb

if (Test-Path $InstallRoot) {
    Remove-Item -Path $InstallRoot -Recurse -Force
}

Write-Host "Uninstalled BOMPipe from $InstallRoot"

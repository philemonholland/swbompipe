param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [string]$InstallRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe'),
    [switch]$ForceRebuild
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

Install-BomPipePayload -InstallRoot $InstallRoot -Configuration $Configuration -ForceRebuild:$ForceRebuild
Register-BomPipeShellVerb -InstallRoot $InstallRoot

Write-Host "Installed BOMPipe to $InstallRoot"
Write-Host 'Explorer and PDM Professional right-click integration is registered for .SLDASM files.'

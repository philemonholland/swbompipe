param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [string]$InstallRoot = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe'),
    [switch]$StartAtSolidWorksStartup,
    [string[]]$PdmVaultName = @(),
    [switch]$SkipSolidWorksAddinRegistration,
    [switch]$SkipPdmRegistration,
    [switch]$ForceRebuild
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

Install-BomPipePayload -InstallRoot $InstallRoot -Configuration $Configuration -ForceRebuild:$ForceRebuild
$solidWorksRegistered = $false
$pdmRegistered = $false

if (-not $SkipSolidWorksAddinRegistration) {
    Register-BomPipeAddin -ComHostPath (Get-BomPipeInstalledComHostPath -InstallRoot $InstallRoot) -StartAtSolidWorksStartup:$StartAtSolidWorksStartup
    if (-not (Test-BomPipeComActivation -ProgId 'AFCA.PipingBom.Generator')) {
        throw 'BOMPipe COM activation failed after registration.'
    }

    Register-BomPipeSolidWorksLoader -InstallRoot $InstallRoot
    Start-BomPipeSolidWorksLoader -InstallRoot $InstallRoot
    $solidWorksRegistered = $true
}
Register-BomPipeShellVerb -InstallRoot $InstallRoot
if (-not $SkipPdmRegistration) {
    $pdmRegistered = Register-BomPipePdmAddin -InstallRoot $InstallRoot -VaultNames $PdmVaultName
}

Write-Host "Installed BOMPipe to $InstallRoot"
if ($solidWorksRegistered) {
    Write-Host 'SolidWorks per-user registration and loader startup have been applied.'
}

Write-Host 'Explorer integration has been applied.'

if ($SkipPdmRegistration) {
    Write-Host 'PDM Professional registration was skipped by request.'
}
elseif ($pdmRegistered) {
    Write-Host 'PDM Professional registration was applied and verified.'
}
else {
    Write-Host 'PDM Professional registration was not applied.'
}

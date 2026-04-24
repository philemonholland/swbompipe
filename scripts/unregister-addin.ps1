param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [string]$ComHostPath
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

if ([string]::IsNullOrWhiteSpace($ComHostPath)) {
    $ComHostPath = Get-BomPipeDefaultComHostPath -Configuration $Configuration
}

Unregister-BomPipeAddin -ComHostPath $ComHostPath

Write-Host "Unregistered AFCA Piping BOM Generator"

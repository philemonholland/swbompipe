param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [switch]$StartAtSolidWorksStartup,
    [string]$ComHostPath
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'BomPipe.Install.Common.ps1')

if ([string]::IsNullOrWhiteSpace($ComHostPath)) {
    $ComHostPath = Get-BomPipeDefaultComHostPath -Configuration $Configuration
}

Register-BomPipeAddin -ComHostPath $ComHostPath -StartAtSolidWorksStartup:$StartAtSolidWorksStartup

Write-Host "Registered AFCA Piping BOM Generator from $ComHostPath"

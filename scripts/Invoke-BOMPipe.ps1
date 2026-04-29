param(
    [Parameter(Mandatory)]
    [string]$AssemblyPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

function Get-BomPipeDotNetPath {
    $userLocalDotNet = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    if (Test-Path $userLocalDotNet) {
        return $userLocalDotNet
    }

    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($dotnetCommand) {
        return $dotnetCommand.Source
    }

    throw 'dotnet.exe was not found and no BomPipeLauncher.exe was installed.'
}

function Get-SafeFileStem {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    foreach ($invalidCharacter in [System.IO.Path]::GetInvalidFileNameChars()) {
        $stem = $stem.Replace($invalidCharacter, '_')
    }

    if ([string]::IsNullOrWhiteSpace($stem)) {
        return 'Assembly'
    }

    $stem
}

try {
    if (-not (Test-Path $AssemblyPath)) {
        throw "Assembly not found: $AssemblyPath"
    }

    $launcherExe = Join-Path $PSScriptRoot 'BomPipeLauncher\BomPipeLauncher.exe'
    $launcherDll = Join-Path $PSScriptRoot 'BomPipeLauncher\BomPipeLauncher.dll'
    $exportsRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'AFCA\BOMPipe\Exports'
    $fileStem = Get-SafeFileStem -Path $AssemblyPath
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

    New-Item -Path $exportsRoot -ItemType Directory -Force | Out-Null

    $outputPath = Join-Path $exportsRoot "$fileStem.$timestamp.xlsx"
    $bomDbOutputPath = Join-Path $exportsRoot "$fileStem.$timestamp.bomdb.json"
    $debugReportPath = Join-Path $exportsRoot "$fileStem.$timestamp.debug.json"
    $arguments = @(
        '--assembly', $AssemblyPath,
        '--format', 'xlsx',
        '--output', $outputPath,
        '--bomdb-output', $bomDbOutputPath,
        '--debug-report', $debugReportPath
    )

    if (Test-Path $launcherExe) {
        $commandOutput = (& $launcherExe @arguments 2>&1 | Out-String).Trim()
    }
    elseif (Test-Path $launcherDll) {
        $dotnet = Get-BomPipeDotNetPath
        $commandOutput = (& $dotnet $launcherDll @arguments 2>&1 | Out-String).Trim()
    }
    else {
        throw "BomPipeLauncher was not found under '$PSScriptRoot'."
    }

    if ($LASTEXITCODE -ne 0) {
        throw "BOMPipe exited with code $LASTEXITCODE.`n$commandOutput"
    }

    $message = "BOM exported to:`n$outputPath`n`nBOMDB import JSON:`n$bomDbOutputPath`n`nDebug report:`n$debugReportPath"
    [System.Windows.Forms.MessageBox]::Show(
        $message,
        'BOMPipe',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message,
        'BOMPipe',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit 1
}

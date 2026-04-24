Set-StrictMode -Version Latest

function Get-BomPipeRepoRoot {
    Split-Path -Parent $PSScriptRoot
}

function Get-BomPipeFramework {
    'net8.0-windows'
}

function Get-BomPipeDefaultInstallRoot {
    Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe'
}

function Get-BomPipeBuildOutputPath {
    param(
        [Parameter(Mandatory)]
        [string]$ProjectName,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration
    )

    Join-Path (Get-BomPipeRepoRoot) "src\$ProjectName\bin\$Configuration\$(Get-BomPipeFramework)"
}

function Get-BomPipeDotNetPath {
    $userLocalDotNet = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    if (Test-Path $userLocalDotNet) {
        return $userLocalDotNet
    }

    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    if ($dotnetCommand) {
        return $dotnetCommand.Source
    }

    throw 'dotnet.exe was not found. Install the .NET SDK or build the solution first so the installer can reuse existing outputs.'
}

function Initialize-BomPipeDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force
    }

    New-Item -Path $Path -ItemType Directory -Force | Out-Null
}

function Copy-BomPipePayload {
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$DestinationPath
    )

    if (-not (Test-Path $SourcePath)) {
        throw "Build output not found at '$SourcePath'."
    }

    Initialize-BomPipeDirectory -Path $DestinationPath
    Copy-Item -Path (Join-Path $SourcePath '*') -Destination $DestinationPath -Recurse -Force
}

function Publish-BomPipeProject {
    param(
        [Parameter(Mandatory)]
        [string]$ProjectRelativePath,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration
    )

    Initialize-BomPipeDirectory -Path $OutputPath

    $dotnet = Get-BomPipeDotNetPath
    $projectPath = Join-Path (Get-BomPipeRepoRoot) $ProjectRelativePath

    & $dotnet publish $projectPath -c $Configuration -o $OutputPath --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed for '$projectPath'."
    }
}

function Install-BomPipePayload {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration,

        [switch]$ForceRebuild
    )

    $launcherTarget = Join-Path $InstallRoot 'BomPipeLauncher'
    $addinTarget = Join-Path $InstallRoot 'SolidWorksBOMAddin'
    $invokerTarget = Join-Path $InstallRoot 'Invoke-BOMPipe.ps1'
    $invokerSource = Join-Path $PSScriptRoot 'Invoke-BOMPipe.ps1'

    New-Item -Path $InstallRoot -ItemType Directory -Force | Out-Null

    $launcherSource = Get-BomPipeBuildOutputPath -ProjectName 'BomPipeLauncher' -Configuration $Configuration
    $addinSource = Get-BomPipeBuildOutputPath -ProjectName 'SolidWorksBOMAddin' -Configuration $Configuration

    if ($ForceRebuild -or -not (Test-Path $launcherSource) -or -not (Test-Path $addinSource)) {
        Publish-BomPipeProject -ProjectRelativePath 'src\BomPipeLauncher\BomPipeLauncher.csproj' -OutputPath $launcherTarget -Configuration $Configuration
        Publish-BomPipeProject -ProjectRelativePath 'src\SolidWorksBOMAddin\SolidWorksBOMAddin.csproj' -OutputPath $addinTarget -Configuration $Configuration
    }
    else {
        Copy-BomPipePayload -SourcePath $launcherSource -DestinationPath $launcherTarget
        Copy-BomPipePayload -SourcePath $addinSource -DestinationPath $addinTarget
    }

    Copy-Item -Path $invokerSource -Destination $invokerTarget -Force
}

function Get-BomPipeInstalledComHostPath {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    Join-Path $InstallRoot 'SolidWorksBOMAddin\AFCA.PipingBom.Generator.comhost.dll'
}

function Get-BomPipeDefaultComHostPath {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration
    )

    Join-Path (Get-BomPipeBuildOutputPath -ProjectName 'SolidWorksBOMAddin' -Configuration $Configuration) 'AFCA.PipingBom.Generator.comhost.dll'
}

function Invoke-BomPipeRegSvr32 {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments
    )

    $process = Start-Process -FilePath 'regsvr32.exe' -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) {
        throw "regsvr32.exe failed for arguments: $($Arguments -join ' ')"
    }
}

function Register-BomPipeAddin {
    param(
        [Parameter(Mandatory)]
        [string]$ComHostPath,

        [switch]$StartAtSolidWorksStartup
    )

    $guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'
    $title = 'AFCA Piping BOM Generator'
    $description = 'SolidWorks piping BOM add-in for profile-driven grouped CSV/XLSX exports.'

    if (-not (Test-Path $ComHostPath)) {
        throw "COM host not found at '$ComHostPath'."
    }

    Invoke-BomPipeRegSvr32 -Arguments @('/s', $ComHostPath)

    $addinsKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\SolidWorks\AddIns\$guid")
    $addinsKey.SetValue($null, 0, [Microsoft.Win32.RegistryValueKind]::DWord)
    $addinsKey.SetValue('Title', $title, [Microsoft.Win32.RegistryValueKind]::String)
    $addinsKey.SetValue('Description', $description, [Microsoft.Win32.RegistryValueKind]::String)
    $addinsKey.Dispose()

    $startupValue = if ($StartAtSolidWorksStartup) { 1 } else { 0 }
    $startupKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\SolidWorks\AddInsStartup')
    $startupKey.SetValue($guid, $startupValue, [Microsoft.Win32.RegistryValueKind]::DWord)
    $startupKey.Dispose()
}

function Unregister-BomPipeAddin {
    param(
        [Parameter(Mandatory)]
        [string]$ComHostPath
    )

    $guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'

    if (Test-Path $ComHostPath) {
        Invoke-BomPipeRegSvr32 -Arguments @('/u', '/s', $ComHostPath)
    }

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\SolidWorks\AddIns\$guid", $false)

    $startupKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\SolidWorks\AddInsStartup', $true)
    if ($startupKey -and $startupKey.GetValue($guid) -ne $null) {
        $startupKey.DeleteValue($guid, $false)
    }

    if ($startupKey) {
        $startupKey.Dispose()
    }
}

function Get-BomPipeInvokerPath {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    Join-Path $InstallRoot 'Invoke-BOMPipe.ps1'
}

function Get-BomPipeShellCommand {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    $powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $invokerPath = Get-BomPipeInvokerPath -InstallRoot $InstallRoot
    '"{0}" -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}" "%1"' -f $powerShellExe, $invokerPath
}

function Register-BomPipeShellVerb {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    $launcherExe = Join-Path $InstallRoot 'BomPipeLauncher\BomPipeLauncher.exe'
    $iconPath = if (Test-Path $launcherExe) { $launcherExe } else { Join-Path $InstallRoot 'SolidWorksBOMAddin\AFCA.PipingBom.Generator.dll' }
    $command = Get-BomPipeShellCommand -InstallRoot $InstallRoot
    $registryPaths = @(
        'Software\Classes\SystemFileAssociations\.sldasm\shell\AFCA.BOMPipe',
        'Software\Classes\.sldasm\shell\AFCA.BOMPipe'
    )

    foreach ($registryPath in $registryPaths) {
        $shellKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($registryPath)
        $shellKey.SetValue('MUIVerb', 'Generate BOM with BOMPipe', [Microsoft.Win32.RegistryValueKind]::String)
        $shellKey.SetValue('Icon', $iconPath, [Microsoft.Win32.RegistryValueKind]::String)
        $shellKey.SetValue('MultiSelectModel', 'Single', [Microsoft.Win32.RegistryValueKind]::String)
        $shellKey.Dispose()

        $commandKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("$registryPath\command")
        $commandKey.SetValue($null, $command, [Microsoft.Win32.RegistryValueKind]::String)
        $commandKey.Dispose()
    }
}

function Unregister-BomPipeShellVerb {
    $registryPaths = @(
        'Software\Classes\SystemFileAssociations\.sldasm\shell\AFCA.BOMPipe',
        'Software\Classes\.sldasm\shell\AFCA.BOMPipe'
    )

    foreach ($registryPath in $registryPaths) {
        [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($registryPath, $false)
    }
}

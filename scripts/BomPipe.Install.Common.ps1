Set-StrictMode -Version Latest

function Get-BomPipeRepoRoot {
    Split-Path -Parent $PSScriptRoot
}

function Get-BomPipeDefaultInstallRoot {
    Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AFCA\BOMPipe'
}

function Get-BomPipeProjectFramework {
    param(
        [Parameter(Mandatory)]
        [string]$ProjectName
    )

    switch ($ProjectName) {
        'BomPipeLauncher' { 'net8.0-windows' }
        'SolidWorksBOMAddin' { 'net8.0-windows' }
        'BomPipePdmAddin' { 'net48' }
        'BomPipePdmVaultInstaller' { 'net48' }
        default { throw "Unknown BOMPipe project '$ProjectName'." }
    }
}

function Get-BomPipeProjectRelativePath {
    param(
        [Parameter(Mandatory)]
        [string]$ProjectName
    )

    "src\$ProjectName\$ProjectName.csproj"
}

function Get-BomPipeBuildOutputPath {
    param(
        [Parameter(Mandatory)]
        [string]$ProjectName,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration
    )

    Join-Path (Get-BomPipeRepoRoot) "src\$ProjectName\bin\$Configuration\$(Get-BomPipeProjectFramework -ProjectName $ProjectName)"
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

    throw 'dotnet.exe was not found. Install the .NET SDK so BOMPipe can publish the current source during installation.'
}

function Get-BomPipeSourceRevision {
    $repoRoot = Get-BomPipeRepoRoot
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCommand) {
        return 'git-unavailable'
    }

    $revision = (& $gitCommand.Source -C $repoRoot rev-parse --short=12 HEAD 2>$null | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($revision)) {
        return 'unknown'
    }

    $status = (& $gitCommand.Source -C $repoRoot status --short 2>$null | Out-String).Trim()
    if (-not [string]::IsNullOrWhiteSpace($status)) {
        return "$revision-dirty"
    }

    return $revision
}

function Write-BomPipeBuildManifest {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration,

        [Parameter(Mandatory)]
        [object[]]$PayloadTargets
    )

    $manifest = [ordered]@{
        product = 'BOMPipe'
        configuration = $Configuration
        installed_at_utc = (Get-Date).ToUniversalTime().ToString('o')
        source_root = Get-BomPipeRepoRoot
        source_revision = Get-BomPipeSourceRevision
        machine = $env:COMPUTERNAME
        user = [Environment]::UserName
    }
    $json = $manifest | ConvertTo-Json -Depth 3
    $manifestFileName = 'bompipe-build-info.json'
    $manifestPaths = @((Join-Path $InstallRoot $manifestFileName))
    foreach ($payloadTarget in $PayloadTargets) {
        $manifestPaths += (Join-Path $payloadTarget.Target $manifestFileName)
    }

    foreach ($path in $manifestPaths) {
        $parent = Split-Path -Parent $path
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
        Set-Content -Path $path -Value $json -Encoding UTF8
    }

    Write-Host ("BOMPipe build manifest written: {0} ({1})" -f $manifest.source_revision, $manifest.installed_at_utc)
}

function Initialize-BomPipeDirectory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        try {
            Remove-Item -Path $Path -Recurse -Force
        }
        catch [System.UnauthorizedAccessException] {
            throw "Could not replace '$Path' because one or more files are locked. Close SolidWorks and any BOMPipe/SolidWorks loader processes, then run install again."
        }
        catch [System.IO.IOException] {
            throw "Could not replace '$Path' because one or more files are in use. Close SolidWorks and any BOMPipe/SolidWorks loader processes, then run install again."
        }
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
        [string]$ProjectName,

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [Parameter(Mandatory)]
        [ValidateSet('Debug', 'Release')]
        [string]$Configuration
    )

    $dotnet = Get-BomPipeDotNetPath
    $projectPath = Join-Path (Get-BomPipeRepoRoot) (Get-BomPipeProjectRelativePath -ProjectName $ProjectName)

    Initialize-BomPipeDirectory -Path $OutputPath

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

    $dotnet = Get-BomPipeDotNetPath
    Write-Host "Using .NET SDK host: $dotnet"

    $payloadTargets = @(
        @{ Project = 'BomPipeLauncher'; Target = Join-Path $InstallRoot 'BomPipeLauncher' },
        @{ Project = 'SolidWorksBOMAddin'; Target = Join-Path $InstallRoot 'SolidWorksBOMAddin' },
        @{ Project = 'BomPipePdmAddin'; Target = Join-Path $InstallRoot 'BomPipePdmAddin' },
        @{ Project = 'BomPipePdmVaultInstaller'; Target = Join-Path $InstallRoot 'BomPipePdmVaultInstaller' }
    )

    $invokerTarget = Join-Path $InstallRoot 'Invoke-BOMPipe.ps1'
    $invokerSource = Join-Path $PSScriptRoot 'Invoke-BOMPipe.ps1'
    $solidWorksLoaderTarget = Join-Path $InstallRoot 'Watch-BomPipeSolidWorksAddin.ps1'
    $solidWorksLoaderSource = Join-Path $PSScriptRoot 'Watch-BomPipeSolidWorksAddin.ps1'
    $solidWorksEnsureTarget = Join-Path $InstallRoot 'Ensure-BomPipeSolidWorksAddin.vbs'
    $solidWorksEnsureSource = Join-Path $PSScriptRoot 'Ensure-BomPipeSolidWorksAddin.vbs'

    New-Item -Path $InstallRoot -ItemType Directory -Force | Out-Null

    foreach ($payloadTarget in $payloadTargets) {
        Publish-BomPipeProject -ProjectName $payloadTarget.Project -OutputPath $payloadTarget.Target -Configuration $Configuration
    }

    Copy-Item -Path $invokerSource -Destination $invokerTarget -Force
    Copy-Item -Path $solidWorksLoaderSource -Destination $solidWorksLoaderTarget -Force
    Copy-Item -Path $solidWorksEnsureSource -Destination $solidWorksEnsureTarget -Force
    Write-BomPipeBuildManifest -InstallRoot $InstallRoot -Configuration $Configuration -PayloadTargets $payloadTargets
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

function Register-BomPipeComClass {
    param(
        [Parameter(Mandatory)]
        [string]$ClassId,

        [Parameter(Mandatory)]
        [string]$ProgId,

        [Parameter(Mandatory)]
        [string]$ComHostPath,

        [Parameter(Mandatory)]
        [string]$DisplayName
    )

    $clsidKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Classes\CLSID\$ClassId")
    $clsidKey.SetValue($null, $DisplayName, [Microsoft.Win32.RegistryValueKind]::String)
    $clsidKey.SetValue('ProgId', $ProgId, [Microsoft.Win32.RegistryValueKind]::String)

    $inprocKey = $clsidKey.CreateSubKey('InprocServer32')
    $inprocKey.SetValue($null, $ComHostPath, [Microsoft.Win32.RegistryValueKind]::String)
    $inprocKey.SetValue('ThreadingModel', 'Both', [Microsoft.Win32.RegistryValueKind]::String)
    $inprocKey.Dispose()
    $clsidKey.Dispose()

    $progIdKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Classes\$ProgId")
    $progIdKey.SetValue($null, $DisplayName, [Microsoft.Win32.RegistryValueKind]::String)
    $progIdClsidKey = $progIdKey.CreateSubKey('CLSID')
    $progIdClsidKey.SetValue($null, $ClassId, [Microsoft.Win32.RegistryValueKind]::String)
    $progIdClsidKey.Dispose()
    $progIdKey.Dispose()
}

function Unregister-BomPipeComClass {
    param(
        [Parameter(Mandatory)]
        [string]$ClassId,

        [Parameter(Mandatory)]
        [string]$ProgId
    )

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Classes\CLSID\$ClassId", $false)
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Classes\$ProgId", $false)
}

function Register-BomPipeAddin {
    param(
        [Parameter(Mandatory)]
        [string]$ComHostPath,

        [switch]$StartAtSolidWorksStartup
    )

    $guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'
    $progId = 'AFCA.PipingBom.Generator'
    $title = 'AFCA Piping BOM Generator'
    $description = 'SolidWorks piping BOM add-in for profile-driven grouped CSV/XLSX exports.'

    if (-not (Test-Path $ComHostPath)) {
        throw "COM host not found at '$ComHostPath'."
    }

    Register-BomPipeComClass -ClassId $guid -ProgId $progId -ComHostPath $ComHostPath -DisplayName $title

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

function Test-BomPipeComActivation {
    param(
        [Parameter(Mandatory)]
        [string]$ProgId
    )

    $scriptPath = Join-Path ([System.IO.Path]::GetTempPath()) ("bompipe-com-activation-{0}.vbs" -f [Guid]::NewGuid().ToString('N'))

    @"
On Error Resume Next
Set obj = CreateObject("$ProgId")
If Err.Number <> 0 Then
  WScript.Echo Err.Description
  WScript.Quit 1
End If
WScript.Quit 0
"@ | Set-Content -Path $scriptPath -Encoding ASCII

    try {
        $cscriptPath = Join-Path $env:SystemRoot 'System32\cscript.exe'
        & $cscriptPath //NoLogo $scriptPath | Out-Null
        return $LASTEXITCODE -eq 0
    }
    finally {
        if (Test-Path $scriptPath) {
            Remove-Item -Path $scriptPath -Force
        }
    }
}

function Unregister-BomPipeAddin {
    param(
        [Parameter(Mandatory)]
        [string]$ComHostPath
    )

    $guid = '{E1DB31B7-4F08-4B0D-99E4-69AFC34C5B1A}'
    $progId = 'AFCA.PipingBom.Generator'

    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\SolidWorks\AddIns\$guid", $false)

    $startupKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\SolidWorks\AddInsStartup', $true)
    if ($startupKey -and $startupKey.GetValue($guid) -ne $null) {
        $startupKey.DeleteValue($guid, $false)
    }

    if ($startupKey) {
        $startupKey.Dispose()
    }

    Unregister-BomPipeComClass -ClassId $guid -ProgId $progId
}

function Get-BomPipeSolidWorksLoaderPath {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    Join-Path $InstallRoot 'Watch-BomPipeSolidWorksAddin.ps1'
}

function Get-BomPipeSolidWorksLoaderCommand {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    $powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $loaderPath = Get-BomPipeSolidWorksLoaderPath -InstallRoot $InstallRoot
    '"{0}" -NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}"' -f $powershellExe, $loaderPath
}

function Register-BomPipeSolidWorksLoader {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    $runKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey('Software\Microsoft\Windows\CurrentVersion\Run')
    $runKey.SetValue('AFCA.BOMPipe.SolidWorksLoader', (Get-BomPipeSolidWorksLoaderCommand -InstallRoot $InstallRoot), [Microsoft.Win32.RegistryValueKind]::String)
    $runKey.Dispose()
}

function Unregister-BomPipeSolidWorksLoader {
    $runKey = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Run', $true)
    if ($runKey -and $runKey.GetValue('AFCA.BOMPipe.SolidWorksLoader') -ne $null) {
        $runKey.DeleteValue('AFCA.BOMPipe.SolidWorksLoader', $false)
    }

    if ($runKey) {
        $runKey.Dispose()
    }
}

function Stop-BomPipeSolidWorksLoader {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    $loaderPath = Get-BomPipeSolidWorksLoaderPath -InstallRoot $InstallRoot
    $processes = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" |
        Where-Object { $_.CommandLine -like "*$loaderPath*" }

    foreach ($process in $processes) {
        Stop-Process -Id $process.ProcessId -Force
    }
}

function Start-BomPipeSolidWorksLoader {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    Stop-BomPipeSolidWorksLoader -InstallRoot $InstallRoot

    $loaderPath = Get-BomPipeSolidWorksLoaderPath -InstallRoot $InstallRoot
    if (-not (Test-Path $loaderPath)) {
        throw "The SolidWorks loader script was not found at '$loaderPath'."
    }

    $powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    Start-Process -FilePath $powershellExe -ArgumentList @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', $loaderPath) -WindowStyle Hidden | Out-Null
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

function Get-BomPipeInstalledPdmVaultInstallerPath {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    Join-Path $InstallRoot 'BomPipePdmVaultInstaller\BomPipePdmVaultInstaller.exe'
}

function Test-BomPipePdmClientInstalled {
    $candidatePaths = @(
        'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS PDM\EPDM.Interop.epdm.dll',
        'C:\Program Files\SOLIDWORKS PDM\EPDM.Interop.epdm.dll'
    )

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path $candidatePath) {
            return $true
        }
    }

    try {
        $vaultType = [type]::GetTypeFromProgID('ConisioLib.EdmVault')
        if ($vaultType) {
            return $true
        }
    }
    catch {
        return $false
    }

    return $false
}

function Invoke-BomPipePdmVaultInstaller {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('register', 'unregister', 'list-vaults', 'verify')]
        [string]$Command,

        [Parameter(Mandatory)]
        [string]$InstallRoot,

        [string[]]$VaultNames = @()
    )

    $installerPath = Get-BomPipeInstalledPdmVaultInstallerPath -InstallRoot $InstallRoot
    if (-not (Test-Path $installerPath)) {
        if ($Command -eq 'unregister') {
            Write-Host "Skipping PDM Professional unregistration because the vault installer was not found at '$installerPath'."
            return @()
        }

        throw "PDM vault installer executable was not found at '$installerPath'."
    }

    $arguments = @($Command)
    if ($Command -ne 'list-vaults') {
        $arguments += @('--install-root', $InstallRoot)
    }

    foreach ($vaultName in $VaultNames) {
        if (-not [string]::IsNullOrWhiteSpace($vaultName)) {
            $arguments += @('--vault', $vaultName)
        }
    }

    $stdoutPath = [System.IO.Path]::GetTempFileName()
    $stderrPath = [System.IO.Path]::GetTempFileName()

    try {
        $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $stdout = if (Test-Path $stdoutPath) { Get-Content -Path $stdoutPath -Raw } else { '' }
        $stderr = if (Test-Path $stderrPath) { Get-Content -Path $stderrPath -Raw } else { '' }

        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            Write-Host $stdout.TrimEnd()
        }

        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            Write-Warning $stderr.TrimEnd()
        }

        if ($process.ExitCode -ne 0) {
            throw "PDM vault installer exited with code $($process.ExitCode)."
        }

        return @($stdout -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    finally {
        Remove-Item -Path $stdoutPath, $stderrPath -Force -ErrorAction SilentlyContinue
    }
}

function Get-BomPipePdmVaultNames {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot
    )

    if (-not (Test-BomPipePdmClientInstalled)) {
        return @()
    }

    $vaultNames = Invoke-BomPipePdmVaultInstaller -Command 'list-vaults' -InstallRoot $InstallRoot
    return @($vaultNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Register-BomPipePdmAddin {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot,

        [string[]]$VaultNames = @()
    )

    if (-not (Test-BomPipePdmClientInstalled)) {
        Write-Host 'Skipping PDM Professional registration because the local PDM client was not detected.'
        return $false
    }

    $targetVaultNames = @($VaultNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($targetVaultNames.Count -eq 0) {
        $targetVaultNames = @(Get-BomPipePdmVaultNames -InstallRoot $InstallRoot)
    }

    if ($targetVaultNames.Count -eq 0) {
        Write-Host 'Skipping PDM Professional registration because no local PDM vault views were found.'
        return $false
    }

    $null = Invoke-BomPipePdmVaultInstaller -Command 'register' -InstallRoot $InstallRoot -VaultNames $VaultNames
    $null = Invoke-BomPipePdmVaultInstaller -Command 'verify' -InstallRoot $InstallRoot -VaultNames $VaultNames
    return $true
}

function Unregister-BomPipePdmAddin {
    param(
        [Parameter(Mandatory)]
        [string]$InstallRoot,

        [string[]]$VaultNames = @()
    )

    if (-not (Test-BomPipePdmClientInstalled)) {
        Write-Host 'Skipping PDM Professional unregistration because the local PDM client was not detected.'
        return
    }

    $null = Invoke-BomPipePdmVaultInstaller -Command 'unregister' -InstallRoot $InstallRoot -VaultNames $VaultNames
}

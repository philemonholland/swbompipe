param(
    [int]$PollIntervalSeconds = 5
)

$ErrorActionPreference = 'Stop'

$mutexName = 'Local\AFCA.BOMPipe.SolidWorksLoader'
$createdNew = $false
$mutex = [System.Threading.Mutex]::new($true, $mutexName, [ref]$createdNew)

if (-not $createdNew) {
    $mutex.Dispose()
    return
}

try {
    $cscriptPath = Join-Path $env:SystemRoot 'System32\cscript.exe'
    $ensureScriptPath = Join-Path $PSScriptRoot 'Ensure-BomPipeSolidWorksAddin.vbs'

    if (-not (Test-Path $cscriptPath)) {
        throw "cscript.exe was not found at '$cscriptPath'."
    }

    if (-not (Test-Path $ensureScriptPath)) {
        throw "The SolidWorks add-in ensure script was not found at '$ensureScriptPath'."
    }

    while ($true) {
        $solidWorksRunning = @(Get-Process -Name 'SLDWORKS' -ErrorAction SilentlyContinue).Count -gt 0

        if ($solidWorksRunning) {
            & $cscriptPath //NoLogo $ensureScriptPath | Out-Null
        }

        Start-Sleep -Seconds $PollIntervalSeconds
    }
}
finally {
    if ($mutex) {
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}

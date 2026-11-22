$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Set-Location $PSScriptRoot

function Invoke-Checked {
    param([string]$FilePath, [string[]]$ArgumentList)
    & $FilePath @ArgumentList
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code ${LASTEXITCODE}: $FilePath"
    }
}

if ($env:OS -ne 'Windows_NT') {
    throw 'Build DBImport.exe on Windows with 64-bit Python 3.10 or newer (3.12 recommended).'
}

$buildPython = Join-Path $PSScriptRoot '.venv-build\Scripts\python.exe'
if (-not (Test-Path $buildPython)) {
    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        Invoke-Checked -FilePath $pyLauncher.Source -ArgumentList @('-3', '-m', 'venv', '.venv-build')
    } elseif ($python) {
        Invoke-Checked -FilePath $python.Source -ArgumentList @('-m', 'venv', '.venv-build')
    } else {
        throw 'Install 64-bit Python 3.12 on this build PC, then run build_windows.bat again.'
    }
}

$runtimeCheck = "import struct,sys; sys.exit(0 if sys.platform == 'win32' and struct.calcsize('P') == 8 and sys.version_info >= (3,10) else '64-bit Windows Python 3.10 or newer is required')"
Invoke-Checked -FilePath $buildPython -ArgumentList @('-c', $runtimeCheck)
Invoke-Checked -FilePath $buildPython -ArgumentList @('-m', 'pip', 'install', '--disable-pip-version-check', '-r', 'requirements-build.txt')
Invoke-Checked -FilePath $buildPython -ArgumentList @('-m', 'unittest', 'discover', '-s', 'tests', '-v')
Invoke-Checked -FilePath $buildPython -ArgumentList @('-m', 'PyInstaller', '--clean', '--noconfirm', 'DBImport.spec')
Invoke-Checked -FilePath $buildPython -ArgumentList @('tools\smoke_executable.py', 'dist\DBImport.exe')

$checksum = Get-FileHash 'dist\DBImport.exe' -Algorithm SHA256
"$($checksum.Hash.ToLower())  DBImport.exe" | Set-Content -Encoding ascii 'dist\DBImport.exe.sha256'
Copy-Item 'WINDOWS.md' 'dist\WINDOWS.md'
Write-Host "`nReady to share: $PSScriptRoot\dist\DBImport.exe"

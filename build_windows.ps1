param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SpecPath = Join-Path $ProjectRoot "JjukfundLauncher.spec"

if (-not $IsWindows -and $PSVersionTable.PSEdition -eq "Core") {
    throw "Windows에서만 빌드할 수 있습니다."
}

& $Python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller가 없습니다. 먼저 '$Python -m pip install -r requirements-build.txt'를 실행해주세요."
}

Push-Location $ProjectRoot
try {
    & $Python -m PyInstaller --noconfirm --clean $SpecPath
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 빌드가 실패했습니다."
    }
} finally {
    Pop-Location
}

$ExePath = Join-Path $ProjectRoot "dist\JjukfundLauncher\JjukfundLauncher.exe"
if (-not (Test-Path $ExePath)) {
    throw "빌드 결과를 찾지 못했습니다: $ExePath"
}
Write-Output "빌드 완료: $ExePath"

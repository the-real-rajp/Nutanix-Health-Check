$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Write-Host ""
Write-Host "Nutanix Health Check 1.0.0 - Windows portable build"
Write-Host "===================================================="
Write-Host ""

foreach ($Command in @("python", "pyinstaller", "node", "npm")) {
    if (-not (Get-Command $Command -ErrorAction SilentlyContinue)) {
        throw "Required command '$Command' was not found."
    }
}

Write-Host "[1/6] Validating application..."
python -m py_compile nutanix_health_check.py
python nutanix_health_check.py --version

Write-Host "[2/6] Preparing bundled Node.js..."
New-Item -ItemType Directory -Force "vendor\node" | Out-Null
New-Item -ItemType Directory -Force "vendor\node-runtime" | Out-Null
Copy-Item (Get-Command node.exe).Source "vendor\node\node.exe" -Force

Write-Host "[3/6] Installing the pinned Word report library..."
npm install `
    --prefix "vendor\node-runtime" `
    --no-save `
    --no-package-lock `
    "docx@9.7.1"

Write-Host "[4/6] Building the Windows application..."
Remove-Item "build", "dist" -Recurse -Force -ErrorAction SilentlyContinue
pyinstaller --noconfirm --clean "NutanixHealthCheck.spec"

$PackageDir = Join-Path $ProjectRoot "dist\Nutanix-Health-Check-1.0.0-Windows-x64"
$ZipPath = Join-Path $ProjectRoot "dist\Nutanix-Health-Check-1.0.0-Windows-x64.zip"

Write-Host "[5/6] Adding launcher and output directories..."
Copy-Item "Run-Nutanix-Health-Check.cmd" $PackageDir -Force
New-Item -ItemType Directory -Force (Join-Path $PackageDir "output\logs") | Out-Null

Write-Host "[6/6] Creating portable ZIP..."
Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $PackageDir "*") -DestinationPath $ZipPath -Force

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $ZipPath"
Write-Host ""
Write-Host "Extract the ZIP on a Windows test computer and run:"
Write-Host "  Run-Nutanix-Health-Check.cmd"

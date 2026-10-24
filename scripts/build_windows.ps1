# Build a portable Windows directory; keep all generated files outside the repository.
param(
    [string]$Python = "python",
    [string]$OutputDirectory = (Join-Path $env:TEMP "vcard2qr-build")
)
$ErrorActionPreference = "Stop"
$projectRoot = Split-Path $PSScriptRoot -Parent
$buildRoot = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
$pythonExecutable = (Get-Command $Python -ErrorAction Stop).Source
$pythonBase = & $pythonExecutable -c "import sys; print(sys.base_prefix)"
$originalPath = $env:PATH

# Use the invoking environment's pinned PyInstaller and runtime dependencies.
Push-Location $projectRoot
try {
    # Prevent unrelated tools on PATH from supplying incompatible ICU/UCRT DLLs.
    $env:PATH = "$(Split-Path $pythonExecutable);$pythonBase;$env:SystemRoot\System32;$env:SystemRoot"
    & $pythonExecutable -m PyInstaller --clean --noconfirm --windowed --onedir --name vcard2qr `
        --copy-metadata PySide6-Essentials --copy-metadata shiboken6 `
        --copy-metadata Pillow --copy-metadata qrcode --copy-metadata openpyxl `
        --copy-metadata et-xmlfile --copy-metadata colorama `
        --distpath (Join-Path $buildRoot "dist") `
        --workpath (Join-Path $buildRoot "work") `
        --specpath $buildRoot "vcard-2-qr.py"
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

    # Ship usage and dependency notices alongside the executable.
    $appFolder = Join-Path $buildRoot "dist/vcard2qr"
    Copy-Item -LiteralPath (Join-Path $projectRoot "README.md") -Destination $appFolder
    Copy-Item -LiteralPath (Join-Path $projectRoot "THIRD_PARTY_NOTICES.md") -Destination $appFolder
    Copy-Item -LiteralPath (Join-Path $projectRoot "sample_contacts.xlsx") -Destination $appFolder
    Copy-Item -LiteralPath (Join-Path $projectRoot "docs") -Destination $appFolder -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $projectRoot "CHANGELOG.md") -Destination $appFolder
    $pythonLicense = & $pythonExecutable -c "import sys; from pathlib import Path; print(Path(sys.base_prefix) / 'LICENSE.txt')"
    if (Test-Path -LiteralPath $pythonLicense) {
        Copy-Item -LiteralPath $pythonLicense -Destination (Join-Path $appFolder "PYTHON-LICENSE.txt")
    }
    Write-Output "Portable application: $appFolder"
} finally {
    $env:PATH = $originalPath
    Pop-Location
}

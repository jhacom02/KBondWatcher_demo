# Nuitka onefolder Pilot build (Windows)
# Requires: pip install nuitka ordered-set zstandard cryptography
#
# Usage (from repo root, with venv active):
#   powershell -ExecutionPolicy Bypass -File build\build_nuitka.ps1
#
# Output: dist\main.dist\ (rename to KBondWatcher)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$Version = "0.2.0"
$DemoExpiry = (Get-Date).AddDays(90).ToString("yyyy-MM-dd")
$BuildFlags = Join-Path $Root "app\build_flags.py"

$OriginalFlags = Get-Content -Path $BuildFlags -Raw

@"
from __future__ import annotations

# Injected by build/build_nuitka.ps1 — frozen Pilot cannot be downgraded via env.
DEPLOY_MODE_BUILD: str | None = "pilot"
"@ | Set-Content -Path $BuildFlags -Encoding utf8

try {
  python -m pip install -U nuitka ordered-set zstandard cryptography

  python -m nuitka `
    --standalone `
    --assume-yes-for-downloads `
    --windows-console-mode=force `
    --enable-plugin=anti-bloat `
    --include-package=app `
    --include-package=admin `
    --include-package=config `
    --include-package=core `
    --include-package=excel `
    --include-package=send `
    --include-package=source `
    --include-package-data=app `
    --include-package-data=admin `
    --output-dir=dist `
    main.py

  $Dist = Join-Path $Root "dist\main.dist"
  if (-not (Test-Path $Dist)) {
    throw "Nuitka output folder missing: $Dist"
  }

  Set-Content -Path (Join-Path $Dist "VERSION.txt") -Value $Version -Encoding ascii
  Set-Content -Path (Join-Path $Dist "demo_expiry.txt") -Value $DemoExpiry -Encoding ascii

  $pyLeft = Get-ChildItem -Path $Dist -Recurse -Filter *.py -ErrorAction SilentlyContinue
  if ($pyLeft) {
    throw ("Post-build .py sources found (refuse ship): " + ($pyLeft.FullName -join ", "))
  }

  Write-Host "Build done: $Dist"
  Write-Host "VERSION=$Version demo_expiry=$DemoExpiry DEPLOY_MODE_BUILD=pilot"
  Write-Host "Optional Authenticode: signtool sign /fd SHA256 /a dist\main.dist\main.exe"
  Write-Host "Trader: main.exe --serve   Admin: main.exe --serve-admin"
}
finally {
  Set-Content -Path $BuildFlags -Value $OriginalFlags -Encoding utf8
}

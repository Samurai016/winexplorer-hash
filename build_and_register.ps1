# Ensure you run this as Administrator!

Write-Host "Building Rust Property Handler..." -ForegroundColor Cyan
cargo build --release

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed. Please ensure Rust is installed (rustup.rs)." -ForegroundColor Red
    exit 1
}

$dllPath = Resolve-Path ".\target\release\explorer_hash.dll"
Write-Host "Copying schema file to release directory..." -ForegroundColor Cyan
Copy-Item ".\hash_schema.propdesc" -Destination ".\target\release\" -Force

Write-Host "Registering DLL at: $dllPath" -ForegroundColor Cyan

# regsvr32 /s registers silently, /c is for console output (optional)
Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s `"$dllPath`"" -Verb RunAs -Wait

Write-Host "Registered! Please restart explorer.exe if you don't see the column." -ForegroundColor Green
Write-Host "To view the column, right-click Explorer headers -> More -> 'File Hash (MD5)'" -ForegroundColor Yellow

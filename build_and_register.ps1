# Ensure you run this as Administrator!

Write-Host "Building Rust Property Handler..." -ForegroundColor Cyan
cargo build --release

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed. Please ensure Rust is installed (rustup.rs)." -ForegroundColor Red
    exit 1
}

Write-Host "Copying schema file to release directory..." -ForegroundColor Cyan
Copy-Item ".\hash_schema.propdesc" -Destination ".\target\release\" -Force

Push-Location ".\target\release"
try {
    & "..\..\register.ps1"
} finally {
    Pop-Location
}

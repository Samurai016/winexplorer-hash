Write-Host "Building release DLL..." -ForegroundColor Cyan
cargo build --release

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed! Make sure Rust is installed." -ForegroundColor Red
    exit 1
}

$dllPath = ".\target\release\explorer_hash.dll"
if (-not (Test-Path $dllPath)) {
    Write-Host "Error: DLL not found at $dllPath" -ForegroundColor Red
    exit 1
}

$zipName = "WinExplorerHash_Release.zip"

Write-Host "Creating zip archive $zipName..." -ForegroundColor Cyan

# Specify the files to include
$filesToZip = @(
    $dllPath,
    ".\hash_schema.propdesc",
    ".\register.ps1",
    ".\README.md"
)

# Compress-Archive handles the creation
Compress-Archive -Path $filesToZip -DestinationPath ".\$zipName" -Force

Write-Host "Release created successfully: $zipName" -ForegroundColor Green

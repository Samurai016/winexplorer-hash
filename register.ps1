# Ensure you run this as Administrator!

$dllPath = Resolve-Path ".\explorer_hash.dll" -ErrorAction SilentlyContinue

if (-not $dllPath) {
    Write-Host "Error: explorer_hash.dll not found in the current directory." -ForegroundColor Red
    Write-Host "Make sure you extracted all files from the release zip." -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path ".\hash_schema.propdesc")) {
    Write-Host "Warning: hash_schema.propdesc not found. Schema registration may fail." -ForegroundColor Yellow
}

Write-Host "Registering DLL at: $dllPath" -ForegroundColor Cyan

# regsvr32 /s registers silently
Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s `"$dllPath`"" -Verb RunAs -Wait

Write-Host "Registered! Please restart explorer.exe if you don't see the column." -ForegroundColor Green
Write-Host "To view the column, right-click Explorer headers -> More -> 'File Hash (MD5)'" -ForegroundColor Yellow

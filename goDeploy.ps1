# PDFusion Deployment Script
# Builds and publishes the application as a self-contained executable

param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64"
)

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  PDFusion Build & Publish" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Check for running processes
Write-Host "Checking for running instances..." -ForegroundColor Yellow
$process = Get-Process -Name "PDFusion" -ErrorAction SilentlyContinue
if ($process) {
    Write-Host "PDFusion is currently running. Please close it first!" -ForegroundColor Red
    $close = Read-Host "Press Y to force close, or any other key to exit"
    if ($close -eq "Y" -or $close -eq "y") {
        Stop-Process -Name "PDFusion" -Force
        Start-Sleep -Seconds 2
        Write-Host "Process closed" -ForegroundColor Green
    } else {
        exit 1
    }
}
Write-Host ""

# Clean previous builds (skip if locked)
Write-Host "[1/4] Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "bin") {
    try {
        Remove-Item -Path "bin" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Clean complete" -ForegroundColor Green
    } catch {
        Write-Host "Skipping clean (files in use - not a problem)" -ForegroundColor Yellow
    }
}
if (Test-Path "obj") {
    try {
        Remove-Item -Path "obj" -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        # Ignore
    }
}
Write-Host ""

# Restore dependencies
Write-Host "[2/4] Restoring NuGet packages..." -ForegroundColor Yellow
dotnet restore PDFusion.csproj
if ($LASTEXITCODE -ne 0) {
    Write-Host "Restore failed" -ForegroundColor Red
    exit 1
}
Write-Host "Restore complete" -ForegroundColor Green
Write-Host ""

# Build
Write-Host "[3/4] Building project..." -ForegroundColor Yellow
dotnet build PDFusion.csproj -c $Configuration
if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed" -ForegroundColor Red
    exit 1
}
Write-Host "Build complete" -ForegroundColor Green
Write-Host ""

# Publish as self-contained
Write-Host "[4/4] Publishing self-contained executable..." -ForegroundColor Yellow
dotnet publish PDFusion.csproj -c $Configuration -r $Runtime --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
if ($LASTEXITCODE -ne 0) {
    Write-Host "Publish failed" -ForegroundColor Red
    exit 1
}
Write-Host "Publish complete" -ForegroundColor Green
Write-Host ""

# Copy executable to Desktop
$publishPath = "bin\$Configuration\net9.0-windows\$Runtime\publish"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$exeName = "PDFusion.exe"

Write-Host ""
Write-Host "Copying executable to Desktop..." -ForegroundColor Yellow

try {
    Copy-Item -Path "$publishPath\$exeName" -Destination "$desktopPath\$exeName" -Force
    Write-Host "Copied to Desktop successfully" -ForegroundColor Green
} catch {
    Write-Host "Failed to copy to Desktop: $_" -ForegroundColor Red
}

# Also copy the data folder next to the executable
Write-Host "Copying data folder..." -ForegroundColor Yellow
try {
    if (Test-Path "$publishPath\data") {
        Remove-Item -Path "$publishPath\data" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -Path "data" -Destination "$publishPath\data" -Recurse -Force
    Write-Host "Data folder copied" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not copy data folder: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Deployment Complete!" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Executable locations:" -ForegroundColor Yellow
Write-Host "  Desktop: $desktopPath\$exeName" -ForegroundColor White
Write-Host "  Source:  $publishPath\$exeName" -ForegroundColor White
Write-Host ""
Write-Host "IMPORTANT: Copy the 'data' folder next to PDFusion.exe!" -ForegroundColor Red
Write-Host "  From: $publishPath\data" -ForegroundColor Yellow
Write-Host "  To:   Same folder as PDFusion.exe" -ForegroundColor Yellow
Write-Host ""
Write-Host "You can now run PDFusion.exe from your Desktop!" -ForegroundColor Green
Write-Host ""

# Offer to open Desktop folder
$openFolder = Read-Host "Open Desktop folder? (Y/N)"
if ($openFolder -eq "Y" -or $openFolder -eq "y") {
    Start-Process explorer.exe -ArgumentList $desktopPath
}
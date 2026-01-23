# PDFusion Deployment Script
# Builds and publishes the application as a self-contained executable
# Supports Windows, macOS (Intel & Apple Silicon), and Linux

param(
    [string]$Configuration = "Release",
    [ValidateSet("win-x64", "osx-x64", "osx-arm64", "linux-x64", "current")]
    [string]$Runtime = "current"
)

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  PDFusion Cross-Platform Build" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Determine current platform if "current" is specified
if ($Runtime -eq "current") {
    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        $Runtime = "win-x64"
    } elseif ($IsMacOS) {
        # Check for Apple Silicon
        $arch = uname -m
        if ($arch -eq "arm64") {
            $Runtime = "osx-arm64"
        } else {
            $Runtime = "osx-x64"
        }
    } elseif ($IsLinux) {
        $Runtime = "linux-x64"
    } else {
        Write-Host "Could not detect platform, defaulting to current OS" -ForegroundColor Yellow
        $Runtime = "win-x64"
    }
}

Write-Host "Building for: $Runtime" -ForegroundColor Yellow
Write-Host ""

# Check for running processes (Windows only)
if ($Runtime -eq "win-x64") {
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

# Determine paths based on runtime
$publishPath = "bin/$Configuration/net10.0/$Runtime/publish"

# Set executable name based on platform
if ($Runtime -eq "win-x64") {
    $exeName = "PDFusion.exe"
} else {
    $exeName = "PDFusion"
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Build Complete!" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Executable location:" -ForegroundColor Yellow
Write-Host "  $publishPath/$exeName" -ForegroundColor White
Write-Host ""

# Platform-specific post-build steps
if ($Runtime -eq "win-x64") {
    # Copy to Desktop on Windows
    $desktopPath = [Environment]::GetFolderPath("Desktop")

    Write-Host "Copying executable to Desktop..." -ForegroundColor Yellow
    try {
        Copy-Item -Path "$publishPath/$exeName" -Destination "$desktopPath/$exeName" -Force
        Write-Host "Copied to Desktop successfully" -ForegroundColor Green
        Write-Host "  Desktop: $desktopPath/$exeName" -ForegroundColor White
    } catch {
        Write-Host "Failed to copy to Desktop: $_" -ForegroundColor Red
    }

    # Offer to open Desktop folder
    $openFolder = Read-Host "Open Desktop folder? (Y/N)"
    if ($openFolder -eq "Y" -or $openFolder -eq "y") {
        Start-Process explorer.exe -ArgumentList $desktopPath
    }
} elseif ($Runtime -like "osx-*") {
    Write-Host "macOS Notes:" -ForegroundColor Yellow
    Write-Host "  - You may need to allow the app in System Preferences > Security" -ForegroundColor White
    Write-Host "  - To run: ./$publishPath/$exeName" -ForegroundColor White
    Write-Host "  - Or double-click the executable in Finder" -ForegroundColor White
    Write-Host ""

    # Make executable on macOS
    if ($IsMacOS) {
        chmod +x "$publishPath/$exeName"
        Write-Host "Made executable with chmod +x" -ForegroundColor Green
    }
} elseif ($Runtime -eq "linux-x64") {
    Write-Host "Linux Notes:" -ForegroundColor Yellow
    Write-Host "  - To run: ./$publishPath/$exeName" -ForegroundColor White
    Write-Host ""

    # Make executable on Linux
    if ($IsLinux) {
        chmod +x "$publishPath/$exeName"
        Write-Host "Made executable with chmod +x" -ForegroundColor Green
    }
}

# Copy data folder if it exists
if (Test-Path "data") {
    Write-Host ""
    Write-Host "Copying data folder..." -ForegroundColor Yellow
    try {
        if (Test-Path "$publishPath/data") {
            Remove-Item -Path "$publishPath/data" -Recurse -Force -ErrorAction SilentlyContinue
        }
        Copy-Item -Path "data" -Destination "$publishPath/data" -Recurse -Force
        Write-Host "Data folder copied" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not copy data folder: $_" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "You can now run PDFusion!" -ForegroundColor Green
Write-Host ""

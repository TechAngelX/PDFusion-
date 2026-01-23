# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PDFusion is a Windows Forms desktop application (.NET 9.0) for batch renaming student PDF application forms using CSV or Excel data. It matches PDF filenames to student records and generates standardized filenames.

## Build Commands

```powershell
# Full build and publish (recommended)
.\goDeploy.ps1

# Manual build steps
dotnet restore PDFusion.csproj
dotnet build PDFusion.csproj -c Release
dotnet publish PDFusion.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

Output: `bin\Release\net9.0-windows\win-x64\publish\PDFusion.exe`

## Architecture

### File Structure
- `MainForm.cs` - Main application logic, UI initialization, file processing workflow
- `UIComponents.cs` - Custom UI components (`ModernFilePanel`, `ModernButton`) and data models (`StudentRecord`, `RenameOperation`)
- `Program.cs` - Application entry point

### Data Flow
1. User loads CSV/Excel file → parsed into `List<StudentRecord>`
2. User selects PDF folder → files scanned for student number pattern
3. Preview matches PDFs to student records using `StudentNo` as key
4. Apply copies files to "Renamed" subfolder with new names (originals unchanged)

### Key Functions in MainForm.cs
- `LoadFromCSV()` / `LoadFromExcel()` - Parse data files, map column names flexibly
- `ExtractStudentNumber()` - Extracts 7-10 digit number from PDF filename prefix (before first `-`)
- `GenerateNewFilename()` - Creates output format: `b{batch} {Forename} {Surname} {StudentNo} {FeeCode} {Grade}.pdf`
- `DetermineFeeStatusCode()` - Maps "home" → "H", "overseas" → "OH"
- `FormatUKGrade()` - Normalizes grades (e.g., "2.1" → "2_1")

### Dependencies
- EPPlus 7.5.2 - Excel file reading (requires `ExcelPackage.LicenseContext = LicenseContext.NonCommercial`)
- Windows Forms (.NET 9.0)

## Input/Output Formats

**PDF input pattern:** `XXXXXXXX-XX-XX-OVERVIEW.PDF` (8-digit student number prefix)

**Required CSV/Excel columns:** StudentNo, Batch, Forename, Surname, FeeStatus, UKGrade

**Output filename:** `b1 John Smith 26049530 H 2_1.pdf`

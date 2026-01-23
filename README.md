# PDFusion

**Data-driven document reorganisation tool**

A cross-platform desktop application for batch renaming student PDF application forms using CSV or Excel data. Built for admissions offices and educational institutions.

![PDFusion Screenshot](readme_images/screenshot1.png)

---

## Features

- Smart data loading from CSV or Excel files
- Batch process hundreds of PDFs in seconds
- Preview all changes before applying
- Renamed files saved to separate subfolder
- Undo support for last operation
- Cross-platform support (Windows, macOS, Linux)
- Modern Avalonia UI interface
- Automatic proper case formatting and status codes

---

## Use Case

Designed for office administrators who need to:
- Organize large batches of student application PDFs
- Rename files based on spreadsheet data
- Maintain consistent file naming conventions
- Process applications efficiently

---

## How It Works

### Input Format

**PDF Files:**
```
26034824-01-01-OVERVIEW.PDF
26043635-02-01-OVERVIEW.PDF
26049530-01-01-OVERVIEW.PDF
```

**CSV/Excel Data:**

Required columns:
- `StudentNo` - Student number (matches PDF filename prefix)
- `Batch` - Batch identifier (e.g., "Batch 1", "Batch 2")
- `Forename` - Student's first name
- `Surname` - Student's last name
- `FeeStatus` - Fee classification (Home, Overseas, or European)
- `UKGrade` - UK grade classification (e.g., "2.1", "1", "2.2", "3")

### Output Format

```
b1 John Smith 26049530 H 2_1.pdf
b1 Jane Doe 26038447 H 1.pdf
b1 Alex Johnson 26045428 OS 2_1.pdf
```

**Format breakdown:**
- `b1` - Batch number
- `John Smith` - Proper case name (converted from ALL CAPS)
- `26049530` - Student number
- `H` - Home fee status (`OS` = Overseas/European)
- `2_1` - UK grade classification

---

## Quick Start

### Usage

1. Launch PDFusion
2. **Step 1:** Click or drag your CSV/Excel file with student data
3. **Step 2:** Choose the folder containing PDF files to rename
4. Click **Preview Renames** to see all proposed changes
5. Click **Apply Renames** to process files
6. Renamed files appear in the "Renamed" subfolder

Original PDF files remain untouched in the source folder.

---

## Installation

### Option A: Download Release (Recommended)
1. Download the appropriate release for your platform:
   - Windows: `PDFusion-win-x64.zip`
   - macOS (Apple Silicon): `PDFusion-osx-arm64.zip`
   - macOS (Intel): `PDFusion-osx-x64.zip`
   - Linux: `PDFusion-linux-x64.zip`
2. Extract and run (no installation required)

### Option B: Build from Source
See [Building from Source](#building-from-source) below.

---

## System Requirements

| Platform | Requirements |
|----------|-------------|
| **Windows** | Windows 10/11 (64-bit) |
| **macOS** | macOS 11+ (Intel or Apple Silicon) |
| **Linux** | Ubuntu 20.04+, Fedora 33+, or equivalent |

- **Memory:** Minimum 100MB RAM
- **Storage:** 50MB disk space
- **Runtime:** Included in self-contained builds

---

## Project Structure

```
PDFusion/
├── readme_images/
│   └── screenshot1.png      # Application screenshot
├── data/                     # Sample data folder
├── App.axaml                # Avalonia application definition
├── App.axaml.cs             # Application startup
├── MainWindow.axaml         # Main window UI layout
├── MainWindow.axaml.cs      # Main application logic
├── Models.cs                # Data models
├── Styles.axaml             # Custom UI styles
├── Program.cs               # Application entry point
├── PDFusion.csproj          # Project configuration
├── goDeploy.ps1             # Build & deployment script
└── README.md                # This file
```

---

## Smart Features

### Automatic Name Formatting
Converts names from ALL CAPS to proper case:
- `JOHN SMITH` → `John Smith`
- `MARIA GARCIA` → `Maria Garcia`
- `DAVID CHEN` → `David Chen`

### Fee Status Codes
- `HOME` or `HOME (PROVISIONAL)` → `H`
- `OVERSEAS` or `OVERSEAS STUDENTS` → `OS`
- `EUROPEAN` or `EUROPEAN (PROVISIONAL)` → `OS`

### UK Grade Formatting
- `2.1` or `2:1` → `2_1`
- `First Class` or `1st` → `1`
- `2.2` or `2:2` → `2_2`
- `3rd` or `3.0` → `3`

### Safety Features
- Original PDFs never modified
- Renamed files saved to separate folder
- Preview all changes before applying
- Undo last batch operation
- Detailed status logging
- Handles duplicate filenames gracefully

---

## Troubleshooting

### Common Issues

**"Could not find StudentNo column"**
- Ensure your CSV/Excel has a column named exactly `StudentNo`
- Check for extra spaces or special characters in header names

**"Student XXXXXXXX not found in data"**
- The PDF's student number doesn't exist in your data file
- Verify student numbers match between PDFs and spreadsheet

**"Could not extract student number"**
- PDF filename doesn't match expected format
- Required format: `XXXXXXXX-XX-XX-OVERVIEW.PDF` (8 digits at start)

**"Target already exists"**
- A file with the new name already exists in the Renamed folder
- Check for duplicate student records in your data

### macOS Security
On first launch, macOS may block the app. To allow it:
1. Go to **System Preferences > Security & Privacy**
2. Click **Open Anyway** for PDFusion

---

## Building from Source

### Prerequisites
- .NET 10.0 SDK (or .NET 9.0+)

### Quick Build & Run

```bash
# Clone repository
git clone https://github.com/yourusername/pdfusion.git
cd pdfusion

# Run directly
dotnet run

# Build release
dotnet build -c Release
```

### macOS Build (Recommended)

Use the included shell script to build and copy to Desktop:

```bash
# Make executable (first time only)
chmod +x build-mac.sh

# Build and publish to Desktop
./build-mac.sh
```

This automatically:
- Detects your Mac architecture (Intel or Apple Silicon)
- Builds a self-contained single-file executable
- Copies the app to `~/Desktop/PDFusion/`

### Windows Build

Use PowerShell:

```powershell
# Build for current platform
.\goDeploy.ps1

# Or specify runtime explicitly
.\goDeploy.ps1 -Runtime win-x64
```

### Manual Cross-Platform Publishing

```bash
# Windows (64-bit)
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o ~/Desktop/PDFusion

# macOS (Apple Silicon)
dotnet publish -c Release -r osx-arm64 --self-contained -p:PublishSingleFile=true -o ~/Desktop/PDFusion

# macOS (Intel)
dotnet publish -c Release -r osx-x64 --self-contained -p:PublishSingleFile=true -o ~/Desktop/PDFusion

# Linux (64-bit)
dotnet publish -c Release -r linux-x64 --self-contained -p:PublishSingleFile=true -o ~/Desktop/PDFusion
```

### Build Scripts

| Platform | Script | Usage |
|----------|--------|-------|
| macOS/Linux | `build-mac.sh` | `./build-mac.sh` |
| Windows | `goDeploy.ps1` | `.\goDeploy.ps1` |

### Output Locations

| Platform | Executable Path |
|----------|-----------------|
| Windows | `~/Desktop/PDFusion/PDFusion.exe` |
| macOS | `~/Desktop/PDFusion/PDFusion` |
| Linux | `~/Desktop/PDFusion/PDFusion` |

---

## License

© 2025 Ricki Angel

Free to use for educational and personal purposes.

---

## Contributing

Contributions are welcome. Please submit a Pull Request.

---

## Acknowledgments

**Built with:**
- .NET 10.0 - Runtime
- Avalonia UI 11.2 - Cross-platform UI framework
- EPPlus 7.5.2 - Excel file processing
- C# - Programming language

---

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

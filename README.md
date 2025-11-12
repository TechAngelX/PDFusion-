# Batch Renamer

A Windows Forms application for batch renaming PDF application forms using student data from CSV or Excel files.

## Features

- **Loads student data** from CSV or Excel files (any filename, from desktop)
- **Automatically matches** PDF files by student number
- **Renames PDFs** in format: `b[Batch] [Forename] [Surname] [StudentNo] [H/O] [Grade].pdf`
- **Preview before renaming** - see what will change before committing
- **Undo functionality** - restore original filenames if needed
- **Modern UI** with drag-and-drop support

## Renaming Format

Original: `26034824-01-01-OVERVIEW.PDF`

New: `b1 Lisa Jackson 26034824 H 2_1.pdf`

Where:
- **b1** = Batch number (extracted from "Batch 1", "Batch 2", etc.)
- **Lisa Jackson** = Forename and Surname from CSV
- **26034824** = Student number
- **H** = Home student (or **O** for Overseas)
- **2_1** = UK Grade (2.1, 1, 2.2, 3)

## Required CSV/Excel Columns

Your data file must contain these columns (names must match exactly):

- `StudentNo` - Student number (must match PDF filename)
- `Batch` - Batch information (e.g., "Batch 1", "Batch 2")
- `Forename` - Student's first name
- `Surname` - Student's last name
- `FeeStatus` - Must contain "Home" or "Overseas"
- `UKGrade` - UK grade classification (e.g., "2.1", "1", "2.2", "3")

## How to Use

1. **Load Student Data**
   - Click or drag CSV/Excel file into Step 1 panel
   - File can be named anything, but must have correct column headers
   - Application will validate and load student records

2. **Select PDF Folder**
   - Click Step 2 panel to browse for folder containing PDFs
   - PDFs should be named like: `XXXXXXXX-XX-XX-OVERVIEW.PDF`
   - Where XXXXXXXX is the student number

3. **Preview Renames**
   - Click "Preview Renames" button
   - Review the proposed changes in the preview list
   - Green rows = successful match
   - Red rows = student not found in data
   - Gray rows = could not extract student number

4. **Apply Renames**
   - Click "Apply Renames" button
   - Confirm the operation
   - Files will be renamed
   - Status log shows progress

5. **Undo (if needed)**
   - Click "Undo Last Batch" to restore original filenames
   - Only works for the most recent rename operation

## Building the Application

### Requirements
- .NET 9.0 SDK or later
- Windows OS

### Build Instructions

```bash
# Restore dependencies
dotnet restore

# Build the application
dotnet build -c Release

# Run the application
dotnet run

# Or publish as single executable
dotnet publish -c Release -r win-x64 --self-contained
```

The built executable will be in: `bin/Release/net9.0-windows/win-x64/publish/BatchRenamer.exe`

## Project Structure

```
BatchRenamer/
├── MainForm.cs          # Main application form and logic
├── UIComponents.cs      # Custom UI components (panels, buttons)
├── Program.cs           # Application entry point
├── BatchRenamer.csproj  # Project configuration
└── README.md           # This file
```

## Dependencies

- **EPPlus 7.5.2** - For reading Excel files (.xlsx, .xls)
- **.NET 9.0 Windows Forms** - UI framework

## Notes

- PDF files not matching any student number will be skipped
- Duplicate filenames will be detected and skipped during rename
- All operations are logged to the status window
- Undo only works for the most recent batch of renames

## Error Handling

The application handles common errors:
- Missing or invalid CSV/Excel columns
- PDF files with unexpected naming format
- File access permission issues
- Duplicate target filenames

## License

Free to use for educational and personal purposes.

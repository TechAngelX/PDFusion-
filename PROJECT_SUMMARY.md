# BatchRenamer - Project Summary

## ✅ Project Complete!

Your BatchRenamer application is ready. All files have been created based on your ADMerger template with the new renaming logic.

## 📦 What You Got

### Core Application Files
- **MainForm.cs** - Main application logic with CSV/Excel loading and renaming
- **UIComponents.cs** - Modern UI components (panels, buttons)
- **Program.cs** - Application entry point
- **BatchRenamer.csproj** - Project configuration with EPPlus dependency

### Documentation
- **README.md** - Full documentation
- **QUICKSTART.md** - Quick start guide
- **.gitignore** - Git ignore rules

### Build Script
- **goDeploy.ps1** - PowerShell script to build and publish

## 🎯 Renaming Logic Implemented

```
INPUT:  26034824-01-01-OVERVIEW.PDF
OUTPUT: b1 Lisa Jackson 26034824 H 2_1.pdf
```

### Format Breakdown:
- **b1** ← Extracted from "Batch 1" column
- **Lisa Jackson** ← Forename + Surname columns
- **26034824** ← StudentNo column (matched from PDF filename)
- **H** ← FeeStatus column ("Home" → H, "Overseas" → O)
- **2_1** ← UKGrade column (formatted: 2.1 → 2_1)

## 🚀 Next Steps

### 1. Copy Files to Your Project
Copy all files from this output to:
```
C:\Users\Ricki\Documents\LOCALDEV-PC\BRenamer\
```

### 2. Build the Application
Open PowerShell in that directory and run:
```powershell
.\goDeploy.ps1
```

This will:
- Clean previous builds
- Restore NuGet packages
- Build the project
- Publish as a self-contained executable

### 3. Find Your Executable
After building, you'll find it at:
```
bin\Release\net9.0-windows\win-x64\publish\BatchRenamer.exe
```

### 4. Test It!
1. Create a test CSV with the required columns
2. Put some test PDFs in a folder (format: XXXXXXXX-XX-XX-OVERVIEW.PDF)
3. Run BatchRenamer.exe
4. Load your CSV
5. Select the PDF folder
6. Preview and apply!

## 📋 Required CSV Columns

Your data file MUST have these exact column names:
- `StudentNo`
- `Batch`
- `Forename`
- `Surname`
- `FeeStatus`
- `UKGrade`

## 🎨 UI Features

✅ Modern purple gradient header  
✅ Drag-and-drop file loading  
✅ Real-time preview of renames  
✅ Detailed status logging  
✅ Undo functionality  
✅ Color-coded results (green = match, red = error)  

## 🔧 Technical Details

### Dependencies
- **.NET 9.0** - Target framework
- **EPPlus 7.5.2** - Excel file reading
- **Windows Forms** - UI framework

### Key Features
- Reads CSV and Excel (.xlsx, .xls) files
- Extracts student number from PDF filenames
- Matches with student data
- Generates new filenames with proper formatting
- Previews all changes before applying
- Full undo support for last operation
- Comprehensive error handling

## 💡 Code Highlights

### Student Number Extraction
```csharp
// Extracts "26034824" from "26034824-01-01-OVERVIEW.PDF"
private string ExtractStudentNumber(string filename)
{
    var parts = filename.Split('-');
    string firstPart = parts[0].Trim();
    if (firstPart.All(char.IsDigit) && firstPart.Length >= 7)
        return firstPart;
    return null;
}
```

### Filename Generation
```csharp
// Format: b[Batch] [Forename] [Surname] [StudentNo] [H/O] [Grade].pdf
private string GenerateNewFilename(StudentRecord student)
{
    string batchNum = ExtractBatchNumber(student.Batch);
    string feeCode = DetermineFeeStatusCode(student.FeeStatus);
    string gradeCode = FormatUKGrade(student.UKGrade);
    
    return $"b{batchNum} {student.Forename} {student.Surname} " +
           $"{student.StudentNo} {feeCode} {gradeCode}.pdf";
}
```

## 🎉 Done!

Your BatchRenamer is complete and ready to use. The UI matches your ADMerger style with a fresh purple theme, and all the renaming logic is implemented exactly as specified.

Good luck with your batch renaming! 🚀

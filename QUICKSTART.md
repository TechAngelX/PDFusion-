# BatchRenamer - Quick Start Guide

## What This Does

Renames PDF files like `26034824-01-01-OVERVIEW.PDF` to `b1 Lisa Jackson 26034824 H 2_1.pdf`

## Setup (One Time)

1. **Copy all files** to your project folder:
   ```
   C:\Users\Ricki\Documents\LOCALDEV-PC\BRenamer\
   ```

2. **Open PowerShell** in that folder and run:
   ```powershell
   .\goDeploy.ps1
   ```

3. **Find the executable** at:
   ```
   bin\Release\net9.0-windows\win-x64\publish\BatchRenamer.exe
   ```

## How to Use

### Step 1: Prepare Your Data File
Make sure your CSV or Excel file has these columns:
- **StudentNo** - e.g., 26034824
- **Batch** - e.g., "Batch 1", "Batch 2"
- **Forename** - e.g., "Lisa"
- **Surname** - e.g., "Jackson"
- **FeeStatus** - Must contain "Home" or "Overseas"
- **UKGrade** - e.g., "2.1", "1", "2.2", "3"

### Step 2: Run BatchRenamer.exe

1. **Load your data file** (Step 1)
   - Click or drag your CSV/Excel file
   - Application will show how many records loaded

2. **Select PDF folder** (Step 2)
   - Click to browse for the folder with your PDFs
   - Shows how many PDF files found

3. **Preview renames**
   - Click "Preview Renames"
   - Check the list - green = will rename, red = no match

4. **Apply renames**
   - Click "Apply Renames"
   - Confirm when prompted
   - Done!

5. **Undo if needed**
   - Click "Undo Last Batch" to restore original names

## Example

**Before:**
```
26034824-01-01-OVERVIEW.PDF
26043635-02-01-OVERVIEW.PDF
26051392-02-01-OVERVIEW.PDF
```

**After:**
```
b1 Lisa Jackson 26034824 H 2_1.pdf
b2 Maxwell Rahn 26043635 O 1.pdf
b2 Yufei Lin 26051392 H 2_2.pdf
```

## Troubleshooting

**"Could not find StudentNo column"**
- Check your CSV/Excel has a column named exactly "StudentNo"

**"Student XXXXXXXX not found"**
- The student number from the PDF doesn't exist in your data file
- Check for typos or missing records

**"Could not extract student number"**
- PDF filename doesn't match expected format
- Should be: `XXXXXXXX-XX-XX-OVERVIEW.PDF`

## Tips

✓ Always preview before applying  
✓ Keep a backup of original files  
✓ Use Undo feature if something goes wrong  
✓ Check the Status Log for details  

## Need Help?

Check the full README.md for more detailed information.

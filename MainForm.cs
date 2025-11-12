using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Globalization;

namespace PDFusion
{
    public partial class MainForm : Form
    {
        private string dataFilePath = "";
        private string pdfFolderPath = "";
        private string outputFolderPath = "";
        private List<StudentRecord> studentData = new List<StudentRecord>();
        private List<RenameOperation> renameOperations = new List<RenameOperation>();
        
        private ModernFilePanel dataPanel;
        private ModernFilePanel folderPanel;
        private ModernButton previewButton;
        private ModernButton processButton;
        private ModernButton exitButton;
        private ModernButton undoButton;
        private RichTextBox statusBox;
        private ListView previewListView;
        
        private List<RenameOperation> completedOperations = new List<RenameOperation>();
        
        public MainForm()
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            InitializeComponent();
        }
        
        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            this.Text = "PDFusion - Student Application Manager";
            this.ClientSize = new Size(1100, 800);
            this.BackColor = Color.White;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Header Panel
            Panel headerPanel = new Panel();
            headerPanel.Location = new Point(0, 0);
            headerPanel.Size = new Size(1100, 90);
            headerPanel.BackColor = ColorTranslator.FromHtml("#8B5CF6");
            headerPanel.Paint += (s, e) =>
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(
                    headerPanel.ClientRectangle,
                    ColorTranslator.FromHtml("#8B5CF6"),
                    ColorTranslator.FromHtml("#7C3AED"),
                    LinearGradientMode.Horizontal))
                {
                    e.Graphics.FillRectangle(brush, headerPanel.ClientRectangle);
                }
            };

            Label titleLabel = new Label();
            titleLabel.Text = "PDFusion";
            titleLabel.Font = new Font("Segoe UI", 24F, FontStyle.Bold);
            titleLabel.ForeColor = Color.White;
            titleLabel.AutoSize = true;
            titleLabel.Location = new Point(30, 15);
            titleLabel.BackColor = Color.Transparent;
            headerPanel.Controls.Add(titleLabel);

            Label subtitleLabel = new Label();
            subtitleLabel.Text = "Data-driven File Management Document Renaming Tool";
            subtitleLabel.Font = new Font("Segoe UI", 11F, FontStyle.Regular);
            subtitleLabel.ForeColor = ColorTranslator.FromHtml("#E9D5FF");
            subtitleLabel.AutoSize = true;
            subtitleLabel.Location = new Point(33, 55);
            subtitleLabel.BackColor = Color.Transparent;
            headerPanel.Controls.Add(subtitleLabel);

            Label copyrightLabel = new Label();
            copyrightLabel.Text = "© 2025 Ricki Angel";
            copyrightLabel.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            copyrightLabel.ForeColor = ColorTranslator.FromHtml("#C4B5FD");
            copyrightLabel.AutoSize = true;
            copyrightLabel.Location = new Point(920, 65);
            copyrightLabel.BackColor = Color.Transparent;
            copyrightLabel.TextAlign = ContentAlignment.MiddleRight;
            headerPanel.Controls.Add(copyrightLabel);

            this.Controls.Add(headerPanel);

            int yPos = 120;

            // Data File Panel
            Label dataLabel = new Label();
            dataLabel.Text = "Step 1: Load Student Data (CSV or Excel)";
            dataLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dataLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            dataLabel.Location = new Point(30, yPos);
            dataLabel.AutoSize = true;
            this.Controls.Add(dataLabel);

            dataPanel = new ModernFilePanel("Click or drag CSV/Excel file with student data", 30, yPos + 30);
            dataPanel.Click += (s, e) => SelectDataFile();
            dataPanel.SetDropHandler(filePath => 
            {
                dataFilePath = filePath;
                LoadStudentData();
            });
            this.Controls.Add(dataPanel);

            yPos += 140;

            // PDF Folder Panel
            Label folderLabel = new Label();
            folderLabel.Text = "Step 2: Select Folder Containing PDF Files";
            folderLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            folderLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            folderLabel.Location = new Point(30, yPos);
            folderLabel.AutoSize = true;
            this.Controls.Add(folderLabel);

            folderPanel = new ModernFilePanel("Click to select folder with PDF files to rename", 30, yPos + 30);
            folderPanel.Click += (s, e) => SelectPDFFolder();
            folderPanel.Enabled = false;
            this.Controls.Add(folderPanel);

            yPos += 140;

            // Buttons
            previewButton = new ModernButton();
            previewButton.Text = "Preview Renames";
            previewButton.Location = new Point(30, yPos);
            previewButton.Size = new Size(180, 45);
            previewButton.Enabled = false;
            previewButton.Click += PreviewRenames_Click;
            previewButton.SetRounded();
            this.Controls.Add(previewButton);

            processButton = new ModernButton();
            processButton.Text = "Apply Renames";
            processButton.Location = new Point(220, yPos);
            processButton.Size = new Size(180, 45);
            processButton.Enabled = false;
            processButton.Click += ProcessRenames_Click;
            processButton.SetRounded();
            this.Controls.Add(processButton);

            yPos += 65;

            // Preview ListView
            Label previewLabel = new Label();
            previewLabel.Text = "Rename Preview";
            previewLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            previewLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            previewLabel.Location = new Point(30, yPos);
            previewLabel.AutoSize = true;
            this.Controls.Add(previewLabel);

            previewListView = new ListView();
            previewListView.Location = new Point(30, yPos + 25);
            previewListView.Size = new Size(1030, 180);
            previewListView.View = View.Details;
            previewListView.FullRowSelect = true;
            previewListView.GridLines = true;
            previewListView.Font = new Font("Consolas", 9F);
            
            previewListView.Columns.Add("Current Filename", 280);
            previewListView.Columns.Add("→", 30);
            previewListView.Columns.Add("New Filename", 380);
            previewListView.Columns.Add("Status", 320);
            
            this.Controls.Add(previewListView);

            yPos += 215;

            // Status Box
            Label statusLabel = new Label();
            statusLabel.Text = "Status Log";
            statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            statusLabel.ForeColor = ColorTranslator.FromHtml("#334155");
            statusLabel.Location = new Point(30, yPos);
            statusLabel.AutoSize = true;
            this.Controls.Add(statusLabel);

            statusBox = new RichTextBox();
            statusBox.Location = new Point(30, yPos + 25);
            statusBox.Size = new Size(1030, 100);
            statusBox.ReadOnly = true;
            statusBox.BackColor = ColorTranslator.FromHtml("#F8FAFC");
            statusBox.BorderStyle = BorderStyle.FixedSingle;
            statusBox.Font = new Font("Consolas", 9F);
            statusBox.ForeColor = ColorTranslator.FromHtml("#475569");
            statusBox.Text = "Ready. Load a CSV or Excel file with student data to begin...";
            this.Controls.Add(statusBox);

            yPos += 135;

            // Bottom Buttons
            exitButton = new ModernButton();
            exitButton.Text = "Exit";
            exitButton.Location = new Point(30, yPos);
            exitButton.Size = new Size(120, 40);
            exitButton.Click += (s, e) => Application.Exit();
            exitButton.SetSecondary();
            this.Controls.Add(exitButton);

            undoButton = new ModernButton();
            undoButton.Text = "Undo Last Batch";
            undoButton.Location = new Point(160, yPos);
            undoButton.Size = new Size(180, 40);
            undoButton.Enabled = false;
            undoButton.Click += UndoRename_Click;
            this.Controls.Add(undoButton);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void UpdateStatus(string message)
        {
            if (statusBox.InvokeRequired)
            {
                statusBox.Invoke(new Action(() => UpdateStatus(message)));
                return;
            }
            statusBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
            statusBox.SelectionStart = statusBox.Text.Length;
            statusBox.ScrollToCaret();
        }

        private void SelectDataFile()
        {
            var dialog = new OpenFileDialog();
            dialog.Title = "Select Student Data File (CSV or Excel)";
            dialog.Filter = "Data files (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|CSV files (*.csv)|*.csv|Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                dataFilePath = dialog.FileName;
                LoadStudentData();
            }
        }

        private void LoadStudentData()
        {
            try
            {
                studentData.Clear();
                string extension = Path.GetExtension(dataFilePath).ToLower();
                
                if (extension == ".csv")
                {
                    LoadFromCSV();
                }
                else if (extension == ".xlsx" || extension == ".xls")
                {
                    LoadFromExcel();
                }
                else
                {
                    throw new Exception("Unsupported file format. Please use CSV or Excel files.");
                }
                
                dataPanel.SetFileLoaded(Path.GetFileName(dataFilePath), studentData.Count);
                UpdateStatus($"✓ Loaded {studentData.Count} student records from {Path.GetFileName(dataFilePath)}");
                
                folderPanel.Enabled = true;
                folderPanel.UpdateText("Click to select folder with PDF files to rename");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading student data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"✗ ERROR: {ex.Message}");
            }
        }

        private void LoadFromCSV()
        {
            using var reader = new StreamReader(dataFilePath);
            
            // Read header
            string headerLine = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(headerLine))
                throw new Exception("CSV file is empty");
            
            var headers = headerLine.Split(',').Select(h => h.Trim().Trim('"')).ToList();
            
            // Find required columns
            int studentNoIdx = FindColumnIndex(headers, "StudentNo", "Student No");
            int batchIdx = FindColumnIndex(headers, "Batch");
            int forenameIdx = FindColumnIndex(headers, "Forename", "First Name");
            int surnameIdx = FindColumnIndex(headers, "Surname", "Last Name");
            int feeStatusIdx = FindColumnIndex(headers, "FeeStatus", "Fee Status");
            int ukGradeIdx = FindColumnIndex(headers, "UKGrade", "UK Grade");
            
            if (studentNoIdx == -1)
                throw new Exception("Could not find 'StudentNo' column");
            if (batchIdx == -1)
                throw new Exception("Could not find 'Batch' column");
            if (forenameIdx == -1)
                throw new Exception("Could not find 'Forename' column");
            if (surnameIdx == -1)
                throw new Exception("Could not find 'Surname' column");
            if (feeStatusIdx == -1)
                throw new Exception("Could not find 'FeeStatus' column");
            if (ukGradeIdx == -1)
                throw new Exception("Could not find 'UKGrade' column");
            
            // Read data
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;
                
                var values = ParseCSVLine(line);
                
                if (studentNoIdx < values.Count && !string.IsNullOrWhiteSpace(values[studentNoIdx]))
                {
                    studentData.Add(new StudentRecord
                    {
                        StudentNo = values[studentNoIdx].Trim(),
                        Batch = batchIdx < values.Count ? values[batchIdx].Trim() : "",
                        Forename = forenameIdx < values.Count ? values[forenameIdx].Trim() : "",
                        Surname = surnameIdx < values.Count ? values[surnameIdx].Trim() : "",
                        FeeStatus = feeStatusIdx < values.Count ? values[feeStatusIdx].Trim() : "",
                        UKGrade = ukGradeIdx < values.Count ? values[ukGradeIdx].Trim() : ""
                    });
                }
            }
        }

        private List<string> ParseCSVLine(string line)
        {
            var values = new List<string>();
            bool inQuotes = false;
            var currentValue = "";
            
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentValue += '"';
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    values.Add(currentValue.Trim());
                    currentValue = "";
                }
                else
                {
                    currentValue += c;
                }
            }
            
            values.Add(currentValue.Trim());
            return values;
        }

        private void LoadFromExcel()
        {
            using var package = new ExcelPackage(new FileInfo(dataFilePath));
            var worksheet = package.Workbook.Worksheets[0];
            
            if (worksheet.Dimension == null)
                throw new Exception("Excel sheet is empty");
            
            // Read headers
            var headers = new List<string>();
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                var headerCell = worksheet.Cells[1, col].Value;
                headers.Add(headerCell?.ToString()?.Trim() ?? "");
            }
            
            // Find required columns
            int studentNoCol = FindColumnIndex(headers, "StudentNo", "Student No") + 1;
            int batchCol = FindColumnIndex(headers, "Batch") + 1;
            int forenameCol = FindColumnIndex(headers, "Forename", "First Name") + 1;
            int surnameCol = FindColumnIndex(headers, "Surname", "Last Name") + 1;
            int feeStatusCol = FindColumnIndex(headers, "FeeStatus", "Fee Status") + 1;
            int ukGradeCol = FindColumnIndex(headers, "UKGrade", "UK Grade") + 1;
            
            if (studentNoCol == 0)
                throw new Exception("Could not find 'StudentNo' column");
            if (batchCol == 0)
                throw new Exception("Could not find 'Batch' column");
            if (forenameCol == 0)
                throw new Exception("Could not find 'Forename' column");
            if (surnameCol == 0)
                throw new Exception("Could not find 'Surname' column");
            if (feeStatusCol == 0)
                throw new Exception("Could not find 'FeeStatus' column");
            if (ukGradeCol == 0)
                throw new Exception("Could not find 'UKGrade' column");
            
            // Read data
            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                var studentNo = worksheet.Cells[row, studentNoCol].Value?.ToString()?.Trim();
                
                if (!string.IsNullOrWhiteSpace(studentNo))
                {
                    studentData.Add(new StudentRecord
                    {
                        StudentNo = studentNo,
                        Batch = worksheet.Cells[row, batchCol].Value?.ToString()?.Trim() ?? "",
                        Forename = worksheet.Cells[row, forenameCol].Value?.ToString()?.Trim() ?? "",
                        Surname = worksheet.Cells[row, surnameCol].Value?.ToString()?.Trim() ?? "",
                        FeeStatus = worksheet.Cells[row, feeStatusCol].Value?.ToString()?.Trim() ?? "",
                        UKGrade = worksheet.Cells[row, ukGradeCol].Value?.ToString()?.Trim() ?? ""
                    });
                }
            }
        }

        private int FindColumnIndex(List<string> headers, params string[] possibleNames)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                string header = headers[i].Replace(" ", "").Replace("_", "").ToLower();
                
                foreach (var name in possibleNames)
                {
                    string cleanName = name.Replace(" ", "").Replace("_", "").ToLower();
                    if (header == cleanName)
                        return i;
                }
            }
            return -1;
        }

        private void SelectPDFFolder()
        {
            if (studentData.Count == 0)
            {
                MessageBox.Show("Please load student data first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new FolderBrowserDialog();
            dialog.Description = "Select folder containing PDF files to rename";
            dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                pdfFolderPath = dialog.SelectedPath;
                
                var pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                
                folderPanel.SetFileLoaded(Path.GetFileName(pdfFolderPath) + " folder", pdfFiles.Length);
                UpdateStatus($"✓ Selected folder: {pdfFolderPath}");
                UpdateStatus($"  Found {pdfFiles.Length} PDF files");
                
                previewButton.Enabled = true;
            }
        }

        private void PreviewRenames_Click(object sender, EventArgs e)
        {
            try
            {
                previewListView.Items.Clear();
                renameOperations.Clear();
                
                UpdateStatus("\n--- Generating rename preview ---");
                
                var pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                var studentLookup = studentData.ToDictionary(s => s.StudentNo, s => s);
                
                int matchCount = 0;
                int skipCount = 0;
                
                foreach (var pdfFile in pdfFiles)
                {
                    string currentFilename = Path.GetFileName(pdfFile);
                    string studentNo = ExtractStudentNumber(currentFilename);
                    
                    if (string.IsNullOrEmpty(studentNo))
                    {
                        var item = new ListViewItem(currentFilename);
                        item.SubItems.Add("✗");
                        item.SubItems.Add("(no change)");
                        item.SubItems.Add("Could not extract student number");
                        item.ForeColor = ColorTranslator.FromHtml("#94A3B8");
                        previewListView.Items.Add(item);
                        skipCount++;
                        continue;
                    }
                    
                    if (studentLookup.ContainsKey(studentNo))
                    {
                        var student = studentLookup[studentNo];
                        string newFilename = GenerateNewFilename(student);
                        
                        var operation = new RenameOperation
                        {
                            OriginalPath = pdfFile,
                            OriginalFilename = currentFilename,
                            NewFilename = newFilename,
                            NewPath = Path.Combine(pdfFolderPath, newFilename),
                            StudentNo = studentNo
                        };
                        
                        renameOperations.Add(operation);
                        
                        var item = new ListViewItem(currentFilename);
                        item.SubItems.Add("→");
                        item.SubItems.Add(newFilename);
                        item.SubItems.Add($"Match: {student.Forename} {student.Surname}");
                        item.ForeColor = ColorTranslator.FromHtml("#059669");
                        previewListView.Items.Add(item);
                        
                        matchCount++;
                    }
                    else
                    {
                        var item = new ListViewItem(currentFilename);
                        item.SubItems.Add("✗");
                        item.SubItems.Add("(no change)");
                        item.SubItems.Add($"Student {studentNo} not found in data");
                        item.ForeColor = ColorTranslator.FromHtml("#DC2626");
                        previewListView.Items.Add(item);
                        skipCount++;
                    }
                }
                
                UpdateStatus($"✓ Preview complete: {matchCount} files will be renamed, {skipCount} skipped");
                
                if (matchCount > 0)
                {
                    processButton.Enabled = true;
                }
                else
                {
                    MessageBox.Show("No matching files found. Please check that:\n\n" +
                        "1. PDF files are named like: XXXXXXXX-XX-XX-OVERVIEW.PDF\n" +
                        "2. Student numbers in the data file match the PDFs", 
                        "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating preview: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"✗ ERROR: {ex.Message}");
            }
        }

        private string ExtractStudentNumber(string filename)
        {
            // Extract student number from format: XXXXXXXX-XX-XX-OVERVIEW.PDF
            var parts = filename.Split('-');
            if (parts.Length >= 1)
            {
                string firstPart = parts[0].Trim();
                // Check if it's numeric and reasonable length
                if (firstPart.All(char.IsDigit) && firstPart.Length >= 7 && firstPart.Length <= 10)
                {
                    return firstPart;
                }
            }
            return null;
        }

        private string GenerateNewFilename(StudentRecord student)
        {
            // Format: b[BatchNumber] [Forename] [Surname] [StudentNo] [H/OH] [UKGrade].pdf
            
            // Extract batch number
            string batchNum = ExtractBatchNumber(student.Batch);
            
            // Convert names to proper case
            string forename = ToProperCase(student.Forename);
            string surname = ToProperCase(student.Surname);
            
            // Determine fee status code
            string feeCode = DetermineFeeStatusCode(student.FeeStatus);
            
            // Format UK grade
            string gradeCode = FormatUKGrade(student.UKGrade);
            
            // Build filename
            string filename = $"b{batchNum} {forename} {surname} {student.StudentNo} {feeCode} {gradeCode}.pdf";
            
            // Clean filename (remove invalid characters)
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                filename = filename.Replace(c, '_');
            }
            
            return filename;
        }

        private string ExtractBatchNumber(string batch)
        {
            if (string.IsNullOrWhiteSpace(batch))
                return "0";
            
            // Extract number from "Batch 1", "Batch 2", etc.
            var digits = new string(batch.Where(char.IsDigit).ToArray());
            return string.IsNullOrEmpty(digits) ? "0" : digits;
        }

        private string ToProperCase(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;
            
            // Convert to proper case (Title Case)
            var textInfo = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(text.ToLower());
        }

        private string DetermineFeeStatusCode(string feeStatus)
        {
            if (string.IsNullOrWhiteSpace(feeStatus))
                return "?";
            
            string lower = feeStatus.ToLower();
            
            if (lower.Contains("home"))
                return "H";
            else if (lower.Contains("overseas"))
                return "OH";
            else
                return "?";
        }

        private string FormatUKGrade(string ukGrade)
        {
            if (string.IsNullOrWhiteSpace(ukGrade))
                return "?";
            
            string clean = ukGrade.Trim().Replace(".", "_");
            
            // Handle common formats
            if (clean == "2_1" || clean == "2:1")
                return "2_1";
            else if (clean == "2_2" || clean == "2:2")
                return "2_2";
            else if (clean == "1" || clean == "1_0" || clean == "1st")
                return "1";
            else if (clean == "3" || clean == "3_0" || clean == "3rd")
                return "3";
            else
                return clean;
        }

        private void ProcessRenames_Click(object sender, EventArgs e)
        {
            if (renameOperations.Count == 0)
            {
                MessageBox.Show("No files to rename. Please preview first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            var result = MessageBox.Show(
                $"Are you sure you want to rename {renameOperations.Count} files?\n\n" +
                "Renamed files will be saved to a 'Renamed' subfolder.\n" +
                "This action can be undone using the 'Undo Last Batch' button.",
                "Confirm Rename",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            
            if (result != DialogResult.Yes)
                return;
            
            try
            {
                processButton.Enabled = false;
                previewButton.Enabled = false;
                
                // Create output folder
                outputFolderPath = Path.Combine(pdfFolderPath, "Renamed");
                if (!Directory.Exists(outputFolderPath))
                {
                    Directory.CreateDirectory(outputFolderPath);
                    UpdateStatus($"\n✓ Created output folder: {outputFolderPath}");
                }
                
                UpdateStatus("\n--- Starting rename process ---");
                
                int successCount = 0;
                int errorCount = 0;
                
                completedOperations.Clear();
                
                foreach (var operation in renameOperations)
                {
                    try
                    {
                        // Update the new path to be in the Renamed subfolder
                        string newPath = Path.Combine(outputFolderPath, operation.NewFilename);
                        
                        if (File.Exists(newPath))
                        {
                            UpdateStatus($"✗ Skipped: {operation.OriginalFilename} (target already exists)");
                            errorCount++;
                            continue;
                        }
                        
                        // Copy file to new location with new name
                        File.Copy(operation.OriginalPath, newPath);
                        
                        // Store the actual paths used
                        operation.NewPath = newPath;
                        completedOperations.Add(operation);
                        
                        UpdateStatus($"✓ Renamed: {operation.OriginalFilename} → {operation.NewFilename}");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"✗ Error renaming {operation.OriginalFilename}: {ex.Message}");
                        errorCount++;
                    }
                }
                
                UpdateStatus($"\n--- Rename complete: {successCount} succeeded, {errorCount} failed ---");
                UpdateStatus($"✓ Renamed files saved to: {outputFolderPath}");
                
                MessageBox.Show(
                    $"Rename process complete!\n\n" +
                    $"Successfully renamed: {successCount} files\n" +
                    $"Errors: {errorCount} files\n\n" +
                    $"Renamed files saved to:\n{outputFolderPath}",
                    "Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                if (successCount > 0)
                {
                    undoButton.Enabled = true;
                    
                    // Ask if user wants to open the output folder
                    var openFolder = MessageBox.Show(
                        "Open the Renamed folder?",
                        "Open Folder",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (openFolder == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
                    }
                }
                
                // Refresh preview
                PreviewRenames_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during rename process: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"✗ FATAL ERROR: {ex.Message}");
            }
            finally
            {
                processButton.Enabled = renameOperations.Count > 0;
                previewButton.Enabled = true;
            }
        }

        private void UndoRename_Click(object sender, EventArgs e)
        {
            if (completedOperations.Count == 0)
            {
                MessageBox.Show("No operations to undo.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            var result = MessageBox.Show(
                $"Undo the last batch of {completedOperations.Count} renames?\n\n" +
                "This will delete the renamed files from the 'Renamed' folder.\n" +
                "Original files in the source folder remain unchanged.",
                "Confirm Undo",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            
            if (result != DialogResult.Yes)
                return;
            
            try
            {
                UpdateStatus("\n--- Undoing renames ---");
                
                int successCount = 0;
                int errorCount = 0;
                
                foreach (var operation in completedOperations)
                {
                    try
                    {
                        if (File.Exists(operation.NewPath))
                        {
                            File.Delete(operation.NewPath);
                            UpdateStatus($"✓ Deleted: {operation.NewFilename}");
                            successCount++;
                        }
                        else
                        {
                            UpdateStatus($"✗ Could not find: {operation.NewFilename}");
                            errorCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"✗ Error deleting {operation.NewFilename}: {ex.Message}");
                        errorCount++;
                    }
                }
                
                UpdateStatus($"--- Undo complete: {successCount} deleted, {errorCount} failed ---");
                
                MessageBox.Show(
                    $"Undo complete!\n\n" +
                    $"Deleted: {successCount} files\n" +
                    $"Errors: {errorCount} files\n\n" +
                    $"Original PDF files remain in the source folder.",
                    "Undo Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                completedOperations.Clear();
                undoButton.Enabled = false;
                
                // Refresh preview
                PreviewRenames_Click(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during undo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus($"✗ ERROR: {ex.Message}");
            }
        }
    }
}
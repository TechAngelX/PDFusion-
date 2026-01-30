using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using OfficeOpenXml;

namespace PDFusion
{
    public partial class MainWindow : Window
    {
        private string dataFilePath = "";
        private string pdfFolderPath = "";
        private string outputFolderPath = "";
        private List<StudentRecord> studentData = new List<StudentRecord>();
        private List<RenameOperation> renameOperations = new List<RenameOperation>();
        private List<RenameOperation> completedOperations = new List<RenameOperation>();
        private ObservableCollection<PreviewItem> previewItems = new ObservableCollection<PreviewItem>();

        public MainWindow()
        {
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            InitializeComponent();
            SetupDragDrop();
            SetWindowIcon();

            PreviewGrid.ItemsSource = previewItems;
        }

        private void SetWindowIcon()
        {
            try
            {
                // Create a simple purple document icon programmatically
                var bitmap = CreateAppIcon(64, 64);
                Icon = new WindowIcon(bitmap);
            }
            catch
            {
                // Ignore icon errors
            }
        }

        private Bitmap CreateAppIcon(int width, int height)
        {
            var pixelData = new byte[width * height * 4];

            // Colors (BGRA format)
            byte[] purple = { 0xF6, 0x5C, 0x8B, 0xFF };      // #8B5CF6
            byte[] darkPurple = { 0xED, 0x3A, 0x7C, 0xFF };  // #7C3AED
            byte[] white = { 0xFF, 0xFF, 0xFF, 0xFF };
            byte[] green = { 0x81, 0xB9, 0x10, 0xFF };       // #10B981
            byte[] transparent = { 0x00, 0x00, 0x00, 0x00 };

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    int idx = (y * width + x) * 4;
                    byte[] color = transparent;

                    // Rounded rectangle background (with corner radius ~12)
                    int margin = 4;
                    int cornerRadius = 12;
                    bool inRect = x >= margin && x < width - margin && y >= margin && y < height - margin;
                    bool inCorner = false;

                    if (inRect)
                    {
                        // Check corners
                        int dx, dy;
                        if (x < margin + cornerRadius && y < margin + cornerRadius)
                        {
                            dx = x - (margin + cornerRadius);
                            dy = y - (margin + cornerRadius);
                            inCorner = dx * dx + dy * dy > cornerRadius * cornerRadius;
                        }
                        else if (x >= width - margin - cornerRadius && y < margin + cornerRadius)
                        {
                            dx = x - (width - margin - cornerRadius - 1);
                            dy = y - (margin + cornerRadius);
                            inCorner = dx * dx + dy * dy > cornerRadius * cornerRadius;
                        }
                        else if (x < margin + cornerRadius && y >= height - margin - cornerRadius)
                        {
                            dx = x - (margin + cornerRadius);
                            dy = y - (height - margin - cornerRadius - 1);
                            inCorner = dx * dx + dy * dy > cornerRadius * cornerRadius;
                        }
                        else if (x >= width - margin - cornerRadius && y >= height - margin - cornerRadius)
                        {
                            dx = x - (width - margin - cornerRadius - 1);
                            dy = y - (height - margin - cornerRadius - 1);
                            inCorner = dx * dx + dy * dy > cornerRadius * cornerRadius;
                        }

                        if (!inCorner)
                        {
                            // Gradient from purple to dark purple
                            float t = (float)(x + y) / (width + height);
                            color = new byte[] {
                                (byte)(purple[0] + t * (darkPurple[0] - purple[0])),
                                (byte)(purple[1] + t * (darkPurple[1] - purple[1])),
                                (byte)(purple[2] + t * (darkPurple[2] - purple[2])),
                                0xFF
                            };
                        }
                    }

                    // Draw white document shape (left side)
                    int docLeft = 12, docTop = 14, docWidth = 18, docHeight = 36;
                    if (x >= docLeft && x < docLeft + docWidth && y >= docTop && y < docTop + docHeight)
                    {
                        // Folded corner
                        int foldSize = 6;
                        if (x >= docLeft + docWidth - foldSize && y < docTop + foldSize)
                        {
                            if (x - (docLeft + docWidth - foldSize) + (y - docTop) < foldSize)
                                color = white;
                        }
                        else
                        {
                            color = white;
                        }
                    }

                    // Draw arrow (middle)
                    int arrowX = 33, arrowY = 28;
                    int ax = x - arrowX, ay = y - arrowY;
                    if (ax >= 0 && ax < 12 && ay >= 3 && ay < 9) color = green; // shaft
                    if (ax >= 6 && ax < 14 && ay >= 0 && ay < 12) // head
                    {
                        int tipX = 13, tipY = 6;
                        if (Math.Abs(ay - tipY) <= (tipX - ax) / 2 + 1 && ax >= 6)
                            color = green;
                    }

                    // Draw second document with checkmark (right side)
                    int doc2Left = 44, doc2Top = 14;
                    if (x >= doc2Left && x < doc2Left + docWidth && y >= doc2Top && y < doc2Top + docHeight)
                    {
                        int foldSize = 6;
                        if (x >= doc2Left + docWidth - foldSize && y < doc2Top + foldSize)
                        {
                            if (x - (doc2Left + docWidth - foldSize) + (y - doc2Top) < foldSize)
                                color = white;
                        }
                        else
                        {
                            color = white;
                        }
                    }

                    // Green checkmark circle on right doc
                    int cx = 53, cy = 28, cr = 7;
                    int cdx = x - cx, cdy = y - cy;
                    if (cdx * cdx + cdy * cdy <= cr * cr)
                    {
                        color = green;
                        // White checkmark inside
                        if ((x == 50 && y == 28) || (x == 51 && y == 29) || (x == 52 && y == 30) ||
                            (x == 53 && y == 29) || (x == 54 && y == 28) || (x == 55 && y == 27) || (x == 56 && y == 26))
                            color = white;
                    }

                    pixelData[idx] = color[0];
                    pixelData[idx + 1] = color[1];
                    pixelData[idx + 2] = color[2];
                    pixelData[idx + 3] = color[3];
                }
            }

            using var ms = new MemoryStream();
            // Create a simple BMP-like structure that Avalonia can read
            var bitmap = new WriteableBitmap(new PixelSize(width, height), new Vector(96, 96), Avalonia.Platform.PixelFormat.Bgra8888, Avalonia.Platform.AlphaFormat.Premul);
            using (var fb = bitmap.Lock())
            {
                System.Runtime.InteropServices.Marshal.Copy(pixelData, 0, fb.Address, pixelData.Length);
            }
            return bitmap;
        }

        private void SetupDragDrop()
        {
            DragDrop.SetAllowDrop(DataPanel, true);
            DataPanel.AddHandler(DragDrop.DropEvent, DataPanel_Drop);
            DataPanel.AddHandler(DragDrop.DragEnterEvent, DataPanel_DragEnter);
            DataPanel.AddHandler(DragDrop.DragLeaveEvent, DataPanel_DragLeave);
        }

        private void DataPanel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.Contains(DataFormats.Files))
            {
                var files = e.Data.GetFiles()?.ToList();
                if (files != null && files.Count == 1)
                {
                    var file = files[0] as IStorageFile;
                    if (file != null)
                    {
                        var ext = Path.GetExtension(file.Name).ToLower();
                        if (ext == ".csv" || ext == ".xlsx" || ext == ".xls")
                        {
                            e.DragEffects = DragDropEffects.Copy;
                            DataPanel.Background = new SolidColorBrush(Color.Parse("#DDD6FE"));
                            return;
                        }
                    }
                }
            }
            e.DragEffects = DragDropEffects.None;
        }

        private void DataPanel_DragLeave(object sender, DragEventArgs e)
        {
            DataPanel.Background = new SolidColorBrush(Color.Parse("#F8FAFC"));
        }

        private async void DataPanel_Drop(object sender, DragEventArgs e)
        {
            DataPanel.Background = new SolidColorBrush(Color.Parse("#F8FAFC"));

            if (e.Data.Contains(DataFormats.Files))
            {
                var files = e.Data.GetFiles()?.ToList();
                if (files != null && files.Count == 1)
                {
                    var file = files[0] as IStorageFile;
                    if (file != null)
                    {
                        var ext = Path.GetExtension(file.Name).ToLower();
                        if (ext == ".csv" || ext == ".xlsx" || ext == ".xls")
                        {
                            dataFilePath = file.Path.LocalPath;
                            await LoadStudentData();
                        }
                    }
                }
            }
        }

        private void UpdateStatus(string message)
        {
            Dispatcher.UIThread.Post(() =>
            {
                StatusBox.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
                StatusBox.CaretIndex = StatusBox.Text.Length;
            });
        }

        private async void DataPanel_Click(object sender, PointerPressedEventArgs e)
        {
            await SelectDataFile();
        }

        private async Task SelectDataFile()
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) return;

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Select Student Data File (CSV or Excel)",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("Data files")
                    {
                        Patterns = new[] { "*.csv", "*.xlsx", "*.xls" }
                    },
                    new FilePickerFileType("CSV files")
                    {
                        Patterns = new[] { "*.csv" }
                    },
                    new FilePickerFileType("Excel files")
                    {
                        Patterns = new[] { "*.xlsx", "*.xls" }
                    },
                    new FilePickerFileType("All files")
                    {
                        Patterns = new[] { "*.*" }
                    }
                }
            });

            if (files.Count > 0)
            {
                dataFilePath = files[0].Path.LocalPath;
                await LoadStudentData();
            }
        }

        private async Task LoadStudentData()
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

                SetFilePanelLoaded(DataPanel, DataPanelText, Path.GetFileName(dataFilePath), studentData.Count);
                UpdateStatus($"Loaded {studentData.Count} student records from {Path.GetFileName(dataFilePath)}");

                // Enable folder panel
                FolderPanel.IsEnabled = true;
                FolderPanel.Classes.Remove("disabled");
                FolderPanelText.Foreground = new SolidColorBrush(Color.Parse("#64748B"));
                FolderPanelText.Text = "Click to select folder with PDF files to rename";
            }
            catch (Exception ex)
            {
                await ShowMessageBox("Error", "Error loading student data: " + ex.Message);
                UpdateStatus($"ERROR: {ex.Message}");
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
            int qualityRankIdx = FindColumnIndex(headers, "ApplicationQualityRank", "Application Quality Rank");

            if (studentNoIdx == -1) throw new Exception("Could not find 'StudentNo' column");
            if (batchIdx == -1) throw new Exception("Could not find 'Batch' column");
            if (forenameIdx == -1) throw new Exception("Could not find 'Forename' column");
            if (surnameIdx == -1) throw new Exception("Could not find 'Surname' column");
            if (feeStatusIdx == -1) throw new Exception("Could not find 'FeeStatus' column");
            if (ukGradeIdx == -1) throw new Exception("Could not find 'UKGrade' column");
            // ApplicationQualityRank is optional

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
                        UKGrade = ukGradeIdx < values.Count ? values[ukGradeIdx].Trim() : "",
                        ApplicationQualityRank = qualityRankIdx >= 0 && qualityRankIdx < values.Count ? values[qualityRankIdx].Trim() : ""
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
            int qualityRankCol = FindColumnIndex(headers, "ApplicationQualityRank", "Application Quality Rank") + 1;

            if (studentNoCol == 0) throw new Exception("Could not find 'StudentNo' column");
            if (batchCol == 0) throw new Exception("Could not find 'Batch' column");
            if (forenameCol == 0) throw new Exception("Could not find 'Forename' column");
            if (surnameCol == 0) throw new Exception("Could not find 'Surname' column");
            if (feeStatusCol == 0) throw new Exception("Could not find 'FeeStatus' column");
            if (ukGradeCol == 0) throw new Exception("Could not find 'UKGrade' column");
            // ApplicationQualityRank is optional (qualityRankCol will be 0 if not found)

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
                        UKGrade = worksheet.Cells[row, ukGradeCol].Value?.ToString()?.Trim() ?? "",
                        ApplicationQualityRank = qualityRankCol > 0 ? worksheet.Cells[row, qualityRankCol].Value?.ToString()?.Trim() ?? "" : ""
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

        private async void FolderPanel_Click(object sender, PointerPressedEventArgs e)
        {
            if (!FolderPanel.IsEnabled) return;
            await SelectPDFFolder();
        }

        private async Task SelectPDFFolder()
        {
            if (studentData.Count == 0)
            {
                await ShowMessageBox("Info", "Please load student data first.");
                return;
            }

            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) return;

            var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
            {
                Title = "Select folder containing PDF files to rename",
                AllowMultiple = false
            });

            if (folders.Count > 0)
            {
                pdfFolderPath = folders[0].Path.LocalPath;

                var pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                SetFilePanelLoaded(FolderPanel, FolderPanelText, Path.GetFileName(pdfFolderPath) + " folder", pdfFiles.Length);
                UpdateStatus($"Selected folder: {pdfFolderPath}");
                UpdateStatus($"  Found {pdfFiles.Length} PDF files");

                PreviewButton.IsEnabled = true;

                // Enable Append Ranking button if any students have a quality rank
                bool hasQualityRanks = studentData.Any(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank));
                QualityRankButton.IsEnabled = hasQualityRanks;
                if (hasQualityRanks)
                {
                    int rankCount = studentData.Count(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank));
                    UpdateStatus($"  Found {rankCount} students with Application Quality Rank");
                }
            }
        }

        private void SetFilePanelLoaded(Border panel, TextBlock textBlock, string filename, int recordCount)
        {
            panel.Background = new SolidColorBrush(Color.Parse("#DCFCE7"));
            panel.BorderBrush = new SolidColorBrush(Color.Parse("#86EFAC"));
            panel.Classes.Add("loaded");

            string displayName = filename.Length > 60 ? filename.Substring(0, 57) + "..." : filename;
            textBlock.Text = $"{displayName}\n{recordCount} records loaded";
            textBlock.Foreground = new SolidColorBrush(Color.Parse("#16A34A"));
            textBlock.FontWeight = FontWeight.Bold;
        }

        private void PreviewRenames_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                previewItems.Clear();
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
                        previewItems.Add(new PreviewItem
                        {
                            CurrentFilename = currentFilename,
                            Arrow = "X",
                            NewFilename = "(no change)",
                            Status = "Could not extract student number",
                            StatusColor = "#94A3B8"
                        });
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

                        previewItems.Add(new PreviewItem
                        {
                            CurrentFilename = currentFilename,
                            Arrow = "->",
                            NewFilename = newFilename,
                            Status = $"Match: {student.Forename} {student.Surname}",
                            StatusColor = "#059669"
                        });

                        matchCount++;
                    }
                    else
                    {
                        previewItems.Add(new PreviewItem
                        {
                            CurrentFilename = currentFilename,
                            Arrow = "X",
                            NewFilename = "(no change)",
                            Status = $"Student {studentNo} not found in data",
                            StatusColor = "#DC2626"
                        });
                        skipCount++;
                    }
                }

                UpdateStatus($"Preview complete: {matchCount} files will be renamed, {skipCount} skipped");

                if (matchCount > 0)
                {
                    ProcessButton.IsEnabled = true;
                }
                else
                {
                    _ = ShowMessageBox("No Matches",
                        "No matching files found. Please check that:\n\n" +
                        "1. PDF files are named like: XXXXXXXX-XX-XX-OVERVIEW.PDF\n" +
                        "2. Student numbers in the data file match the PDFs");
                }
            }
            catch (Exception ex)
            {
                _ = ShowMessageBox("Error", "Error generating preview: " + ex.Message);
                UpdateStatus($"ERROR: {ex.Message}");
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
            var textInfo = CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(text.ToLower());
        }

        private string DetermineFeeStatusCode(string feeStatus)
        {
            if (string.IsNullOrWhiteSpace(feeStatus))
                return "?";

            string lower = feeStatus.ToLower();

            // European students (including provisional) are classified as Overseas
            if (lower.Contains("european") || lower.Contains("uropean"))
                return "OS";
            else if (lower.Contains("overseas"))
                return "OS";
            else if (lower.Contains("home"))
                return "H";
            else
                return "?";
        }

        private string FormatUKGrade(string ukGrade)
        {
            if (string.IsNullOrWhiteSpace(ukGrade))
                return "XX";

            string lower = ukGrade.Trim().ToLower();
            string clean = ukGrade.Trim().Replace(".", "_");

            // Handle common formats
            if (lower.Contains("2.1") || lower.Contains("2:1") || lower == "2_1")
                return "2_1";
            else if (lower.Contains("2.2") || lower.Contains("2:2") || lower == "2_2")
                return "2_2";
            else if (lower.Contains("1st") || lower == "1" || lower == "1.0" || lower == "first")
                return "1";
            else if (lower.Contains("3rd") || lower == "3" || lower == "3.0" || lower == "third")
                return "3";
            else if (lower.Contains("master"))
                return "Masters";
            else if (lower.Contains("?") || lower == "n/a" || lower == "na" || lower == "unknown" || lower == "-")
                return "XX";
            else
                return "XX";
        }

        private async void ProcessRenames_Click(object sender, RoutedEventArgs e)
        {
            if (renameOperations.Count == 0)
            {
                await ShowMessageBox("Info", "No files to rename. Please preview first.");
                return;
            }

            var result = await ShowConfirmDialog("Confirm Rename",
                $"Are you sure you want to rename {renameOperations.Count} files?\n\n" +
                "Renamed files will be saved to a 'Renamed' subfolder.\n" +
                "This action can be undone using the 'Undo Last Batch' button.");

            if (!result) return;

            try
            {
                ProcessButton.IsEnabled = false;
                PreviewButton.IsEnabled = false;

                // Create output folder
                outputFolderPath = Path.Combine(pdfFolderPath, "Renamed");
                if (!Directory.Exists(outputFolderPath))
                {
                    Directory.CreateDirectory(outputFolderPath);
                    UpdateStatus($"\nCreated output folder: {outputFolderPath}");
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
                            UpdateStatus($"Skipped: {operation.OriginalFilename} (target already exists)");
                            errorCount++;
                            continue;
                        }

                        // Copy file to new location with new name
                        File.Copy(operation.OriginalPath, newPath);

                        // Store the actual paths used
                        operation.NewPath = newPath;
                        completedOperations.Add(operation);

                        UpdateStatus($"Renamed: {operation.OriginalFilename} -> {operation.NewFilename}");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"Error renaming {operation.OriginalFilename}: {ex.Message}");
                        errorCount++;
                    }
                }

                UpdateStatus($"\n--- Rename complete: {successCount} succeeded, {errorCount} failed ---");
                UpdateStatus($"Renamed files saved to: {outputFolderPath}");

                await ShowMessageBox("Complete",
                    $"Rename process complete!\n\n" +
                    $"Successfully renamed: {successCount} files\n" +
                    $"Errors: {errorCount} files\n\n" +
                    $"Renamed files saved to:\n{outputFolderPath}");

                if (successCount > 0)
                {
                    UndoButton.IsEnabled = true;

                    // Ask if user wants to open the output folder
                    var openFolder = await ShowConfirmDialog("Open Folder", "Open the Renamed folder?");
                    if (openFolder)
                    {
                        OpenFolder(outputFolderPath);
                    }
                }

                // Refresh preview
                PreviewRenames_Click(null, null);
            }
            catch (Exception ex)
            {
                await ShowMessageBox("Error", "Error during rename process: " + ex.Message);
                UpdateStatus($"FATAL ERROR: {ex.Message}");
            }
            finally
            {
                ProcessButton.IsEnabled = renameOperations.Count > 0;
                PreviewButton.IsEnabled = true;
            }
        }

        private async void UndoRename_Click(object sender, RoutedEventArgs e)
        {
            if (completedOperations.Count == 0)
            {
                await ShowMessageBox("Info", "No operations to undo.");
                return;
            }

            var result = await ShowConfirmDialog("Confirm Undo",
                $"Undo the last batch of {completedOperations.Count} renames?\n\n" +
                "This will delete the renamed files from the 'Renamed' folder.\n" +
                "Original files in the source folder remain unchanged.");

            if (!result) return;

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
                            UpdateStatus($"Deleted: {operation.NewFilename}");
                            successCount++;
                        }
                        else
                        {
                            UpdateStatus($"Could not find: {operation.NewFilename}");
                            errorCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"Error deleting {operation.NewFilename}: {ex.Message}");
                        errorCount++;
                    }
                }

                UpdateStatus($"--- Undo complete: {successCount} deleted, {errorCount} failed ---");

                await ShowMessageBox("Undo Complete",
                    $"Undo complete!\n\n" +
                    $"Deleted: {successCount} files\n" +
                    $"Errors: {errorCount} files\n\n" +
                    $"Original PDF files remain in the source folder.");

                completedOperations.Clear();
                UndoButton.IsEnabled = false;

                // Refresh preview
                PreviewRenames_Click(null, null);
            }
            catch (Exception ex)
            {
                await ShowMessageBox("Error", "Error during undo: " + ex.Message);
                UpdateStatus($"ERROR: {ex.Message}");
            }
        }

        private async void AddQualityRankPrefix_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(pdfFolderPath) || !Directory.Exists(pdfFolderPath))
            {
                await ShowMessageBox("Info", "Please select a PDF folder first.");
                return;
            }

            // Build lookup of student number to quality rank
            var qualityRankLookup = studentData
                .Where(s => !string.IsNullOrWhiteSpace(s.ApplicationQualityRank))
                .ToDictionary(s => s.StudentNo, s => s.ApplicationQualityRank.Trim());

            if (qualityRankLookup.Count == 0)
            {
                await ShowMessageBox("Info", "No Application Quality Rank data found in the loaded data file.");
                return;
            }

            var result = await ShowConfirmDialog("Append Ranking",
                $"This will prefix PDF files with their Application Quality Rank.\n\n" +
                $"Example: 'A - filename.pdf'\n\n" +
                $"Found {qualityRankLookup.Count} students with quality ranks.\n\n" +
                "Files will be renamed in place. Continue?");

            if (!result) return;

            try
            {
                UpdateStatus("\n--- Appending Quality Rank Prefixes ---");

                var pdfFiles = Directory.GetFiles(pdfFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);
                int successCount = 0;
                int skipCount = 0;
                int errorCount = 0;

                foreach (var filePath in pdfFiles)
                {
                    string filename = Path.GetFileName(filePath);

                    // Try to extract student number from various filename formats
                    string studentNo = ExtractStudentNumberForRanking(filename);

                    if (string.IsNullOrEmpty(studentNo))
                    {
                        UpdateStatus($"Skipped (no student number): {filename}");
                        skipCount++;
                        continue;
                    }

                    if (!qualityRankLookup.ContainsKey(studentNo))
                    {
                        UpdateStatus($"Skipped (no quality rank for {studentNo}): {filename}");
                        skipCount++;
                        continue;
                    }

                    string qualityRank = qualityRankLookup[studentNo];

                    // Skip if already has a quality rank prefix (character followed by " - ")
                    if (filename.Length > 4 && filename.Substring(1, 3) == " - ")
                    {
                        UpdateStatus($"Skipped (already has prefix): {filename}");
                        skipCount++;
                        continue;
                    }

                    string newFilename = $"{qualityRank} - {filename}";
                    string newPath = Path.Combine(pdfFolderPath, newFilename);

                    try
                    {
                        if (File.Exists(newPath))
                        {
                            UpdateStatus($"Skipped (target exists): {newFilename}");
                            skipCount++;
                            continue;
                        }

                        File.Move(filePath, newPath);
                        UpdateStatus($"Prefixed: {filename} -> {newFilename}");
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        UpdateStatus($"Error renaming {filename}: {ex.Message}");
                        errorCount++;
                    }
                }

                UpdateStatus($"\n--- Append Ranking complete: {successCount} renamed, {skipCount} skipped, {errorCount} errors ---");

                await ShowMessageBox("Complete",
                    $"Append Ranking complete!\n\n" +
                    $"Successfully prefixed: {successCount} files\n" +
                    $"Skipped: {skipCount} files\n" +
                    $"Errors: {errorCount} files");

                // Disable the button after use to prevent double-prefixing
                QualityRankButton.IsEnabled = false;
            }
            catch (Exception ex)
            {
                await ShowMessageBox("Error", "Error appending ranking prefixes: " + ex.Message);
                UpdateStatus($"ERROR: {ex.Message}");
            }
        }

        private string ExtractStudentNumberForRanking(string filename)
        {
            // Try multiple formats:
            // 1. Original format: XXXXXXXX-XX-XX-OVERVIEW.PDF (student number before first dash)
            // 2. Renamed format: b1 John Smith 26049530 H 2_1.pdf (7-10 digit number in filename)

            // First try the original format (number before first dash)
            var dashParts = filename.Split('-');
            if (dashParts.Length >= 1)
            {
                string firstPart = dashParts[0].Trim();
                if (firstPart.All(char.IsDigit) && firstPart.Length >= 7 && firstPart.Length <= 10)
                {
                    return firstPart;
                }
            }

            // Then try finding any 7-10 digit number in the filename
            var spaceParts = filename.Split(' ');
            foreach (var part in spaceParts)
            {
                string cleanPart = part.Replace(".pdf", "").Replace(".PDF", "");
                if (cleanPart.All(char.IsDigit) && cleanPart.Length >= 7 && cleanPart.Length <= 10)
                {
                    return cleanPart;
                }
            }

            return null;
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void OpenFolder(string path)
        {
            try
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Process.Start("explorer.exe", path);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", path);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", path);
                }
            }
            catch
            {
                // Ignore errors opening folder
            }
        }

        private async Task ShowMessageBox(string title, string message)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 400,
                Height = 200,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                CanResize = false
            };

            var panel = new StackPanel
            {
                Margin = new Thickness(20),
                Spacing = 20
            };

            panel.Children.Add(new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap
            });

            var button = new Button
            {
                Content = "OK",
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                Width = 100
            };
            button.Click += (s, e) => dialog.Close();
            panel.Children.Add(button);

            dialog.Content = panel;
            await dialog.ShowDialog(this);
        }

        private async Task<bool> ShowConfirmDialog(string title, string message)
        {
            var result = false;
            var dialog = new Window
            {
                Title = title,
                Width = 450,
                Height = 220,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                CanResize = false
            };

            var panel = new StackPanel
            {
                Margin = new Thickness(20),
                Spacing = 20
            };

            panel.Children.Add(new TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap
            });

            var buttonPanel = new StackPanel
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                Spacing = 10
            };

            var yesButton = new Button
            {
                Content = "Yes",
                Width = 80
            };
            yesButton.Click += (s, e) =>
            {
                result = true;
                dialog.Close();
            };

            var noButton = new Button
            {
                Content = "No",
                Width = 80
            };
            noButton.Click += (s, e) =>
            {
                result = false;
                dialog.Close();
            };

            buttonPanel.Children.Add(yesButton);
            buttonPanel.Children.Add(noButton);
            panel.Children.Add(buttonPanel);

            dialog.Content = panel;
            await dialog.ShowDialog(this);

            return result;
        }
    }
}

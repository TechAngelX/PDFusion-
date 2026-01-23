namespace PDFusion
{
    // Student data record
    public class StudentRecord
    {
        public string StudentNo { get; set; }
        public string Batch { get; set; }
        public string Forename { get; set; }
        public string Surname { get; set; }
        public string FeeStatus { get; set; }
        public string UKGrade { get; set; }
    }

    // Rename operation record for undo functionality
    public class RenameOperation
    {
        public string OriginalPath { get; set; }
        public string OriginalFilename { get; set; }
        public string NewPath { get; set; }
        public string NewFilename { get; set; }
        public string StudentNo { get; set; }
    }

    // Preview item for DataGrid display
    public class PreviewItem
    {
        public string CurrentFilename { get; set; }
        public string Arrow { get; set; }
        public string NewFilename { get; set; }
        public string Status { get; set; }
        public string StatusColor { get; set; }
    }
}

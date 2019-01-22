namespace ReadExcel
{
    public class UploadFileImportModel
    {
        public string FileName { get; set; }
        public string FileNameSystem { get; set; }
        public string UploadPath { get; set; }
        public string SavePathSuccess { get; set; }
        public string SavePathFailed { get; set; }
        public string UploadBy { get; set; }
    }
}
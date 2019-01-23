

namespace ReadExcel.IServices
{
    /// <summary>
    /// The 'AbstractProduct' interface.
    /// </summary>
    public interface IImportExcelService
    {
        
        int Import(string filePath);

        int Import(string filePath, string uploadBy);
    }
}

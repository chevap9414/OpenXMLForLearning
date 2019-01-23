

namespace ReadExcel.IServices
{
    /// <summary>
    /// The 'AbstractFactory' interface.
    /// </summary>
    public interface IExcelService
    {

        IImportExcelService Import();
        IImportExcelService Import(string filePath);
        IExportExcelService Export();
    }
}

using ReadExcel.IServices;

namespace ReadExcel.Services
{
    /// <summary>
    /// The 'ConcreateFactory' class.
    /// </summary>
    public class ModelTypeList : IExcelService
    {
        public IExportExcelService Export()
        {
            return new MTList();
        }

        public IImportExcelService Import()
        {
            return new MTList();
        }

        public IImportExcelService Import(string filePath)
        {
            return new MTList();
        }

        public IReadFile ReadFile(string filePath)
        {
            return new MTList();
        }
    }
}

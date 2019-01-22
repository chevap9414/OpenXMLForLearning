
using ReadExcel.IServices;

namespace ReadExcel.Services
{
    /// <summary>
    /// The 'ConcreateFactory' class.
    /// </summary>
    public class LOTypeList : IExcelService
    {
        public IExportExcelService Export()
        {
            throw new System.NotImplementedException();
        }

        public IImportExcelService Import()
        {
            return new LOList();
        }
    }
}

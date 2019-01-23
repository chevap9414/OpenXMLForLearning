

using ReadExcel.IServices;

namespace ReadExcel
{
    /// <summary>
    /// The 'Product' class.
    /// </summary>
    public class MTList : IImportExcelService, IExportExcelService, IReadFile
    {
        public int Export()
        {
            throw new System.NotImplementedException();
        }

        public int Import(string filePath)
        {
            return 1;
        }

        public int Import(string filePath, string uploadBy)
        {
            return 1;
        }

        public ModelTypeTempSheetModel ReadFile(string fileName)
        {
            return new ModelTypeTempSheetModel();
        }

        void IReadFile.ReadFile(string fileName)
        {
            throw new System.NotImplementedException();
        }
    }
}

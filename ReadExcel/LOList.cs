

using ReadExcel.IServices;

namespace ReadExcel
{
    /// <summary>
    /// The 'Product' class.
    /// </summary>
    public class LOList : IImportExcelService
    {
        public int Import(string filePath)
        {
            return 1;
        }

        public int Import(string filePath, string uploadBy)
        {
            return 1;
        }
    }
}

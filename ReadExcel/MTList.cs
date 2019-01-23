

using ReadExcel.IServices;

namespace ReadExcel
{
    /// <summary>
    /// The 'Product' class.
    /// </summary>
    public class MTList : IImportExcelService, IExportExcelService
    {
        public int Export()
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// asdasdasgfadfa
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
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

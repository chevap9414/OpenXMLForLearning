using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ModelTypeImportExcel : IModelTypeImportExcel
    {
        public ModelTypeImportExcel()
        {
        }

        /// <summary>
        /// Represent's Import excel file to Database.
        /// </summary>
        /// <param name="fileName">
        /// Path file name.
        /// </param>
        /// <param name="uploadBy">
        /// Name of uploader.
        /// </param>
        /// <returns></returns>
        public bool ImportExcel(string fileName, string uploadBy)
        {
            bool IsSucceed = false;

            return IsSucceed;
        }

        private DataTable ReadExcel(string fileName)
        {
            DataTable data = new DataTable();

            return data;
        }
    }
}

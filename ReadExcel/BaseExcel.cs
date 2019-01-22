using ReadExcel.IServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    /// <summary>
    /// The 'Client' class.
    /// </summary>
    class BaseExcel : IBaseExcel
    {
        readonly IImportExcelService importExcelService;
        readonly IExportExcelService exportExcelService;

        public BaseExcel(IExcelService excelService)
        {
            importExcelService = excelService.Import();
            exportExcelService = excelService.Export();
        }

        public int ImportMTList(string fileName)
        {
            return importExcelService.Import(fileName);
        }

        public int ImportMTList(string fileName, string uploadBy)
        {
            return importExcelService.Import(fileName, uploadBy);
        }
    }
}

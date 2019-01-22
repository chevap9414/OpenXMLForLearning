using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ModelTypeTempSheetModel
    {
        public int ModelTypeTempSheetID { get; set; }
        public int ModelTypeUploadID { get; set; }
        public int SheetNo { get; set; }
        public List<ModelTypeTempRowModel> modelTypeTempRowModels { get; set; } = new List<ModelTypeTempRowModel>();
    }
}

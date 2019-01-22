using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ModelTypeTempRowModel
    {
        public int ModelTypeTempRowID { get; set; }
        public int ModelTypeTempSheetID { get; set; }
        public int RowNo { get; set; }
        public string PNo { get; set; }
        public string VIN { get; set; }
        public List<ModelTypeTempEngineModel> modelTypeTempEngines { get; set; } = new List<ModelTypeTempEngineModel>();
        public List<ModelTypeTempEquipmentModel> modelTypeTempEquipmentModels { get; set; } = new List<ModelTypeTempEquipmentModel>();
        public List<ModelTypeTempTypeModel> modelTypeTempTypeModels { get; set; } = new List<ModelTypeTempTypeModel>();


    }
}

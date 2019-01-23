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
        public List<ModelTypeTempEngineModel> ModelTypeTempEngines { get; set; } = new List<ModelTypeTempEngineModel>();
        public List<ModelTypeTempEquipmentModel> ModelTypeTempEquipmentModels { get; set; } = new List<ModelTypeTempEquipmentModel>();
        public List<ModelTypeTempTypeModel> ModelTypeTempTypeModels { get; set; } = new List<ModelTypeTempTypeModel>();


    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ModelTypeUploadModel
    {
        public int ModelTypeUploadID { get; set; }
        public int? CompanyID { get; set; }
        public int? ModelID { get; set; }
        public int? YMID { get; set; }
        public int? ModelTypeTempID { get; set; }
        public int? ModelTypeID { get; set; }
        public string Version { get; set; }
        public string Door { get; set; }
        public string Engine { get; set; }
        public int? I_ModelTypePath { get; set; }
        public string I_ModelTypeFileNameOriginal { get; set; }
        public int? O_ModelTypePath { get; set; }
        public string O_ModelTypeFile { get; set; }
        public int? UploadStatusID { get; set; }
        public string Result { get; set; }
        public string CreatedBy { get; set; }
        public System.DateTime CreatedDate { get; set; }
        public string UpdatedBy { get; set; }
        public DateTime UpdatedDate { get; set; }
        public string I_ModelTypeFileNameSystem { get; set; }

        //public virtual M_FileManager M_FileManager { get; set; }
        public List<ModelTypeTempSheetModel> ModelTypeTempSheetModels { get; set; } = new List<ModelTypeTempSheetModel>();

    }
}

//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace DataLayer
{
    using System;
    using System.Collections.Generic;
    
    public partial class T_PartDemand
    {
        public int PartDemandID { get; set; }
        public int UploadDetailID { get; set; }
        public int FileLineNo { get; set; }
        public string GroupType { get; set; }
        public string SupplyRegion { get; set; }
        public string SupplyPlant { get; set; }
        public string BasicPartNumber { get; set; }
        public string PartName { get; set; }
        public string MLCode { get; set; }
        public string MLName { get; set; }
        public string ReceivePlant { get; set; }
        public string AFRegion { get; set; }
        public string AFPlant { get; set; }
        public string Model { get; set; }
        public string SalesYM { get; set; }
        public string EngType { get; set; }
        public string Disp { get; set; }
        public string Head { get; set; }
        public string TMType { get; set; }
        public string TMClass { get; set; }
        public string Drive { get; set; }
        public string MOTCap { get; set; }
        public string KeyCode { get; set; }
        public string BOMCodeM { get; set; }
        public string BOMCodeT { get; set; }
        public string ProductionDate { get; set; }
        public int ProductionQty { get; set; }
        public string OperationMonth { get; set; }
        public string Space1 { get; set; }
        public string Space2 { get; set; }
        public string Space3 { get; set; }
    
        public virtual T_PartDemandFileUploadDetail T_PartDemandFileUploadDetail { get; set; }
    }
}

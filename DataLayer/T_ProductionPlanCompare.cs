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
    
    public partial class T_ProductionPlanCompare
    {
        public int ProductionPlanCompareID { get; set; }
        public Nullable<int> ProductionPlanAOID01 { get; set; }
        public Nullable<int> ProductionPlanAOID02 { get; set; }
        public string AdditionalCondition { get; set; }
        public string DownloadedBy { get; set; }
        public Nullable<System.DateTime> DownloadedDate { get; set; }
    
        public virtual T_ProductionPlanAO T_ProductionPlanAO { get; set; }
        public virtual T_ProductionPlanAO T_ProductionPlanAO1 { get; set; }
    }
}

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
    
    public partial class M_CountryListTempModelType
    {
        public int CountryListTempModelTypeID { get; set; }
        public int CountryListTempRowID { get; set; }
        public string CountryListTempModeltypeName { get; set; }
        public int ColumnIndex { get; set; }
    
        public virtual M_CountryListTempRow M_CountryListTempRow { get; set; }
    }
}

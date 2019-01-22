using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel.IServices
{
    public interface IBaseExcel
    {
        /// <summary>
        /// Respresent's MT List import excel file.
        /// </summary>
        /// <param name="fileName">
        /// Full path filename.
        /// </param>
        /// <returns>Upoad status code.</returns>
        int ImportMTList(string fileName);


        int ImportMTList(string fileName, string uploadBy);
    }
}

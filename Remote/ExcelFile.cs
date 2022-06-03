using IronXL;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Remote
{
    class ExcelFile
    {
        /// <summary>
        /// метод читает файл и копирует его в дататейбл
        /// </summary>
        /// <param name="fileName">name of the file</param>
        /// <returns>DataTable</returns>
        public DataTable ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);            
            WorkSheet sheet = workbook.DefaultWorkSheet;            
            return sheet.ToDataTable(true);
        }
    }
}

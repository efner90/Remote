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
        public void SaveExcel(string filename, DataTable dataTable)
        {
            IronXL.Options.CreatingOptions creatingOptions = new IronXL.Options.CreatingOptions();
            creatingOptions.DefaultFileFormat = ExcelFileFormat.XLS;
            WorkBook workbook = WorkBook.Create(creatingOptions);  
            workbook.LoadWorkSheet(dataTable);
            WorkBook sheet = workbook.SaveAs(filename);
            //var newSheet = workbook.DefaultWorkSheet;
            //return newSheet.ToDataTable(true);
        }



    }
}

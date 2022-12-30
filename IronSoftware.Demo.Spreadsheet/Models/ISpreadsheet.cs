using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronSoftware.Demo.Spreadsheet.Models
{
    internal interface ISpreadsheet
    {
        //Loads a workbook
        WorkBook LoadWorkBook(string path);

        //gets worksheet
        WorkSheet GetWorkSheet(WorkBook workBook, int sheetNum = 0);

        (string err, object cellValue) GetCellValue(WorkSheet worksheet,string cell);
        void DisplayCellRangeValue(WorkSheet worksheet, string cellRange);
    }
}

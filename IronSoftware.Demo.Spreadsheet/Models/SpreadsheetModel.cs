using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronSoftware.Demo.Spreadsheet.Models
{
    internal class SpreadsheetModel : ISpreadsheet
    {
        /// <summary>
        /// Gets all worksheets in a specified workbook
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="sheetNum">Default shee is 0</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public WorkSheet GetWorkSheet(WorkBook workBook, int sheetNum = 0)
        {
            if (workBook == null)
            {
                throw new ArgumentNullException("Workbook Cannot be empty");

            }
            else
            {
                return workBook.WorkSheets.ElementAt(sheetNum);
            }
        }

        /// <summary>
        /// Loads a workbook
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public WorkBook LoadWorkBook(string path)
        {
            var workbook = WorkBook.Load(path);
            return workbook;
        }

        /// <summary>
        /// Display cell values in a range
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cellRange"></param>
        public void DisplayCellRangeValue(WorkSheet worksheet, string cellRange)
        {
            foreach (var cell in worksheet[cellRange.ToUpper()])
            {
                //Console.WriteLine("Cell {0} alue '{1}'", cell.AddressString, cell.Text);
                Console.WriteLine($"Cell {cell.Address} Value: {cell.Text}");
            }
        }

        /// <summary>
        /// Get cell value in a worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public (string err, object cellValue) GetCellValue(WorkSheet worksheet, string cell)
        {
            try
            {
                return ("", worksheet[cell.ToUpper()].Value);
            }
            catch (Exception ex)
            {
                return (ex.Message.ToString(), worksheet[cell.ToUpper()].Value);
            }
        }

        
    }
}

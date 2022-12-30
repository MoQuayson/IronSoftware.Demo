using IronSoftware.Demo.Spreadsheet.Models;

namespace IronSoftware.Demo.Spreadsheet
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // This will get the current WORKING directory (i.e. \bin\Debug)
            string workingDirectory = Environment.CurrentDirectory;

            // This will get the current PROJECT directory
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

            //get spreaddsheets folder location (i.e /Data/Spreadsheets)
            var spreadSheetDirectory = $@"{projectDirectory}/Data/Spreadsheets";
            

            var spreadSheetModel = new SpreadsheetModel();
            var wkBook =   spreadSheetModel.LoadWorkBook($@"{spreadSheetDirectory}/sampledatafoodsales.xlsx");
            Console.WriteLine($"Worksheets found: {wkBook.WorkSheets.Count}"); ;
            var wkSheet =  spreadSheetModel.GetWorkSheet(wkBook, 1);

            
            //get cell value
            var cell = "D5";
            var cellRange = "E2:E10";

            (string err, object cellValue) = spreadSheetModel.GetCellValue(wkSheet,cell);

            if(err != "" || err != null)
            {
                ConsoleMessage.DisplayErrorMessage(err);
            }
            Console.WriteLine($"Cell {cell} Value: {cellValue}\n");

            Console.WriteLine($"Cell Range {cellRange} Values:\n");
            //display cell ranges
            spreadSheetModel.DisplayCellRangeValue(wkSheet, cellRange);
        }
    }
}
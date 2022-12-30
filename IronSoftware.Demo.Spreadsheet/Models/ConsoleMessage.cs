using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronSoftware.Demo.Spreadsheet.Models
{
    internal class ConsoleMessage
    {
        public static void DisplayErrorMessage(string message)
        {
            Console.ForegroundColor= ConsoleColor.Red;
            Console.Write(message);
            Console.ResetColor();
        }

        public void DisplaySuccessMessage(string message)
        {
            throw new NotImplementedException();
        }
    }
}

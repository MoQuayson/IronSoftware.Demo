using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronSoftware.Demo.Spreadsheet.Models
{
    internal interface IConsoleMessage
    {
        abstract void DisplayErrorMessage(string message);
        void DisplaySuccessMessage(string message);
    }
}

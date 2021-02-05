using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ConsoleApp1
{

    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(@"C:\proy\Excel\final2.xlsx");
            workbook.RefreshAll();
            System.Threading.Thread.Sleep(5000);
            application.DisplayAlerts = false;
            workbook.Save();
            workbook.Close(false, @"C:\proy\Excel\final2.xlsx", null);
            application.Quit();
            workbook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
        }
    }
}

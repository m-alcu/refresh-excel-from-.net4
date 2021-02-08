using System;
using System.IO;
using Util;

namespace RefreshDatabase
{

    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("**********************************************");
            Console.WriteLine("***************  INICI   *********************");
            Console.WriteLine("******* ACTUALITZACIO DE DOCUMENTS ***********");
            Console.WriteLine("**********************************************");
            Console.WriteLine("**********************************************");

            string configPath = AppDomain.CurrentDomain.BaseDirectory + "\\config.ini";
            PropertiesUtility properties = new PropertiesUtility();
            properties.LoadProperties(configPath);

            string timeoutStr = properties.GetProperty("timeout");
            Console.Write("Espera per document: " + timeoutStr + " millisegons");
            Console.WriteLine();
            int timeout = Int32.Parse(timeoutStr);

            string pathList = properties.GetProperty("pathlist");
            Console.WriteLine("Llista ubicada a: " + pathList + "Lista.xlsx");

            string pathExcels = properties.GetProperty("pathexcels");
            Console.WriteLine("Excels ubicats a: " + pathExcels + "*.*");

            Excel.Application application = new Excel.Application();

            string path = pathList + "Lista.xlsx";
            Excel.Workbook lista = application.Workbooks.Open(path);
            Excel.Worksheet mySheet = (Excel.Worksheet)lista.Sheets["Hojas"];

            string[] terms = new string[400];
            int numTerms = 0;
            Excel.Range dataRange = null;
            for (int row = 1; row < mySheet.Rows.Count; row++) 
            {
                dataRange = (Excel.Range)mySheet.Cells[row, 1];
                if (dataRange.Value2 == null) break;
                if (row > 400) break;
                terms[row-1] = String.Format(dataRange.Value2.ToString());
                numTerms++;
            }
            Console.WriteLine();
            lista.Close(false, path, null);


            for (int i = 0; i < numTerms; i++)
            {
                Excel.Workbook workbook = application.Workbooks.Open(pathExcels+terms[i]);
                workbook.RefreshAll();
                Console.Write(i+1+". Actualitzant... "+ pathExcels+terms[i]);
                System.Threading.Thread.Sleep(timeout);
                Console.Write(" ok");
                Console.WriteLine();
                application.DisplayAlerts = false;
                workbook.Save();
                workbook.Close(false, pathExcels+terms[i], null);
                workbook = null;
            }

            Console.WriteLine("**********************************************");
            Console.WriteLine("***************    FI    *********************");
            Console.WriteLine("******* ACTUALITZACIO DE DOCUMENTS ***********");
            Console.WriteLine("**********************************************");
            Console.WriteLine("**********************************************");

            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
        }
    }
}

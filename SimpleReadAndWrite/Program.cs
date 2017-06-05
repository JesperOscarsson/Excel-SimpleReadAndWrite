using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace SimpleReadAndWrite
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Menu/Flow

            bool loop = true;

            while (loop == true)
            {
                Console.WriteLine("What do you want to do?");
                Console.WriteLine("1. Run");
                string answer = Console.ReadLine();

                switch (answer)
                {
                    case "1":
                        Run();
                        loop = false;
                        break;
                }

            }

            ws.Name = "testSheet";
            //wb.SaveAs(Filename: "testExcelFile");
        }

        public static void Run()
        {
            Console.WriteLine("Run has been executed");
        }
    }
}

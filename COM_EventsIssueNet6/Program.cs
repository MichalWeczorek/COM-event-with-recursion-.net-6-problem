using System;
using Microsoft.Office.Interop.Excel;
namespace COM_EventIssueNet6
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Test");
            OpenExcel();
        }
        private static void OpenExcel()
        {
            var excelApp = new Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

             workSheet = excelApp.ActiveSheet as Worksheet;
            excelApp.AfterCalculate += ExcelApp_AfterCalculate;
            workSheet.Cells[1, "A"] = 2;
        }
        private static int counter = 2;
        private static Worksheet workSheet;


        private static void ExcelApp_AfterCalculate()
        {
            counter--;
            if (counter > 0)
            {
                workSheet.Cells[1, "A"] = 2;
            }
            Console.Write("AfterCalculate");
        }
    }
}

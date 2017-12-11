namespace ExcelFreeze
{
    using System;
    using System.Runtime.InteropServices;
    using Excel = Microsoft.Office.Interop.Excel;

    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Open(@"C:\test.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[1];

            worksheet.Activate();
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);

            workbook.Save();
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            application.Quit();
            Marshal.FinalReleaseComObject(application);
        }
    }
}

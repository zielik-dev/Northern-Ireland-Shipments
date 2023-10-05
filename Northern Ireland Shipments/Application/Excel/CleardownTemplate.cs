using Microsoft.Office.Interop.Excel;
using Northern_Ireland_Shipments.Logs;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Northern_Ireland_Shipments.Application.Excel
{
    public class CleardownTemplate
    {
        public static void Clear(string environment, Worksheet reportWs)
        {
            try
            {
                int nRows = reportWs.UsedRange.Rows.Count;

                if (nRows <= 2)
                    nRows = 3;

                Range rng = reportWs.get_Range("A3", $"AD{nRows}");
                rng.Clear();
                Console.WriteLine("Template cleared");
            }
            catch (Exception e)
            {
                string exception = e.ToString();
                string dbExceptionPrName = "Cleardown Template - Completed";
                Console.WriteLine($"Error: {exception}");
                InsertLogToDb.Exception(dbExceptionPrName, environment);
                ExceptionLogToFile.Instance.WriteExceptionLog(exception);
            }
        }
    }
}
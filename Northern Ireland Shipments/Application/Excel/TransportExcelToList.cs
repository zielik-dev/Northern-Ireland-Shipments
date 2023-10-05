using Microsoft.Office.Interop.Excel;
using Northern_Ireland_Shipments.Infrastructure;
using Northern_Ireland_Shipments.Infrastructure.Smtp;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.Models.Queries;
using System.Reflection;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Northern_Ireland_Shipments.Application.Excel
{
    public class TransportExcelToList : ConnectionStrings
    {

        public static List<TransportSrcFileModel> ExcelToList(string environment, Microsoft.Office.Interop.Excel.Application xlApp, string inboundFile, DateTime dt)
        {
            List<TransportSrcFileModel> list = new();

            DateTime dateTime = dt.AddDays(-31);

            Workbook srcWb = xlApp.Workbooks.Open(sourceWb, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet srcWs = srcWb.Worksheets.get_Item(sheetSrc);

            int nRows = srcWs.UsedRange.Rows.Count;

            int exceptionIndex = 0; // Initialize exceptionIndex to a default value

            for (int i = nRows; i >= 3; i--)
            {
                try
                {
                    DateTime srcDate = DateTime.FromOADate(Convert.ToDouble((srcWs.Cells[i, 1] as Range).Value2));

                    if (srcDate >= dateTime && srcDate <= dt)
                    {
                        list.Add(new TransportSrcFileModel()
                        {
                            Date = srcDate,
                            Client_Job_Number = srcWs.Cells[i, 2].Value.ToString(),
                            Document_Reference = srcWs.Cells[i, 3].Value.ToString(),
                            Header_Information = srcWs.Cells[i, 4].Value.ToString(),
                            Pallet_Count = srcWs.Cells[i, 5].Value.ToString(),
                            Vehicle_Reg = srcWs.Cells[i, 6].Value.ToString(),
                            Trailer_Org = srcWs.Cells[i, 7].Value.ToString(),
                            Trailer_Sys = srcWs.Cells[i, 8].Value.ToString(),
                            Gmr = srcWs.Cells[i, 9].Value.ToString(),
                            Completed_By = srcWs.Cells[i, 10].Value.ToString(),
                            Comment = srcWs.Cells[i, 11].Value.ToString(),
                        });
                    }
                }
                catch (Exception innerException)
                {
                    // Handle the inner exception as needed
                    Console.WriteLine($"Inner Exception: {innerException}");

                    // Store the current value of 'i' in the exceptionIndex variable
                    exceptionIndex = i;

                    // Handle the exception here if needed, with the knowledge of the 'i' value
                    if (exceptionIndex != -1)
                    {
                        Console.WriteLine($"Error occurred at Transport List to Excel at index: {exceptionIndex}");
                        string dbExceptionPrName = "Transport Data To Excel";
                        InsertLogToDb.Exception(dbExceptionPrName, environment);
                        ExceptionLogToFile.Instance.WriteExceptionLog($"Error occurred at index {exceptionIndex}");
                        AlertEmail.Instance.Send(environment, exceptionIndex, dbExceptionPrName);
                    }

                    // Continue to the next iteration of the loop
                    continue;
                }
            }

            object misValue = Missing.Value;
            srcWb.Close(false, misValue, misValue);

            Console.WriteLine("Transport Data Transfer To Excel - Completed");

            return list;
           
        }
    }
}
using Microsoft.Office.Interop.Excel;
using Northern_Ireland_Shipments.Infrastructure.Smtp;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.Models.Queries;
using System.Globalization;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Northern_Ireland_Shipments.Application.Excel
{
    public class TransportListToExcel
    {
        public static void ListToExcel(string environment, Worksheet reportWs, List<TransportSrcFileModel> transportList)
        {

            int nRows = reportWs.UsedRange.Rows.Count;
            int exceptionIndex = 0;

            foreach (var item in transportList)
            {
                //DateTime itemDate = DateTime.ParseExact(item.Date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    
                string itemTrailer = item.Trailer_Sys;

                for (int i = 3; i <= nRows; i++)
                {
                    try
                    {
                        DateTime excelDate = DateTime.FromOADate(Convert.ToDouble((reportWs.Cells[i, 28] as Range).Value2));    //RP Shipped_Date
                        string excelTrailer = reportWs.Cells[i, 12].Value;  //RP Trailer

                        if (!excelTrailer.Contains("SUB"))
                        {
                            if (item.Date == excelDate && itemTrailer == excelTrailer)
                            {
                                reportWs.Cells[i, 1].Value = excelDate;
                                reportWs.Cells[i, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 2].Value = item.Client_Job_Number;
                                reportWs.Cells[i, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 3].Value = item.Document_Reference;
                                reportWs.Cells[i, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 4].Value = item.Header_Information;
                                reportWs.Cells[i, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 5].Value = item.Pallet_Count;
                                reportWs.Cells[i, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 6].Value = item.Vehicle_Reg;
                                reportWs.Cells[i, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 7].Value = item.Trailer_Org;
                                reportWs.Cells[i, 7].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 8].Value = item.Trailer_Sys;
                                reportWs.Cells[i, 8].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 9].Value = item.Gmr;
                                reportWs.Cells[i, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 10].Value = item.Completed_By;
                                reportWs.Cells[i, 10].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 11].Value = item.Comment;
                                reportWs.Cells[i, 11].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            }
                        }
                        else
                        {
                            excelTrailer = "SUB";
                            if (item.Date == excelDate && itemTrailer == excelTrailer)
                            {
                                reportWs.Cells[i, 1].Value = excelDate;
                                reportWs.Cells[i, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 2].Value = item.Client_Job_Number;
                                reportWs.Cells[i, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 3].Value = item.Document_Reference;
                                reportWs.Cells[i, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 4].Value = item.Header_Information;
                                reportWs.Cells[i, 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 5].Value = item.Pallet_Count;
                                reportWs.Cells[i, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 6].Value = item.Vehicle_Reg;
                                reportWs.Cells[i, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 7].Value = item.Trailer_Org;
                                reportWs.Cells[i, 7].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 8].Value = item.Trailer_Sys;
                                reportWs.Cells[i, 8].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 9].Value = item.Gmr;
                                reportWs.Cells[i, 9].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 10].Value = item.Completed_By;
                                reportWs.Cells[i, 10].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                reportWs.Cells[i, 11].Value = item.Comment;
                                reportWs.Cells[i, 11].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            }
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
                            Console.WriteLine($"Error occurred at Transport document's date at index: {exceptionIndex}");
                            string dbExceptionPrName = "Transport List To Excel";
                            InsertLogToDb.Exception(dbExceptionPrName, environment);
                            ExceptionLogToFile.Instance.WriteExceptionLog($"Error occurred at index {exceptionIndex}");
                            AlertEmail.Instance.Send(environment, exceptionIndex, dbExceptionPrName);
                        }

                        // Continue to the next iteration of the loop
                        continue;
                    }

                }
            }

            Console.WriteLine("Transport List To Excel - Completed");
        }
    }
}
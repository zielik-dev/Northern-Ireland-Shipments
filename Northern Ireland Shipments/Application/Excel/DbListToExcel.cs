using Microsoft.Office.Interop.Excel;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.Models.Queries;

namespace Northern_Ireland_Shipments.Application.Excel
{
    public class DbListToExcel
    {
        public static void TransferData(string environment, List<RpDbQueryModel> dbResult, Worksheet reportWs)
        {
            try
            {
                Console.WriteLine("Db data to Excel - started");
                int indexRow = 3;

                foreach (var item in dbResult)
                {
                    reportWs.Cells[indexRow, 12].Value = item.TRAILER;
                    reportWs.Cells[indexRow, 12].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 13].Value = item.ORDER_ID;
                    reportWs.Cells[indexRow, 13].NumberFormat = "0.00000000";
                    reportWs.Cells[indexRow, 13].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 14].Value = item.NAME;
                    reportWs.Cells[indexRow, 14].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 15].Value = item.ADDRESS1;
                    reportWs.Cells[indexRow, 16] = item.ADDRESS2;

                    reportWs.Cells[indexRow, 17].Value = item.TOWN;
                    reportWs.Cells[indexRow, 17].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 18].Value = item.POSTCODE;
                    reportWs.Cells[indexRow, 18].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 19].Value = item.SKU_ID;
                    reportWs.Cells[indexRow, 19].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 20].Value = item.QTY_SHIPPED;
                    reportWs.Cells[indexRow, 21] = item.TOTAL_VALUE;

                    reportWs.Cells[indexRow, 22].Value = item.PALLET_ID;
                    reportWs.Cells[indexRow, 22].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 23].Value = item.DESCRIPTION;

                    reportWs.Cells[indexRow, 24].Value = item.COM_CODE;
                    reportWs.Cells[indexRow, 24].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 25] = item.MANUFACTURE;
                    reportWs.Cells[indexRow, 25].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 26].Value = item.CONSIGNMENT;
                    reportWs.Cells[indexRow, 26].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    reportWs.Cells[indexRow, 27].Value = item.COUNTRY;
                    reportWs.Cells[indexRow, 27].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    if (item.SHIPPED_DATE != null)
                    {
                        DateTime dt = DateTime.ParseExact(item.SHIPPED_DATE, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                        reportWs.Cells[indexRow, 28].Value = dt;
                        reportWs.Cells[indexRow, 28].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    if (item.DELIVERED_BY_DATE != null)
                    {
                        DateTime dt = DateTime.ParseExact(item.DELIVERED_BY_DATE, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                        reportWs.Cells[indexRow, 29].Value = dt;
                        reportWs.Cells[indexRow, 29].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    }

                    reportWs.Cells[indexRow, 30].Value = item.EACH_VALUE;
                    reportWs.Cells[indexRow, 31].Value = item.INSTRUCTION;
                    indexRow++;
                }

                Console.WriteLine("Db List To Excel - Completed");
            }
            catch (Exception e)
            {
                string exception = e.ToString();
                string dbExceptionPrName = "Db List To Excel";
                Console.WriteLine($"Error: {exception}");
                InsertLogToDb.Exception(dbExceptionPrName, environment);
                ExceptionLogToFile.Instance.WriteExceptionLog(exception);
            }
        }
    }
}
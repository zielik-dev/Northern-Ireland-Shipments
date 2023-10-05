using Microsoft.Office.Interop.Excel;
using Northern_Ireland_Shipments.Infrastructure;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.Models.Queries;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Northern_Ireland_Shipments.Application.Excel
{
    public class ExcelMain : ConnectionStrings
    {
        public static void Run(string environment, List<RpDbQueryModel> dbResult, DateTime dt, string inboundFile)
        {
			try
			{
                Microsoft.Office.Interop.Excel.Application xlApp = new()
                {
                    //Visible = true,
                    //ScreenUpdating = true,
                    //DisplayAlerts = true
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };

                Workbook destWb = xlApp.Workbooks.Open(reportTemplate, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet reportWs = destWb.Worksheets.get_Item(sheetTemplate);

                string date = dt.ToString("dd/MM/yyyy");
                DateTime time = DateTime.ParseExact(date, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                reportWs.Cells[1, 9] = time;

                CleardownTemplate.Clear(environment, reportWs);
                DbListToExcel.TransferData(environment, dbResult, reportWs);
                var transportList = TransportExcelToList.ExcelToList(environment, xlApp, inboundFile, dt);

                TransportListToExcel.ListToExcel(environment, reportWs, transportList);

                object misValue = Missing.Value;
                destWb.Save();
                destWb.Close(false, misValue, misValue);

                int pid = -1;
                HandleRef hwnd = new(xlApp, (IntPtr)xlApp.Hwnd);
                GetWindowThreadProcessId(hwnd, out pid);

                xlApp.ScreenUpdating = true;
                xlApp.DisplayAlerts = true;
                xlApp.Quit();

                KillProcess(pid, "EXCEL");

                Console.WriteLine("Excel creation completed");
            }
            catch (Exception e)
            {
                string exception = e.ToString();
                Console.WriteLine($"Error: {exception}");
                string dbExceptionPrName = "Excel Creation";
                InsertLogToDb.Exception(dbExceptionPrName, environment);
                ExceptionLogToFile.Instance.WriteExceptionLog(exception);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowThreadProcessId(HandleRef handle, out int processId);
        public static void KillProcess(int pid, string processName)
        {
            Process[] AllProcesses = Process.GetProcessesByName(processName);
            foreach (Process process in AllProcesses)
            {
                if (process.Id == pid)
                {
                    process.Kill();
                }
            }
            AllProcesses = null;
        }
    }
}

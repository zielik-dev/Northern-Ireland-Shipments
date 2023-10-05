using Northern_Ireland_Shipments.Application.Db;
using Northern_Ireland_Shipments.Application.Excel;
using Northern_Ireland_Shipments.Infrastructure.FileBroker;
using Northern_Ireland_Shipments.Infrastructure.Smtp;
using Northern_Ireland_Shipments.Interfaces;
using Northern_Ireland_Shipments.Logs;

var watch = System.Diagnostics.Stopwatch.StartNew();
DateTime dt = DateTime.Now;

string environment = "Production";
//string environment = "Test";

//Database fetch
IRpDataExtractToList rpDataExtractToList = new RpDataExtractToList();
var dbResult = rpDataExtractToList.GetList();
int dbLines = dbResult.Count;
Console.WriteLine($"Db extract lines: {dbLines}");

//Transport src file
ISrcTransportFileCopy srcTransportFileCopy = new SrcTransportFileCopy();
string inboundFile = srcTransportFileCopy.CopySrcTransportFileToInbound();

//Excel
ExcelMain.Run(environment, dbResult, dt, inboundFile);

//AchiveFile
IArchiveTemplate archiveTemplate = new ArchiveTemplate();
archiveTemplate.CopyTemplateToArchive(environment, dt);

//Email
IEmailReport smtp = new EmailReport();
smtp.Send(environment);

//TV dashboard log
InsertLogToDb.Complete(environment);

//App measure
watch.Stop();
var elapsedMs = watch.ElapsedMilliseconds;
Console.WriteLine(elapsedMs);
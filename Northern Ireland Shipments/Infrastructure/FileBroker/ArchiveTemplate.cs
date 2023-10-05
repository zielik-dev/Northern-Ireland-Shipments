using Northern_Ireland_Shipments.Interfaces;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig;

namespace Northern_Ireland_Shipments.Infrastructure.FileBroker
{
    public class ArchiveTemplate : ConnectionStrings, IArchiveTemplate
    {
        private static ArchiveTemplate archiveTemplate;
        private readonly string archiveDir;

        public ArchiveTemplate()
        {
            archiveDir = ArchiveDir.Read();
        }

        public static ArchiveTemplate Instance
        {
            get
            {
                if (archiveTemplate == null)
                    archiveTemplate = new ArchiveTemplate();
                return archiveTemplate;
            }
        }

        public void CopyTemplateToArchive(string environment, DateTime dt)
        {
            try
            {
                string dateTimeStamp = dt.ToString("dd.MM.yyyy HH.mm - ");
                string fileName = Path.GetFileName(reportTemplate);
                string fileArchiveName = String.Concat(dateTimeStamp, fileName);

                string archiveFullPath = Path.Combine(archiveDir, archiveEndDir, fileArchiveName);

                if (File.Exists(reportTemplate))
                {
                    File.Copy(reportTemplate, archiveFullPath);
                }
            }
            catch (Exception e)
            {
                string exception = e.ToString();
                string dbExceptionPrName = "Report Email";
                InsertLogToDb.Exception(dbExceptionPrName, environment);
                ExceptionLogToFile.Instance.WriteExceptionLog(exception);
            }
        }
    }
}
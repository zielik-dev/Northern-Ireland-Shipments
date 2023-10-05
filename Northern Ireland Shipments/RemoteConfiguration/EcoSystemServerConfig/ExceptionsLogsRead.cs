using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig
{
    public class ExceptionsLogsRead : ConnectionStrings
    {
        public static string ExceptionsLogFile()
        {
            string directory = DirectoryRead();
            string fileName = FileNameRead();

            String str = String.Concat(directory, fileName);

            return str;
        }

        private static string DirectoryRead()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemGeneralConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("EcoSystemServerConfiguration/Directories/LogsDir");
            string str = xmlNodeList[0].InnerText;

            return str;
        }

        private static string FileNameRead()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemGeneralConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("EcoSystemServerConfiguration/Files/ExceptionsLogsFile");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}
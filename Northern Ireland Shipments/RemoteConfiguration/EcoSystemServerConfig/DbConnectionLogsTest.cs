using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig
{
    public class DbConnectionLogsTest : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemDbConnection);

            XmlNodeList xmlNodeList = xml.SelectNodes("/EcoSystemDbConfig/Test/Database/VmReportingLogsDatabaseMichal");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}
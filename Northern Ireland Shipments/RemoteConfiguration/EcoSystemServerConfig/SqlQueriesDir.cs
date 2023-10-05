using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig
{
    public class SqlQueriesDir : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemGeneralConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("/EcoSystemServerConfiguration/Directories/SqlQueriesDir");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}

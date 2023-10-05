using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig
{
    public class TemplatesDir : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemGeneralConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("/EcoSystemServerConfiguration/Directories/TemplatesDir");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}

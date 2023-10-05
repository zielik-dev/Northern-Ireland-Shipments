using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.EcoSystemServerConfig
{
    public class RpOracleDbConnectionProd : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemDbConnection);

            XmlNodeList xmlNodeList = xml.SelectNodes("/EcoSystemDbConfig/Production/Database/Dlx2009RpDatabase");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}
using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.SmtpConfig
{
    public class SmtpEmailRecipients : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/NorthernIrelandShipments");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
        public static string Alarm()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/Developer");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}

using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.SmtpConfig
{
    public class SmtpEmailSignature : ConnectionStrings
    {
        public static string ReadRegards()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/AppsEmailBodyComposition/Signature/Regards");
            string str = xmlNodeList[0].InnerText;

            return str;
        }

        public static string ReadSign()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/AppsEmailBodyComposition/Signature/Sign");
            string str = xmlNodeList[0].InnerText;

            return str;
        }

        public static string ReadLine()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/AppsEmailBodyComposition/Signature/Line");
            string str = xmlNodeList[0].InnerText;

            return str;
        }

        public static string ReadFooter()
        {
            XmlDocument xml = new();
            xml.Load(ecoSystemEmailDistributionConfig);

            XmlNodeList xmlNodeList = xml.SelectNodes("AppsEmailDistribution/AppsEmailBodyComposition/Signature/Footer");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}
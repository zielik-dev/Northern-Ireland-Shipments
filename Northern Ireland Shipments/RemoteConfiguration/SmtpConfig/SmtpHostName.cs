using Northern_Ireland_Shipments.Infrastructure;
using System.Xml;

namespace Northern_Ireland_Shipments.RemoteConfiguration.SmtpConfig
{
    public class SmtpHostName : ConnectionStrings
    {
        public static string Read()
        {
            XmlDocument xml = new();
            xml.Load(smtpConfigFile);

            XmlNodeList xmlNodeList = xml.SelectNodes("/GxoSmtpServerConfiguration/Host");
            string str = xmlNodeList[0].InnerText;

            return str;
        }
    }
}

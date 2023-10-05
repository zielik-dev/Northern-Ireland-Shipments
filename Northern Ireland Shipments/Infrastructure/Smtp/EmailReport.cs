using Northern_Ireland_Shipments.Interfaces;
using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.RemoteConfiguration.SmtpConfig;
using System.Net.Mail;

namespace Northern_Ireland_Shipments.Infrastructure.Smtp
{
    public class EmailReport : ConnectionStrings, IEmailReport
    {
        private static EmailReport emailReport;
        private static string _smptHostName, _senderAddress, _emailRecipients, _regards, _sign, _line, _footer;
        public EmailReport()
        {
            _smptHostName = SmtpHostName.Read();
            _senderAddress = SmtpSenderAddress.Read();
            _emailRecipients = SmtpEmailRecipients.Read();
            _regards = SmtpEmailSignature.ReadRegards();
            _sign = SmtpEmailSignature.ReadSign();
            _line = SmtpEmailSignature.ReadLine();
            _footer = SmtpEmailSignature.ReadFooter();
        }

        public static EmailReport Instance
        {
            get
            {
                if (emailReport == null)
                    emailReport = new EmailReport();
                return emailReport;
            }
        }

        public void Send(string environment)
        {
            try
            {

                MailAddress from = new(_senderAddress);
                Attachment attachment = new(reportTemplate);
                MailMessage message = new()
                {
                    From = from
                };

                message.To.Add(_emailRecipients);
                message.Subject = logTitle;
                message.Attachments.Add(attachment);
                message.IsBodyHtml = false;
                message.Body = @"Hello All," + Environment.NewLine + Environment.NewLine +
                                 $"Please find attached {logTitle} report." + Environment.NewLine + Environment.NewLine +
                                 "Please do not reply directly to this email and if escalating an issue please change subject of the email to avoid the risk of Outlook rule moving the email to specific folder that is less visited." + Environment.NewLine + Environment.NewLine +
                                 _regards + Environment.NewLine +
                                 _sign + Environment.NewLine +
                                 _line + Environment.NewLine +
                                 _footer;

                SmtpClient client = new()
                {
                    Host = _smptHostName
                };

                client.Send(message);
                Console.WriteLine("Email sent");
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

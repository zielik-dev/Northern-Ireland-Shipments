using Northern_Ireland_Shipments.Logs;
using Northern_Ireland_Shipments.RemoteConfiguration.SmtpConfig;
using System.Net.Mail;

namespace Northern_Ireland_Shipments.Infrastructure.Smtp
{
    internal class AlertEmail : ConnectionStrings
    {
        private static AlertEmail alertEmail;
        private static string _smptHostName, _senderAddress, _emailRecipients, _regards, _sign, _line, _footer;
        
        public AlertEmail()
        {
            _smptHostName = SmtpHostName.Read();
            _senderAddress = SmtpSenderAddress.Read();
            _emailRecipients = SmtpEmailRecipients.Alarm();
            _regards = SmtpEmailSignature.ReadRegards();
            _sign = SmtpEmailSignature.ReadSign();
            _line = SmtpEmailSignature.ReadLine();
            _footer = SmtpEmailSignature.ReadFooter();
        }

        public static AlertEmail Instance
        {
            get
            {
                if( alertEmail == null )
                    alertEmail = new AlertEmail();
                return alertEmail;
            }
        }

        public void Send(string environment, int exceptionIndex, string exceptionStage)
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
                message.Subject = $"ALERT: {logTitle}";
                message.Attachments.Add(attachment);
                message.IsBodyHtml = false;
                message.Body = @"Hello All," + Environment.NewLine + Environment.NewLine +
                                 $"Please find attached {logTitle} error location and line:" + Environment.NewLine + Environment.NewLine +
                                 $"Problem at stage: {exceptionStage}" + Environment.NewLine + 
                                 $"Problem at data index: {exceptionIndex}" + Environment.NewLine + Environment.NewLine +
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

using System;
using EASendMail;

namespace EWS_EmailSender
{
    public class SendEmailByEA
    {
        public static void SendEmail()
        {
            var oMail = new SmtpMail("TryIt")
            {
                From = Common.UserEmail,
                To = Common.RecipientEmail,
                Subject = "test email from c# project",
                TextBody = "this is a test email sent from c# project, do not reply"
            };
            
            // Your Exchange Server address/Office365 server address
            var oServer = new SmtpServer("https://outlook.office365.com/EWS/Exchange.asmx");

            // Set Exchange Web Service EWS - Exchange 2007/2010/2013/2016/Office365
            oServer.Protocol = ServerProtocol.ExchangeEWS;

            // User and password for Exchange authentication
            oServer.User = Common.UserEmail;
            oServer.Password = Common.Password;

            // By default, Exchange Web Service requires SSL connection
            oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

            Console.WriteLine("start to send email ...");

            var oSmtp = new SmtpClient();
            oSmtp.SendMail(oServer, oMail);

            Common.ShowMessage("email was sent successfully!");
        }
    }
}

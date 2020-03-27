using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailClass
{
    public class MailKitHelper
    {
        string account_bak = System.Configuration.ConfigurationManager.AppSettings["account_bak"];
        string password_bak = System.Configuration.ConfigurationManager.AppSettings["password_bak"];
        string _smtp = System.Configuration.ConfigurationManager.AppSettings["smtp"];
        int port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["port"]);

        public void Send(string[] tos, string title, string content)
        {

            var messageToSend = new MimeMessage();
            messageToSend.Sender = new MailboxAddress("jcet_test_cj01", account_bak);
            foreach (var to in tos)
            {
                messageToSend.To.Add(new MailboxAddress(to, to));
            }
            messageToSend.Subject = title;
            messageToSend.Body = new TextPart(TextFormat.Html) { Text = content };
            using (var smtp = new SmtpClient())
            {
                smtp.ServerCertificateValidationCallback = (s, c, h, e) => true;
                try
                {
                    smtp.Connect(_smtp, port, SecureSocketOptions.StartTls);
                    smtp.Authenticate(account_bak, password_bak);
                    smtp.Send(messageToSend);
                    smtp.Disconnect(true);
                }
                catch (Exception ex)
                {

                }
            }
        }
    }
}

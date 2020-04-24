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
        string account = System.Configuration.ConfigurationManager.AppSettings["account"];
        string password = System.Configuration.ConfigurationManager.AppSettings["password"];
        string account_jscc = System.Configuration.ConfigurationManager.AppSettings["account_jscc"];
        string password_jscc = System.Configuration.ConfigurationManager.AppSettings["password_jscc"];
        string _smtp = System.Configuration.ConfigurationManager.AppSettings["smtp"];
        int port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["port"]);

        public void Send(string[] tos, string title, string content, string type)
        {

            var messageToSend = new MimeMessage();
            if (type == "jscc")
            {
                messageToSend.From.Add(new MailboxAddress("测试项目组", account_jscc));
            }
            else
            {
                messageToSend.From.Add(new MailboxAddress("测试项目组", account));
            }
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
                    if (type == "jscc")
                    {
                        smtp.Authenticate(account_jscc, password_jscc);
                    }
                    else
                    {
                        smtp.Authenticate(account, password);
                    }
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

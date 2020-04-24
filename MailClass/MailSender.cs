using System.Text;
using System.Net.Mail;
using System.Net;
using System.Threading.Tasks;
using System;

namespace Twinkle.Framework.Utilities
{
    public class MailSender
    {
        string account = System.Configuration.ConfigurationManager.AppSettings["account_bak"];
        string password = System.Configuration.ConfigurationManager.AppSettings["password_bak"];
        string account_jscc = System.Configuration.ConfigurationManager.AppSettings["account_jscc"];
        string password_jscc = System.Configuration.ConfigurationManager.AppSettings["password_jscc"];
        string smtp = System.Configuration.ConfigurationManager.AppSettings["smtp"];
        int port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["port"]);
        //string account = "jcet_test_cj01@jcetglobal.com";
        //string password = "BR8KB@YMJDj";
        //string smtp = "smtp.partner.outlook.cn";
        //int port = 587;
        bool ssl = true;
        string displayName = "测试项目组";

        SmtpClient smtpClient;
        MailMessage mailMessage;
        public MailSender(string type)
        {
            InitClient(type);
        }

        //初始化客户端信息
        private void InitClient(string type)
        {
            smtpClient = new SmtpClient(smtp);
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            //smtpClient.Host = smtp;
            smtpClient.Port = port;
            smtpClient.EnableSsl = ssl;
            if (type == "jscc")
            {
                smtpClient.Credentials = new NetworkCredential(account_jscc, password_jscc);//用户名和密码
                mailMessage = new MailMessage
                {
                    From = new MailAddress(account_jscc, displayName),
                    BodyEncoding = System.Text.Encoding.UTF8,
                    IsBodyHtml = true,
                    Priority = MailPriority.Normal
                };
            }
            else
            {
                smtpClient.Credentials = new NetworkCredential(account, password);//用户名和密码
                mailMessage = new MailMessage
                {
                    From = new MailAddress(account, displayName),
                    BodyEncoding = System.Text.Encoding.UTF8,
                    IsBodyHtml = true,
                    Priority = MailPriority.Normal
                };
            }
             
        }

        /// <summary>
        /// 添加收件人邮箱
        /// </summary>
        /// <param name="tos"></param>
        public void AddTo(string[] tos)
        {
            foreach (var to in tos)
            {
                mailMessage.To.Add(to);
            }
        }

        /// <summary>
        /// 添加抄送人邮箱
        /// </summary>
        /// <param name="ccs"></param>
        public void AddCC(string[] ccs)
        {
            foreach (var cc in ccs)
            {
                mailMessage.CC.Add(cc);
            }
        }

        /// <summary>
        /// 添加密送人邮箱(其他收件者看不到)
        /// </summary>
        /// <param name="bccs"></param>
        public void AddBcc(string[] bccs)
        {
            foreach (var bcc in bccs)
            {
                mailMessage.Bcc.Add(bcc);
            }
        }


        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="subject">邮件标题</param>
        /// <param name="body">邮件正文</param>
        /// <returns></returns>
        public Task SendTask(string subject, string body)
        {

            mailMessage.Subject = subject;

            mailMessage.Body = body;

            try
            {
                return smtpClient.SendMailAsync(mailMessage);
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        public void Send(string subject, string body)
        {

            mailMessage.Subject = subject;

            mailMessage.Body = body;

            try
            {
                smtpClient.Send(mailMessage);
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}

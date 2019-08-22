using System.Text;
using System.Net.Mail;
using System.Net;
using System.Threading.Tasks;
using System;

namespace Twinkle.Framework.Utilities
{
    public class MailSender
    {
        string account = "jcet_admin_b2b_d02@cj-elec.com";
        string password = "spc123456!";
        string smtp = "smtp.cj-elec.com";
        bool ssl = false;
        int port = 25;
        string displayName = "测试项目组";

        SmtpClient smtpClient;
        MailMessage mailMessage;
        public MailSender()
        {
            InitClient();
        }

        //初始化客户端信息
        private void InitClient()
        {
            smtpClient = new SmtpClient();
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.Host = smtp;
            smtpClient.Port = port;
            smtpClient.Credentials = new NetworkCredential(account, password);//用户名和密码
            smtpClient.EnableSsl = ssl;
            mailMessage = new MailMessage
            {
                From = new MailAddress(account, displayName),
                BodyEncoding= Encoding.Default,
                IsBodyHtml = true,
                Priority = MailPriority.Normal
            };
        }

        /// <summary>
        /// 添加收件人邮箱
        /// </summary>
        /// <param name="tos"></param>
        public void AddTo(string [] tos)
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
        public Task Send(string subject,string body)
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
    }
}

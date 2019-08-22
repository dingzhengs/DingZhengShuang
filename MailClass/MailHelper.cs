using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;

namespace MailClass
{
    public class MailHelper
    {
        public MailHelper()
        {
            this.mailMessage = new MailMessage();
            this.smtpClient = new SmtpClient();
        }

        /// <summary>
        /// 设置发件人信息
        /// </summary>
        /// <param name="senderAddress">发件人邮箱</param>
        /// <param name="senderPassword">发件人邮箱密码</param>
        /// <param name="displayName">发件人显示名称</param>
        /// <param name="senderSmtp">SMTP地址</param>
        /// <param name="senderPort">SMTP端口号</param>
        public void SetSender(string senderAddress, string senderPassword, string displayName, string senderSmtp, int? senderPort = null)
        {
            mailMessage.From = new MailAddress(senderAddress, displayName);
            smtpClient.Host = senderSmtp;
            if (senderPort != null)
            {
                smtpClient.Port = Convert.ToInt32(senderPort);
            }
            smtpClient.Credentials = new NetworkCredential(senderAddress, senderPassword);
        }

        /// <summary>
        /// 设置收件人/抄送人列表信息
        /// </summary>
        /// <param name="recevicerList">收件人列表</param>
        /// <param name="ccList">抄送人列表</param>
        public void SetRecevier(Dictionary<string, string> recevicerList, Dictionary<string, string> ccList = null)
        {
            foreach (KeyValuePair<string, string> recevicer in recevicerList)
            {
                mailMessage.To.Add(new MailAddress(recevicer.Key, recevicer.Value));
            }
            if (ccList != null)
            {
                foreach (KeyValuePair<string, string> cc in ccList)
                {
                    mailMessage.CC.Add(new MailAddress(cc.Key, cc.Value));
                }
            }
        }
        public void SetRecevier(string singleRecevicer, string singleDisplayName = null)
        {
            if (singleDisplayName == null)
            {
                mailMessage.To.Add(singleRecevicer);
            }
            else
            {
                mailMessage.CC.Add(new MailAddress(singleRecevicer, singleDisplayName));
            }
        }

        /// <summary>
        /// 设置邮件及附件列表
        /// </summary>
        /// <param name="mailSubject">主题</param>
        /// <param name="mailBody">正文</param>
        /// <param name="attachList">附件地址列表</param>
        public void SetMail(string mailSubject, string mailBody, bool isBodyHtml = true, List<string> attachList = null)
        {
            mailMessage.Subject = mailSubject;
            mailMessage.Body = mailBody;
            mailMessage.IsBodyHtml = isBodyHtml;
            if (attachList != null)
            {
                foreach (string attachPath in attachList)
                {
                    mailMessage.Attachments.Add(new Attachment(attachPath));
                }
            }
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <returns></returns>
        public string SendMail(bool IsDispose = true)
        {
            try
            {
                smtpClient.EnableSsl = true;
                smtpClient.Send(mailMessage);
                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (IsDispose)
                {
                    foreach (var item in mailMessage.Attachments)
                    {
                        item.Dispose();
                    }
                    mailMessage.Dispose();
                    smtpClient.Dispose();
                }
            }
        }

        SmtpClient smtpClient;

        MailMessage mailMessage;

    }
}

using CSRedis;
using GlobalUtility.Core;
using MailClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Twinkle.Framework.Utilities;

namespace MailConsole
{
    class Program
    {

        public Program()
        {

        }

        static void Main(string[] args)
        {
            try
            {
                FileLog.WriteLog("启动邮件任务...");

                //EmailHelper email = new EmailHelper("smtphm.qiye.163.com", 25, "admin01@lkxinyun.com.cn", "Xinyun2020");
                //string recevicer = "418033729@qq.com";
                //string[] recevicerList = recevicer.Split(';');
                //email.SendEmail(recevicerList, "标题", "内容", "<b>这是html标记的文本</b>", "");

                MailKitHelper mailkit = new MailKitHelper();
                ////MailSender mail = new MailSender();
                string recevicer = "418033729@qq.com";
                string[] recevicerList = recevicer.Split(';');
                mailkit.Send(recevicerList, "mailkit邮件测试", "mailkit邮件测试");
                ////mail.AddTo(recevicerList);
                //try
                //{
                //    mailkit.Send(recevicerList, "mailkit邮件测试", "mailkit邮件测试");
                //}
                //catch (Exception ex)
                //{
                //    FileLog.WriteLog("mailkit邮件测试" + ex.Message);
                //}
                //try
                //{
                //    mail.Send("mailsender邮件测试", "mailsender邮件测试");
                //}
                //catch (Exception ex)
                //{
                //    FileLog.WriteLog("mailsender邮件测试" + ex.Message);
                //}
            }
            catch (Exception ex)
            {
                FileLog.WriteLog("邮件发送异常:" + ex.Message.ToString());
                FileLog.WriteLog("邮件发送失败！");
            }
        }
    }
}

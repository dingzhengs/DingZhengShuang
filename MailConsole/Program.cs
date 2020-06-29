using CSRedis;
using GlobalUtility.Core;
using MailClass;
using MailClass.WebServices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
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

                MailRules mail6 = new MailRules();
                RedisService redis = new RedisService();

                //MailHelper mail1 = new MailHelper();
                //mail1.testSend();

                //MailSender mail = new MailSender("");
                //string recevicer = "jcet_test_cj03@jcetglobal.com";
                //string[] recevicerList = recevicer.Split(';');
                //mail.AddTo(recevicerList);
                //mail.Send("测试邮件主题", "测试邮件内容");

                Hashtable pars = new Hashtable();
                pars["Lotid"] = "AFY948F043.001";
                string jobname = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/CIM/Service.asmx/getFTProgram", pars);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(jobname);
                String RetXml = doc.InnerText;


                //redis.Subscribe("PRR", msg =>
                //{
                //    mail6.prrRulseMailAlert(msg);
                //    //FileLog.WriteLog("PRR：" + msg);
                //});

                //redis.Subscribe("PTR", msg =>
                //{
                //    mail6.ptrRulseMailAlert(msg);
                //    //FileLog.WriteLog("PTR：" + msg);
                //});

                //redis.Subscribe("ECID", msg =>
                //{
                //    mail6.ecidRulseMailAlert(msg);
                //    //FileLog.WriteLog("ECID：" + msg);
                //});

                //redis.Subscribe("ECIDWAFER", msg =>
                //{
                //    mail6.ecidWaferRulseMailAlert(msg);
                //    //FileLog.WriteLog("ECIDWAFER：" + msg);
                //});

                //Console.ReadKey();
            }
            catch (Exception ex)
            {
                FileLog.WriteLog("邮件发送异常:" + ex.Message.ToString());
                FileLog.WriteLog("邮件发送失败！");
            }
        }

        public string Webkey()
        {
            //string datatime = sqlHelper.ExecuteScalar("").ToString();
            //string datatime = OraDBUtility.ExecuteScalar(CommandType.Text, "select to_char(sysdate,'YYYY-MM-DD') from dual", null).ToString();
            string value = "JCET" + DateTime.Now.ToString("yyyy-MM-dd");
            //string value = "JCET" + datatime;
            if (value == null || value == "")
            {
                return "";
            }
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            return Convert.ToBase64String(bytes);
        }
    }
}

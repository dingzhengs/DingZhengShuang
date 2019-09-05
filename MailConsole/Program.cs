using CSRedis;
using GlobalUtility.Core;
using MailClass;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

                redis.Subscribe("PRR", msg =>
                {
                    mail6.prrRulseMailAlert(msg);
                    //FileLog.WriteLog("PRR：" + msg);
                });

                redis.Subscribe("PTR", msg =>
                {
                    mail6.ptrRulseMailAlert(msg);
                    //FileLog.WriteLog("PTR：" + msg);
                });

                redis.Subscribe("ECID", msg =>
                {
                    mail6.ecidRulseMailAlert(msg);
                    //FileLog.WriteLog("ECID：" + msg);
                });

                redis.Subscribe("ECIDWAFER", msg =>
                {
                    mail6.ecidWaferRulseMailAlert(msg);
                    //FileLog.WriteLog("ECIDWAFER：" + msg);
                });

                Console.ReadKey();
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

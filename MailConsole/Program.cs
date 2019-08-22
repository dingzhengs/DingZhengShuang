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
                    Console.WriteLine("PRR：" + msg);
                });

                redis.Subscribe("PTR", msg =>
                {
                    mail6.ptrRulseMailAlert(msg);
                    Console.WriteLine("PRR：" + msg);
                });

                //redis.Subscribe("ECID", msg =>
                //{
                //    mail6.ecidRulseMailAlert(msg);
                //    Console.WriteLine("ECID：" + msg);
                //});

                Console.ReadKey();
            }
            catch (Exception ex)
            {
                FileLog.WriteLog("邮件发送异常:" + ex.Message.ToString());
                FileLog.WriteLog("邮件发送失败！");
            }
        }
    }
}

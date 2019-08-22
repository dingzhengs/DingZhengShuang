using CSRedis;
using GlobalUtility.Core;
using MailClass;
using MailConsole;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace MailService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
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
                    //mail6.ecidRulseMailAlert(msg);
                    FileLog.WriteLog("ECIDWAFER：" + msg);
                });

            }
            catch (Exception ex)
            {
                FileLog.WriteLog(ex.StackTrace);
            }
        }

        protected override void OnStop()
        {
        }
    }
}

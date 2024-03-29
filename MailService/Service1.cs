﻿using CSRedis;
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

                redis.Subscribe("LOSSMIR", msg =>
                {
                    mail6.lossMirMailAlert(msg);
                });

                redis.Subscribe("RCS", msg =>
                {
                    mail6.ptrRulseMailAlert(msg);
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

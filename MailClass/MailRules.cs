using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Data;
using MailClass;
using CSRedis;
using Twinkle.Framework.Utilities;
using Newtonsoft.Json.Linq;
using System.Collections;
using MailClass.WebServices;
using Oracle.ManagedDataAccess.Client;
using TDASTest;
using Newtonsoft.Json;

namespace MailConsole
{
    public class MailRules
    {
        MailKitHelper mailkit = new MailKitHelper();
        DatabaseManager dmgr = new DatabaseManager("oracle");
        DateTime sendTime = DateTime.MinValue;
        //MailKitHelper mailkit = new MailKitHelper();

        public void lossMirMailAlert(string result)
        {
            using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
            {
                conn.Open();
                OracleCommand ocmd;
                FileLog.WriteLog("lossmir:" + result);
                string mailTitle = "";
                int rulecheck = 0;
                string mailBody = "";
                string recevicer = "";
                string unlockrole = "";
                string unlockbm = "";
                string handleid = "";
                string stop = "";
                string handler_ip = "";
                int handler_port = 6000;
                var value = JToken.Parse(result).ToObject<dynamic>();
                string eqpname = MidStrEx(value.FILENAME.ToString(), "_", "_");

                ocmd = new OracleCommand(@"select stdfid from mir where stdfid in (select stdfid from stdffile where filename ='" + value.FILENAME.ToString() + "' )", conn);
                rulecheck = Convert.ToInt32(ocmd.ExecuteScalar()?.ToString());

                if (rulecheck > 0)
                {
                    return;
                }

                try
                {
                    ocmd = new OracleCommand(@"select HANDLER_IP from RTM_HANDLER_STOP where eqname='" + eqpname + "'", conn);
                    handler_ip = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select HANDLER_PORT from RTM_HANDLER_STOP where eqname='" + eqpname + "'", conn);
                    handler_port = Convert.ToInt32(ocmd.ExecuteScalar()?.ToString() == "" ? "6000" : ocmd.ExecuteScalar()?.ToString());

                    ocmd = new OracleCommand(@"select distinct RECEVICER from LOSS_MIR where EQPNAME='" + eqpname + "'", conn);
                    recevicer = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select distinct UNLOCKROLE from LOSS_MIR where EQPNAME='" + eqpname + "'", conn);
                    unlockrole = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select distinct UNLOCKBM from LOSS_MIR where EQPNAME='" + eqpname + "'", conn);
                    unlockbm = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select distinct HANDLEID from LOSS_MIR where EQPNAME='" + eqpname + "'", conn);
                    handleid = ocmd.ExecuteScalar()?.ToString();
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }

                try
                {
                    //MailSender mail = new MailSender();

                    recevicer = "admin01@lkxinyun.com.cn;" + recevicer;
                    string[] recevicerList = recevicer.Split(';');
                    mailTitle = "Pause Production_无Mir";
                    mailBody = value.FILENAME;

                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                    try
                    {
                        string json = JsonConvert.SerializeObject(new
                        {
                            Command = "STOP",
                            Title = mailTitle,
                            Message = mailBody
                        }).ToString();
                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                        stop = tc.Submit(json + Environment.NewLine);
                        FileLog.WriteLog("停机返回值:" + stop);
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                    }

                    //锁机成功插入数据，方便前台解锁
                    try
                    {
                        FileLog.WriteLog("---触发插表---");
                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{eqpname}','rtm','{handleid}','无Mir','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                        int res = ocmd.ExecuteNonQuery();
                        FileLog.WriteLog("插库反馈：" + res);
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                    }

                    FileLog.WriteLog("用C#SMTP发送邮件");
                    try
                    {
                        mailkit.Send(recevicerList, mailTitle, mailBody);
                        //mail.Send(mailTitle, mailBody);
                    }
                    catch (Exception e)
                    {
                        FileLog.WriteLog("错误：" + e.Message.ToString());
                        FileLog.WriteLog("使用mailkit发送邮件失败");
                    }
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
            }
        }

        public void ptrRulseMailAlert(string ptr_result)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    PTR_RESULT value = JToken.Parse(ptr_result).ToObject<PTR_RESULT>();

                    FileLog.WriteLog("ptr_result:" + ptr_result);
                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";
                    string handler_ip = "";
                    int handler_port = 6000;
                    string eqpid = value.EQPTID.ToString();
                    if (eqpid == "" || eqpid == null)
                    {
                        ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                        eqpid = ocmd.ExecuteScalar()?.ToString();
                    }
                    try
                    {
                        ocmd = new OracleCommand(@"select HANDLER_IP from RTM_HANDLER_STOP where eqname='" + value.EQPNAME + "'", conn);
                        handler_ip = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select HANDLER_PORT from RTM_HANDLER_STOP where eqname='" + value.EQPNAME + "'", conn);
                        handler_port = Convert.ToInt32(ocmd.ExecuteScalar()?.ToString() == "" ? "6000" : ocmd.ExecuteScalar()?.ToString());

                        ocmd = new OracleCommand(@"select ruletype from sys_rcs_rules_testrun where guid='" + value.GUID + "'", conn);
                        type = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select mail_list from sys_rcs_rules_testrun where guid='" + value.GUID + "'", conn);
                        recevicer = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select unlockrole from sys_rcs_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockrole = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select UNLOCKBM from sys_rcs_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockbm = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select name from sys_rcs_rules_testrun where guid='" + value.GUID + "'", conn);
                        handle_name = ocmd.ExecuteScalar()?.ToString();
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }

                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);

                    try
                    {
                        //MailSender mail = new MailSender();
                        recevicer = "admin01@lkxinyun.com.cn;" + recevicer;

                        string[] recevicerList = recevicer.Split(';');
                        //mail.AddTo(recevicerList);
                        switch (type)
                        {
                            case "固定值检查":
                                mailTitle = value.MAILTITLE;
                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']','" + value.REMARK + @"','" + value.PARTID + @"' from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda = new OracleDataAdapter(ocmd);
                                DataSet ds = new DataSet();
                                oda.Fill(ds);
                                DataTable dt = ds.Tables[0];

                                mailBody = dt.Rows[0][0].ToString() + "," + dt.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt.Rows[0][2].ToString() + "<br/>";
                                mailBody += "SITE=" + value.SITENUM + " VALUE=" + value.REMARK + " PARTID=" + value.PARTID + "<br/>";

                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    }
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                            case "一致性检查":
                                mailTitle = value.MAILTITLE;

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']',
substr('" + value.REMARK + @"',0,instr('" + value.REMARK + @"','、')-1),
substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','、')+1,length('" + value.REMARK + @"')) 
from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_ecid = new OracleDataAdapter(ocmd);
                                DataSet ds_ecid = new DataSet();
                                oda_ecid.Fill(ds_ecid);
                                DataTable dt_ecid = ds_ecid.Tables[0];

                                mailBody = dt_ecid.Rows[0][0].ToString() + "," + dt_ecid.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt_ecid.Rows[0][2].ToString() + "<br/>";

                                mailBody += dt_ecid.Rows[0][3].ToString() + "<br/>";
                                mailBody += dt_ecid.Rows[0][4].ToString();

                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    } 
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                            case "唯一性检查":
                                mailTitle = value.MAILTITLE;

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']',
substr('" + value.REMARK + @"',0,instr('" + value.REMARK + @"','、')-1),
substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','、')+1,length('" + value.REMARK + @"')) 
from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_wafer = new OracleDataAdapter(ocmd);
                                DataSet ds_wafer = new DataSet();
                                oda_wafer.Fill(ds_wafer);
                                DataTable dt_wafer = ds_wafer.Tables[0];

                                mailBody = dt_wafer.Rows[0][0].ToString() + "," + dt_wafer.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt_wafer.Rows[0][2].ToString() + "<br/>";

                                mailBody += dt_wafer.Rows[0][3].ToString() + "<br/>";
                                mailBody += dt_wafer.Rows[0][4].ToString();

                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    }
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                            case "公式计算":
                                mailTitle = value.MAILTITLE;

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']','" + value.REMARK + @"'
from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_cal = new OracleDataAdapter(ocmd);
                                DataSet ds_cal = new DataSet();
                                oda_cal.Fill(ds_cal);
                                DataTable dt_cal = ds_cal.Tables[0];

                                mailBody = dt_cal.Rows[0][0].ToString() + "," + dt_cal.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt_cal.Rows[0][2].ToString() + "<br/>";

                                mailBody += dt_cal.Rows[0][3].ToString();
                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    }
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                            case "连续多少次一样":
                                mailTitle = value.MAILTITLE;

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']','" + value.REMARK + @"'
from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_con = new OracleDataAdapter(ocmd);
                                DataSet ds_con = new DataSet();
                                oda_con.Fill(ds_con);
                                DataTable dt_con = ds_con.Tables[0];

                                mailBody = dt_con.Rows[0][0].ToString() + "," + dt_con.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt_con.Rows[0][2].ToString() + "<br/>";

                                mailBody += dt_con.Rows[0][3].ToString();

                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    }
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                            case "存在性检查":
                                mailTitle = value.MAILTITLE;

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.ruletype||'/'||t1.name||']','" + value.REMARK + @"'
from sys_rcs_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_exist = new OracleDataAdapter(ocmd);
                                DataSet ds_exist = new DataSet();
                                oda_exist.Fill(ds_exist);
                                DataTable dt_exist = ds_exist.Tables[0];

                                mailBody = dt_exist.Rows[0][0].ToString() + "," + dt_exist.Rows[0][1].ToString() + "<br/>";
                                mailBody += dt_exist.Rows[0][2].ToString() + "<br/>";

                                mailBody += dt_exist.Rows[0][3].ToString();

                                //触发停机
                                if (value.ISSTOP.ToString() == "1")
                                {
                                    //锁机成功插入数据，方便前台解锁
                                    try
                                    {
                                        FileLog.WriteLog("---触发插表---");
                                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,USERID,EQPTID,MAILTITLE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK,HANDLER_IP,HANDLER_PORT,MAILBODY) values ('{value.EQPNAME}','rtm','{eqpid}','{mailTitle}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}','{handler_ip}','{handler_port}','{mailBody}')", conn);
                                        int res = ocmd.ExecuteNonQuery();
                                        FileLog.WriteLog("插库反馈：" + res);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                                    }

                                    FileLog.WriteLog("开始触发停机，eqptId：" + value.EQPNAME);
                                    try
                                    {
                                        string json = JsonConvert.SerializeObject(new
                                        {
                                            Command = "STOP",
                                            Title = mailTitle,
                                            Message = mailBody
                                        }).ToString();
                                        TcpClient tc = new TcpClient(handler_ip, handler_port);
                                        stop = tc.Submit(json + Environment.NewLine);
                                        FileLog.WriteLog("停机返回值:" + stop);
                                    }
                                    catch (Exception ex)
                                    {
                                        FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                                    } 
                                }

                                FileLog.WriteLog("用C#SMTP发送邮件");
                                try
                                {
                                    mailkit.Send(recevicerList, mailTitle, mailBody);
                                    FileLog.WriteLog("用C#SMTP发送邮件成功");
                                    //mail.Send(mailTitle, mailBody);
                                }
                                catch (Exception e)
                                {
                                    FileLog.WriteLog("错误：" + e.Message.ToString());
                                    FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                }

                                break;

                        }
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog.WriteLog(ex.Message + ex.StackTrace);
            }
        }

        public string GetHtmlString(DataTable tbl)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<table width='100%' border='1' cellpadding='0' cellspacing='0' style='font-size: 12px' ");
            sb.Append(">");
            sb.Append("<tr valign='middle'>");
            sb.Append("<td nowrap><b><span>序号</span></b></td>");
            foreach (DataColumn column in tbl.Columns)
            {
                sb.Append("<td nowrap><b><span>" + column.ColumnName + "</span></b></td>");
            }
            sb.Append("</tr>");
            int iColsCount = tbl.Columns.Count;
            int rowsCount = tbl.Rows.Count - 1;
            for (int j = 0; j <= rowsCount; j++)
            {
                sb.Append("<tr>");
                sb.Append("<td>" + ((int)(j + 1)).ToString() + "</td>");
                for (int k = 0; k <= iColsCount - 1; k++)
                {
                    sb.Append("<td nowrap");
                    sb.Append(">");
                    object obj = tbl.Rows[j][k];
                    if (obj == DBNull.Value)
                    {
                        // 如果是NULL则在HTML里面使用一个空格替换之  
                        obj = "&nbsp;";
                    }
                    if (obj.ToString() == "")
                    {
                        obj = "&nbsp;";
                    }
                    string strCellContent = obj.ToString().Trim();
                    sb.Append("<span>" + strCellContent + "</span>");
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</TABLE>");
            return sb.ToString();
        }

        public static string MidStrEx(string sourse, string startstr, string endstr)
        {
            string result = string.Empty;
            int startindex, endindex;
            try
            {
                startindex = sourse.IndexOf(startstr);
                if (startindex == -1)
                {
                    return result;
                }
                string tmpstr = sourse.Substring(startindex + startstr.Length);
                endindex = tmpstr.IndexOf(endstr);
                if (endindex == -1)
                {
                    return result;
                }
                result = tmpstr.Remove(endindex);
            }
            catch (Exception ex)
            {
                FileLog.WriteLog("截断字符串:" + ex.Message);
            }
            return result;
        }

        public class PRR_RESULT
        {
            public string GUID { get; set; }
            public string EQPNAME { get; set; }
            public string STDFID { get; set; }
            public string DATETIME { get; set; }
            public string SITENUM { get; set; }
            public string REMARK { get; set; }
            public string EQPTID { get; set; }
            public string ISSTOP { get; set; }

        }

        public class PTR_RESULT
        {
            public string GUID { get; set; }
            public string EQPNAME { get; set; }
            public string STDFID { get; set; }
            public string DATETIME { get; set; }
            public string SITENUM { get; set; }
            public string REMARK { get; set; }
            public string EQPTID { get; set; }
            public string ISSTOP { get; set; }
            public string PARTID { get; set; }
            public string MAILTITLE { get; set; }
            public string PRODUCT { get; set; }
        }
    }
}


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
using System.Xml;

namespace MailConsole
{
    public class MailRules
    {
        MailHelper mailHelper = new MailHelper();
        DatabaseManager dmgr = new DatabaseManager("oracle");
        DateTime sendTime = DateTime.MinValue;
        MailKitHelper mailkit = new MailKitHelper();

        public void prrRulseMailAlert(string prr_result)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    PRR_RESULT value = JToken.Parse(prr_result).ToObject<PRR_RESULT>();
                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";
                    string sublot = "";
                    string mailtype = "";
                    string eqpid = value.EQPTID.ToString();
                    if (eqpid == "" || eqpid == null)
                    {
                        ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                        eqpid = ocmd.ExecuteScalar()?.ToString();
                    }

                    try
                    {
                        ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        type = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        recevicer = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockrole = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockbm = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        handle_name = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select sblotid from mir where stdfid='" + value.STDFID + "'", conn);
                        sublot = ocmd.ExecuteScalar()?.ToString();
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }

                    //如果是JSCC
                    ocmd = new OracleCommand(@"select count(*) from dual where fun_getplant('','" + value.EQPNAME + "')='JSCC'", conn);
                    int jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                    if (jscc_count > 0)
                    {
                        return;
                    }
                    FileLog.WriteLog("prr_result:" + prr_result);
                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);
                    if (value.ISSTOP.ToString() == "1")
                    {
                        FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                        Hashtable pars = new Hashtable();
                        pars["key"] = Webkey();
                        pars["userId"] = "shuxi_newserver";
                        pars["eqptId"] = eqpid;
                        pars["type"] = handle_name;
                        pars["lotId"] = "";
                        pars["Formname"] = "";
                        pars["Stepname"] = "";
                        FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + eqpid);
                        try
                        {
                            stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                            FileLog.WriteLog("PRR停机返回值:" + stop);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("PRR停机失败反馈：" + ex.Message.ToString());
                        }

                        //锁机成功插入数据，方便前台解锁
                        try
                        {
                            FileLog.WriteLog("---触发插表---");
                            ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                            int res = ocmd.ExecuteNonQuery();
                            FileLog.WriteLog("插库反馈：" + res);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                        }

                    }

                    try
                    {
                        MailSender mail = new MailSender("");
                        if (jscc_count > 0)
                        {
                            mail = new MailSender("jscc");
                            mailtype = "jscc";
                        }
                        else
                        {
                            mail = new MailSender("");
                        }
                        recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;
                        string[] recevicerList = recevicer.Split(';');
                        mail.AddTo(recevicerList);
                        switch (type)
                        {
                            case "BINCOUNTTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }
                                //                                mailTitle = dmgr.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
                                //from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ").ToString();

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case when count is not null then
case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end 
when maxvalue is not null then
case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end 
when minvalue is not null then
case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  end ||
' site='||'" + value.SITENUM + @"'||' value='||'" + value.REMARK + @"',
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_BINCOUNTTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_BINCOUNTTRIGGER = new DataSet();
                                oda_BINCOUNTTRIGGER.Fill(dt_BINCOUNTTRIGGER);
                                DataTable ds_BINCOUNTTRIGGER = dt_BINCOUNTTRIGGER.Tables[0];

                                mailBody = ds_BINCOUNTTRIGGER.Rows[0][0].ToString() + "," + ds_BINCOUNTTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_BINCOUNTTRIGGER.Rows[0][2].ToString() + "<br/>";
                                mailBody += ds_BINCOUNTTRIGGER.Rows[0][3].ToString() + "<br/>";
                                mailBody += ds_BINCOUNTTRIGGER.Rows[0][4].ToString();

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }
                                break;

                            case "SITETOSITEYIELDTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||'" + value.SITENUM + @"'||' value='||to_char(((to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1),0))-
to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1),0))))*100,'fm999990.0000000000')||'%',
'MAX_site='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,3)+1,instr('" + value.REMARK + @"',',',1,3)-instr('" + value.REMARK + @"','=',1,3)-1)||
' MAX_value='||to_number(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1))*100||'%'||
' MIN_site='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,4)+1,instr('" + value.REMARK + @"',',',1,4)-instr('" + value.REMARK + @"','=',1,4)-1)||
' MIN_value='||to_number(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1))*100||'%'||
' GAP='||to_char(((to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1),0))-
to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1),0))))*100,'fm999990.0000000000')||'%'
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_SITETOSITEYIELDTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_SITETOSITEYIELDTRIGGER = new DataSet();
                                oda_SITETOSITEYIELDTRIGGER.Fill(dt_SITETOSITEYIELDTRIGGER);
                                DataTable ds_SITETOSITEYIELDTRIGGER = dt_SITETOSITEYIELDTRIGGER.Tables[0];

                                //DataTable ds_SITETOSITEYIELDTRIGGER = dmgr.ExecuteDataTable();

                                mailBody = ds_SITETOSITEYIELDTRIGGER.Rows[0][0].ToString() + "," + ds_SITETOSITEYIELDTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEYIELDTRIGGER.Rows[0][2].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEYIELDTRIGGER.Rows[0][3].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEYIELDTRIGGER.Rows[0][4].ToString();

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "CONSECUTIVEBINTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||'" + value.SITENUM + @"'||' value='||'" + value.REMARK + @"',
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_CONSECUTIVEBINTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_CONSECUTIVEBINTRIGGER = new DataSet();
                                oda_CONSECUTIVEBINTRIGGER.Fill(dt_CONSECUTIVEBINTRIGGER);
                                DataTable ds_CONSECUTIVEBINTRIGGER = dt_CONSECUTIVEBINTRIGGER.Tables[0];

                                mailBody = ds_CONSECUTIVEBINTRIGGER.Rows[0][0].ToString() + "," + ds_CONSECUTIVEBINTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_CONSECUTIVEBINTRIGGER.Rows[0][2].ToString() + "<br/>";
                                mailBody += ds_CONSECUTIVEBINTRIGGER.Rows[0][3].ToString() + "<br/>";
                                mailBody += ds_CONSECUTIVEBINTRIGGER.Rows[0][4].ToString();

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
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

        public void ptrRulseMailAlert(string ptr_result)
        {

            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    PTR_RESULT value = JToken.Parse(ptr_result).ToObject<PTR_RESULT>();


                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";
                    string count = "";
                    string sublot = "";
                    string mailtype = "";
                    string eqpid = value.EQPTID.ToString();
                    if (eqpid == "" || eqpid == null)
                    {
                        ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                        eqpid = ocmd.ExecuteScalar()?.ToString();
                    }

                    try
                    {
                        ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        type = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        recevicer = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockrole = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockbm = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        handle_name = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select sblotid from mir where stdfid='" + value.STDFID + "'", conn);
                        sublot = ocmd.ExecuteScalar()?.ToString();
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }

                    //如果是JSCC
                    ocmd = new OracleCommand(@"select count(*) from dual where fun_getplant('','" + value.EQPNAME + "')='JSCC'", conn);
                    int jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                    if (jscc_count > 0)
                    {
                        return;
                    }
                    FileLog.WriteLog("ptr_result:" + ptr_result);
                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);
                    if (value.ISSTOP.ToString() == "1")
                    {

                        FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                        Hashtable pars = new Hashtable();
                        pars["key"] = Webkey();
                        pars["userId"] = "shuxi_newserver";
                        pars["eqptId"] = eqpid;
                        pars["type"] = handle_name;
                        pars["lotId"] = "";
                        pars["Formname"] = "";
                        pars["Stepname"] = "";
                        FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + eqpid);
                        try
                        {
                            stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                            FileLog.WriteLog("PTR停机返回值:" + stop);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("PTR停机失败反馈：" + ex.Message.ToString());
                        }

                        //锁机成功插入数据，方便前台解锁
                        try
                        {
                            FileLog.WriteLog("---触发插表---");
                            ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                            int res = ocmd.ExecuteNonQuery();
                            FileLog.WriteLog("插库反馈：" + res);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                        }
                    }

                    try
                    {
                        MailSender mail = new MailSender("");
                        if (jscc_count > 0)
                        {
                            mail = new MailSender("jscc");
                            mailtype = "jscc";
                        }
                        else
                        {
                            mail = new MailSender("");
                        }
                        recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;

                        string[] recevicerList = recevicer.Split(';');
                        mail.AddTo(recevicerList);
                        switch (type)
                        {
                            case "PARAMETRICTESTSTATISTICTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||'" + value.SITENUM + @"'||' value='||'" + value.REMARK + @"',
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_PARAMETRICTESTSTATISTICTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_PARAMETRICTESTSTATISTICTRIGGER = new DataSet();
                                oda_PARAMETRICTESTSTATISTICTRIGGER.Fill(dt_PARAMETRICTESTSTATISTICTRIGGER);
                                DataTable ds_PARAMETRICTESTSTATISTICTRIGGER = dt_PARAMETRICTESTSTATISTICTRIGGER.Tables[0];

                                //DataTable ds_PARAMETRICTESTSTATISTICTRIGGER = dmgr.ExecuteDataTable();

                                mailBody = ds_PARAMETRICTESTSTATISTICTRIGGER.Rows[0][0].ToString() + "," + ds_PARAMETRICTESTSTATISTICTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Rows[0][2].ToString() + "<br/>";
                                mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Rows[0][3].ToString() + "<br/>";
                                mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Rows[0][4].ToString();

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }
                                break;

                            case "SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||'" + value.SITENUM + @"'||' value='||
to_char((to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1),0))-
to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1),0))),'fm999990.0000000000'),
'MAX_site='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,3)+1,instr('" + value.REMARK + @"',',',1,3)-instr('" + value.REMARK + @"','=',1,3)-1)||
' MAX_value='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1)||
' MIN_site='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,4)+1,instr('" + value.REMARK + @"',',',1,4)-instr('" + value.REMARK + @"','=',1,4)-1)||
' MIN_value='||substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1)||
' GAP='||
to_char((to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,1)+1,instr('" + value.REMARK + @"',',',1,1)-instr('" + value.REMARK + @"','=',1,1)-1),0))-
to_number(NVL(substr('" + value.REMARK + @"',instr('" + value.REMARK + @"','=',1,2)+1,instr('" + value.REMARK + @"',',',1,2)-instr('" + value.REMARK + @"','=',1,2)-1),0))),'fm999990.0000000000')
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER = new DataSet();
                                oda_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Fill(dt_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER);
                                DataTable ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER = dt_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0];

                                mailBody = ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Rows[0][0].ToString() + "," + ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Rows[0][2].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Rows[0][3].ToString() + "<br/>";
                                mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Rows[0][4].ToString();

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "PTSADDTRIGGER":
                                ocmd = new OracleCommand(@"select count from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                                count = ocmd.ExecuteScalar()?.ToString();

                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',count,'" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_PTSADDTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_PTSADDTRIGGER = new DataSet();
                                oda_PTSADDTRIGGER.Fill(dt_PTSADDTRIGGER);
                                DataTable ds_PTSADDTRIGGER = dt_PTSADDTRIGGER.Tables[0];

                                mailBody = ds_PTSADDTRIGGER.Rows[0][0].ToString() + "," + ds_PTSADDTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_PTSADDTRIGGER.Rows[0][2].ToString() + "<br/>";
                                string[] sArray = ds_PTSADDTRIGGER.Rows[0][4].ToString().Split(',');
                                mailBody += "TouchDown=" + count + " LowLimit=" + sArray[0] + " HighLimit=" + sArray[sArray.Length - 1] + " site=" + value.SITENUM + "<br/>";
                                for (int i = 0; i < sArray.Length; i++)
                                {
                                    mailBody += "Unit" + (i + 1).ToString() + "=" + sArray[i] + "<br/>";
                                }

                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "PTSCUTTRIGGER":
                                ocmd = new OracleCommand(@"select count from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                                count = ocmd.ExecuteScalar()?.ToString();

                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',count,'" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_PTSCUTTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_PTSCUTTRIGGER = new DataSet();
                                oda_PTSCUTTRIGGER.Fill(dt_PTSCUTTRIGGER);
                                DataTable ds_PTSCUTTRIGGER = dt_PTSCUTTRIGGER.Tables[0];

                                mailBody = ds_PTSCUTTRIGGER.Rows[0][0].ToString() + "," + ds_PTSCUTTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_PTSCUTTRIGGER.Rows[0][2].ToString() + "<br/>";
                                string[] ptscuttriggerArray = ds_PTSCUTTRIGGER.Rows[0][4].ToString().Split(',');
                                mailBody += "TouchDown=" + count + " LowLimit=" + ptscuttriggerArray[0] + " HighLimit=" + ptscuttriggerArray[ptscuttriggerArray.Length - 1] + " site=" + value.SITENUM + "<br/>";
                                for (int i = 0; i < ptscuttriggerArray.Length; i++)
                                {
                                    mailBody += "Unit" + (i + 1).ToString() + "=" + ptscuttriggerArray[i] + "<br/>";
                                }
                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "OSPINCOUNTTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:','" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_OSPINCOUNTTRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_OSPINCOUNTTRIGGER = new DataSet();
                                oda_OSPINCOUNTTRIGGER.Fill(dt_OSPINCOUNTTRIGGER);
                                DataTable ds_OSPINCOUNTTRIGGER = dt_OSPINCOUNTTRIGGER.Tables[0];

                                mailBody = ds_OSPINCOUNTTRIGGER.Rows[0][0].ToString() + "," + ds_OSPINCOUNTTRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_OSPINCOUNTTRIGGER.Rows[0][2].ToString() + "<br/>";
                                string[] OSPINCOUNTTRIGGERArray = ds_OSPINCOUNTTRIGGER.Rows[0][3].ToString().Split(',');
                                mailBody += "Unit=" + value.SITENUM + " Site =" + value.PARTID + "<br/>";
                                for (int i = 0; i < OSPINCOUNTTRIGGERArray.Length; i++)
                                {
                                    mailBody += OSPINCOUNTTRIGGERArray[i] + "<br/>";
                                }
                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "OSPINCONSECUTIVETRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:','" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_OSPINCONSECUTIVETRIGGER = new OracleDataAdapter(ocmd);
                                DataSet dt_OSPINCONSECUTIVETRIGGER = new DataSet();
                                oda_OSPINCONSECUTIVETRIGGER.Fill(dt_OSPINCONSECUTIVETRIGGER);
                                DataTable ds_OSPINCONSECUTIVETRIGGER = dt_OSPINCONSECUTIVETRIGGER.Tables[0];

                                mailBody = ds_OSPINCONSECUTIVETRIGGER.Rows[0][0].ToString() + "," + ds_OSPINCONSECUTIVETRIGGER.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_OSPINCONSECUTIVETRIGGER.Rows[0][2].ToString() + "<br/>";
                                string[] OSPINCONSECUTIVETRIGGERArray = ds_OSPINCONSECUTIVETRIGGER.Rows[0][3].ToString().Split(',');
                                mailBody += "Unit=" + value.SITENUM + " Site=" + (Convert.ToDouble(value.PARTID) - 1).ToString() + "<br/>";
                                for (int i = 0; i < OSPINCONSECUTIVETRIGGERArray.Length; i++)
                                {
                                    mailBody += OSPINCONSECUTIVETRIGGERArray[i] + "<br/>";
                                }
                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "SIGMATRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:','" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_sigma = new OracleDataAdapter(ocmd);
                                DataSet dt_sigma = new DataSet();
                                oda_sigma.Fill(dt_sigma);
                                DataTable ds_sigma = dt_sigma.Tables[0];

                                mailBody = ds_sigma.Rows[0][0].ToString() + "," + ds_sigma.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_sigma.Rows[0][2].ToString() + "<br/>";
                                mailBody += "site=" + value.SITENUM + " value=" + ds_sigma.Rows[0][3].ToString() + "<br/>";
                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
                                }

                                break;

                            case "PTRCONSECUTIVEBINTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
                                if (mailTitle == "" || mailTitle == null)
                                {
                                    mailTitle = value.MAILTITLE + ".";
                                }

                                ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:','" + value.REMARK + @"' from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                                OracleDataAdapter oda_SPEC = new OracleDataAdapter(ocmd);
                                DataSet dt_SPEC = new DataSet();
                                oda_SPEC.Fill(dt_SPEC);
                                DataTable ds_SPEC = dt_SPEC.Tables[0];

                                mailBody = ds_SPEC.Rows[0][0].ToString() + "," + ds_SPEC.Rows[0][1].ToString() + "<br/>";
                                mailBody += ds_SPEC.Rows[0][2].ToString() + "<br/>";
                                mailBody += "site=" + value.SITENUM + " value=" + ds_SPEC.Rows[0][3].ToString() + "<br/>";
                                try
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog(value.MAILTITLE + "，错误：" + ex.Message.ToString());
                                    FileLog.WriteLog(value.MAILTITLE + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog(value.MAILTITLE + "，错误：" + e.Message.ToString());
                                        FileLog.WriteLog(value.MAILTITLE + "：使用C#SMTP发送邮件失败");
                                    }
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

        public void ecidRulseMailAlert(string result)
        {
            using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
            {
                conn.Open();
                OracleCommand ocmd;

                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string type = "";
                string unlockrole = "";
                string unlockbm = "";
                string handle_name = "";
                string jscc_stop = "";
                string mailtype = "";
                var value = JToken.Parse(result).ToObject<dynamic>();
                string eqpid = value.EQPTID.ToString();
                if (eqpid == "" || eqpid == null)
                {
                    ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                    eqpid = ocmd.ExecuteScalar()?.ToString();
                }

                try
                {
                    ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    type = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    recevicer = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockrole = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockbm = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    handle_name = ocmd.ExecuteScalar()?.ToString();
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }

                //如果是JSCC的停机
                ocmd = new OracleCommand(@"select count(*) from dual where fun_getplant('','" + value.NODENAM + "')='JSCC'", conn);
                int jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                if (jscc_count > 0)
                {
                    return;
                }
                FileLog.WriteLog("ecid:" + result);
                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid + "；EQPNAME：" + value.NODENAM);
                Hashtable pars = new Hashtable();
                pars["key"] = Webkey();
                pars["userId"] = "shuxi_newserver";
                pars["eqptId"] = eqpid;
                pars["type"] = handle_name;
                pars["lotId"] = "";
                pars["Formname"] = "";
                pars["Stepname"] = "";
                string stop = "";
                try
                {
                    stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                    FileLog.WriteLog("ECID停机返回值:" + stop);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("ECID停机失败反馈：" + ex.Message.ToString());
                }

                //锁机成功插入数据，方便前台解锁
                try
                {
                    FileLog.WriteLog("---触发插表---");
                    ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                    int res = ocmd.ExecuteNonQuery();
                    FileLog.WriteLog("插库反馈：" + res);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                }
                try
                {
                    MailSender mail = new MailSender("");
                    if (jscc_count > 0)
                    {
                        mail = new MailSender("jscc");
                        mailtype = "jscc";
                    }
                    else
                    {
                        mail = new MailSender("");
                    }
                    recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //                    ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from v_eq_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //                    mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECID_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.FLOWID + "_" + value.NODENAM + "_" + value.DATETIME;

                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECID|JCET_ECID_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.RESULT + "<br/>";
                    mailBody += value.REFS;
                    try
                    {
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件...");
                        mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(mailTitle + "，错误：" + ex.Message.ToString());
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                        try
                        {
                            mail.Send(mailTitle, mailBody);
                        }
                        catch (Exception e)
                        {
                            FileLog.WriteLog(mailTitle + "，错误：" + e.Message.ToString());
                            FileLog.WriteLog(mailTitle + "：使用C#SMTP发送邮件失败");
                        }
                    }

                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
            }
        }

        public void ecidWaferRulseMailAlert(string result)
        {
            using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
            {
                conn.Open();
                OracleCommand ocmd;

                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string unlockrole = "";
                string type = "";
                string unlockbm = "";
                string mailtype = "";
                string handle_name = "";
                var value = JToken.Parse(result).ToObject<dynamic>();
                string eqpid = value.EQPTID.ToString();
                if (eqpid == "" || eqpid == null)
                {
                    ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                    eqpid = ocmd.ExecuteScalar()?.ToString();
                }

                try
                {
                    ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    type = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    recevicer = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockrole = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockbm = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    handle_name = ocmd.ExecuteScalar()?.ToString();
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }

                //如果是JSCC的停机
                ocmd = new OracleCommand(@"select count(*) from dual where fun_getplant('','" + value.NODENAM + "')='JSCC'", conn);
                int jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                if (jscc_count > 0)
                {
                    return;
                }
                FileLog.WriteLog("ecidwafer:" + result);
                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                Hashtable pars = new Hashtable();
                pars["key"] = Webkey();
                pars["userId"] = "shuxi_newserver";
                pars["eqptId"] = eqpid;
                pars["type"] = handle_name;
                pars["lotId"] = "";
                pars["Formname"] = "";
                pars["Stepname"] = "";
                FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + eqpid);
                string stop = "";
                try
                {
                    stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                    FileLog.WriteLog("ECIDWAFER停机返回值:" + stop);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("ECIDWAFER停机失败反馈：" + ex.Message.ToString());
                }

                //锁机成功插入数据，方便前台解锁
                try
                {
                    FileLog.WriteLog("---触发插表---");
                    ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                    int res = ocmd.ExecuteNonQuery();
                    FileLog.WriteLog("插库反馈：" + res);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                }

                try
                {
                    MailSender mail = new MailSender("");
                    if (jscc_count > 0)
                    {
                        mail = new MailSender("jscc");
                        mailtype = "jscc";
                    }
                    else
                    {
                        mail = new MailSender("");
                    }
                    recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;
                    //recevicer = "kai.guo@shu-xi.com;zhengshuang.ding@shu-xi.com;jun.lai@cj-elec.com;tdas_it.list@cj-elec.com";
                    //recevicer = "zhengshuang.ding@shu-xi.com";
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECIDWAFER_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.FLOWID + "_" + value.NODENAM + "_" + value.DATETIME;


                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECIDWAFER|JCET_ECIDWAFER_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.REFS + "<br/>";
                    mailBody += value.RESULT;
                    try
                    {
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件...");
                        mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(mailTitle + "，错误：" + ex.Message.ToString());
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                        try
                        {
                            mail.Send(mailTitle, mailBody);
                        }
                        catch (Exception e)
                        {
                            FileLog.WriteLog(mailTitle + "，错误：" + e.Message.ToString());
                            FileLog.WriteLog(mailTitle + "：使用C#SMTP发送邮件失败");
                        }
                    }

                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
            }
        }

        public void ecidWafer_akjRulseMailAlert(string result)
        {
            using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
            {
                conn.Open();
                OracleCommand ocmd;

                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string unlockrole = "";
                string type = "";
                string unlockbm = "";
                string handle_name = "";
                var value = JToken.Parse(result).ToObject<dynamic>();
                string eqpid = value.EQPTID.ToString();
                if (eqpid == "" || eqpid == null)
                {
                    ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                    eqpid = ocmd.ExecuteScalar()?.ToString();
                }

                try
                {
                    ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    type = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    recevicer = ocmd.ExecuteScalar()?.ToString();

                    ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockrole = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    unlockbm = ocmd.ExecuteScalar()?.ToString();
                    ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                    handle_name = ocmd.ExecuteScalar()?.ToString();
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
                FileLog.WriteLog("ecidwaferakj:" + result);
                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                Hashtable pars = new Hashtable();
                pars["key"] = Webkey();
                pars["userId"] = "shuxi_newserver";
                pars["eqptId"] = eqpid;
                pars["type"] = handle_name;
                pars["lotId"] = "";
                pars["Formname"] = "";
                pars["Stepname"] = "";
                FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + eqpid);
                string stop = "";
                try
                {
                    stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                    FileLog.WriteLog("ECIDAKJ停机返回值:" + stop);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("ECIDAKJ停机失败反馈：" + ex.Message.ToString());
                }

                //锁机成功插入数据，方便前台解锁
                try
                {
                    FileLog.WriteLog("---触发插表---");
                    ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                    int res = ocmd.ExecuteNonQuery();
                    FileLog.WriteLog("插库反馈：" + res);
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                }

                try
                {
                    MailSender mail = new MailSender("");
                    recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //                    ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from v_eq_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //                    mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECIDWAFER-AKJ_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.FLOWID + "_" + value.NODENAM + "_" + value.DATETIME;


                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECIDWAFER-AKJ|JCET_ECIDWAFER-AKJ_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.REFS + "<br/>";
                    mailBody += value.RESULT;
                    try
                    {
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件...");
                        mailkit.Send(recevicerList, mailTitle, mailBody, "");

                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(mailTitle + "，错误：" + ex.Message.ToString());
                        FileLog.WriteLog(mailTitle + "：使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                        try
                        {
                            mail.Send(mailTitle, mailBody);
                        }
                        catch (Exception e)
                        {
                            FileLog.WriteLog(mailTitle + "，错误：" + e.Message.ToString());
                            FileLog.WriteLog(mailTitle + "：使用C#SMTP发送邮件失败");
                        }
                    }

                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
            }
        }

        public void touchdownRulseMailAlert(string ptr_result)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    PTR_RESULT value = JToken.Parse(ptr_result).ToObject<PTR_RESULT>();


                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";
                    string sublot = "";
                    string mailtype = "";
                    string eqpid = value.EQPTID.ToString();
                    if (eqpid == "" || eqpid == null)
                    {
                        ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                        eqpid = ocmd.ExecuteScalar()?.ToString();
                    }

                    try
                    {
                        ocmd = new OracleCommand(@"select type from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        type = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select mail_list from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        recevicer = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select unlockrole from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockrole = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select UNLOCKBM from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        unlockbm = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select name from sys_rules_testrun where guid='" + value.GUID + "'", conn);
                        handle_name = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select sblotid from mir where stdfid='" + value.STDFID + "'", conn);
                        sublot = ocmd.ExecuteScalar()?.ToString();
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }

                    //如果是JSCC
                    ocmd = new OracleCommand(@"select count(*) from dual where fun_getplant('','" + value.EQPNAME + "')='JSCC'", conn);
                    int jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                    if (jscc_count > 0)
                    {
                        return;
                    }
                    FileLog.WriteLog("ptr_result:" + ptr_result);
                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);
                    if (value.ISSTOP.ToString() == "1")
                    {
                        FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                        Hashtable pars = new Hashtable();
                        pars["key"] = Webkey();
                        pars["userId"] = "shuxi_newserver";
                        pars["eqptId"] = eqpid;
                        pars["type"] = handle_name;
                        pars["lotId"] = "";
                        pars["Formname"] = "";
                        pars["Stepname"] = "";
                        FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + eqpid);
                        try
                        {
                            stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                            FileLog.WriteLog("PTR停机返回值:" + stop);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("PTR停机失败反馈：" + ex.Message.ToString());
                        }

                        //锁机成功插入数据，方便前台解锁
                        try
                        {
                            FileLog.WriteLog("---触发插表---");
                            ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                            int res = ocmd.ExecuteNonQuery();
                            FileLog.WriteLog("插库反馈：" + res);
                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                        }

                    }

                    try
                    {
                        MailSender mail = new MailSender("");
                        jscc_count = Convert.ToInt32(ocmd.ExecuteScalar());
                        if (jscc_count > 0)
                        {
                            mail = new MailSender("jscc");
                            mailtype = "jscc";
                        }
                        else
                        {
                            mail = new MailSender("");
                        }
                        recevicer = "JCET_D3_IT_TDAS.LIST@jcetglobal.com;jcet_test_cj03@jcetglobal.com;" + recevicer;
                        string[] recevicerList = recevicer.Split(';');
                        mail.AddTo(recevicerList);

                        ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||substr(t2.FLOWID,0,2)||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                        mailTitle = ocmd.ExecuteScalar()?.ToString();

                        ocmd = new OracleCommand(@"select '" + value.DATETIME + @"',t1.action,
'['||t1.type||'/'||t1.name||']:',
'TouchDown=1  Site='||'" + value.SITENUM + @"'
from sys_rules_testrun t1 where t1.guid='" + value.GUID + "'", conn);

                        OracleDataAdapter oda_BINCOUNTTRIGGER = new OracleDataAdapter(ocmd);
                        DataSet dt_BINCOUNTTRIGGER = new DataSet();
                        oda_BINCOUNTTRIGGER.Fill(dt_BINCOUNTTRIGGER);
                        DataTable ds_BINCOUNTTRIGGER = dt_BINCOUNTTRIGGER.Tables[0];

                        mailBody = ds_BINCOUNTTRIGGER.Rows[0][0].ToString() + "," + ds_BINCOUNTTRIGGER.Rows[0][1].ToString() + "<br/>";
                        mailBody += ds_BINCOUNTTRIGGER.Rows[0][2].ToString() + "<br/>";
                        mailBody += ds_BINCOUNTTRIGGER.Rows[0][3].ToString();

                        try
                        {
                            FileLog.WriteLog("使用mailkit发送邮件...");
                            mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                        }
                        catch (Exception ex)
                        {
                            FileLog.WriteLog("错误：" + ex.Message.ToString());
                            FileLog.WriteLog("使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                            try
                            {
                                mail.Send(mailTitle, mailBody);
                            }
                            catch (Exception e)
                            {
                                FileLog.WriteLog("错误：" + e.Message.ToString());
                                FileLog.WriteLog("使用C#SMTP发送邮件失败");
                            }
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

        public void touchdownDiffJOBNAM(string ptr_result)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    DIFF_RESULT value = JToken.Parse(ptr_result).ToObject<DIFF_RESULT>();

                    string jobname = "";
                    string stop = "";
                    string mailTitle = "";
                    string mailBody = "";
                    string recevicer = "";
                    string unlockrole = "";
                    string unlockbm = "PTE";
                    string handle_name = "MES的测试程序名跟MIR里的测试程序名不一致";
                    string mailtype = "";
                    string eqpid = value.EQPTID.ToString();
                    string retXml = "";
                    //if (eqpid == "" || eqpid == null)
                    //{
                    //    ocmd = new OracleCommand(@"select handid from sdr where stdfid='" + value.STDFID + "'", conn);
                    //    eqpid = ocmd.ExecuteScalar()?.ToString();
                    //}

                    try
                    {
                        ocmd = new OracleCommand(@"select maillist from a_eqpmail where device_group=fun_getplant('" + value.SBLOTID + "','" + value.EQPNAME + "')", conn);
                        recevicer = ocmd.ExecuteScalar()?.ToString();
                        ocmd = new OracleCommand(@"select fun_getplant('" + value.SBLOTID + "','" + value.EQPNAME + "') from dual", conn);
                        unlockrole = ocmd.ExecuteScalar()?.ToString();
                    }
                    catch (Exception ex)
                    {
                        FileLog.WriteLog(ex.Message + ex.StackTrace);
                    }

                    if (!unlockrole.Contains("BGA其他"))
                    {
                        return;
                    }

                    try
                    {
                        Hashtable par = new Hashtable();
                        par["Lotid"] = value.SBLOTID;
                        jobname = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/CIM/Service.asmx/getFTProgram", par);
                        FileLog.WriteLog("获取Mes程序名:" + jobname);
                        XmlDocument doc = new XmlDocument();
                        doc.LoadXml(jobname);
                        retXml = doc.InnerText;
                        if (retXml != value.JOBNAM)
                        {
                            //FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + eqpid);
                            //Hashtable pars = new Hashtable();
                            //pars["key"] = Webkey();
                            //pars["userId"] = "shuxi_newserver";
                            //pars["eqptId"] = eqpid;
                            //pars["type"] = handle_name;
                            //pars["lotId"] = "";
                            //pars["Formname"] = "";
                            //pars["Stepname"] = "";
                            //try
                            //{
                            //    stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                            //    FileLog.WriteLog("停机返回值:" + stop);
                            //}
                            //catch (Exception ex)
                            //{
                            //    FileLog.WriteLog("停机失败反馈：" + ex.Message.ToString());
                            //}

                            ////锁机成功插入数据，方便前台解锁
                            //try
                            //{
                            //    FileLog.WriteLog("---触发插表---");
                            //    ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE,REMARK) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{eqpid}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate,'{stop}')", conn);
                            //    int res = ocmd.ExecuteNonQuery();
                            //    FileLog.WriteLog("插库反馈：" + res);
                            //}
                            //catch (Exception ex)
                            //{
                            //    FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                            //}

                            try
                            {
                                MailSender mail = new MailSender("");
                                string[] recevicerList = recevicer.Split(';');
                                mail.AddTo(recevicerList);

                                mailTitle = value.MAILTITLE;
                                mailBody = "MES的测试程序名跟MIR里的测试程序名不一致，MES："+ retXml + ",TDAS："+ value.JOBNAM + "";

                                try
                                {
                                    FileLog.WriteLog("使用mailkit发送邮件...");
                                    mailkit.Send(recevicerList, mailTitle, mailBody, mailtype);

                                }
                                catch (Exception ex)
                                {
                                    FileLog.WriteLog("错误：" + ex.Message.ToString());
                                    FileLog.WriteLog("使用mailkit发送邮件发送失败，换用C#SMTP发送邮件");
                                    try
                                    {
                                        mail.Send(mailTitle, mailBody);
                                    }
                                    catch (Exception e)
                                    {
                                        FileLog.WriteLog("错误：" + e.Message.ToString());
                                        FileLog.WriteLog("使用C#SMTP发送邮件失败");
                                    }
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
                        FileLog.WriteLog("获取Mes程序名失败：" + ex.Message.ToString());
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

        public string Webkey()
        {
            string value = "JCET" + DateTime.Now.ToString("yyyy-MM-dd");
            if (value == null || value == "")
            {
                return "";
            }
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            return Convert.ToBase64String(bytes);
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
            public string MAILTITLE { get; set; }

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
        }

        public class DIFF_RESULT
        {
            public string STDFID { get; set; }
            public string EQPNAME { get; set; }
            public string JOBNAM { get; set; }
            public string SBLOTID { get; set; }
            public string EQPTID { get; set; }
            public string MAILTITLE { get; set; }
        }
    }
}


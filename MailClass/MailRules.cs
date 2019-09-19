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

namespace MailConsole
{
    public class MailRules
    {
        MailHelper mailHelper = new MailHelper();
        DatabaseManager dmgr = new DatabaseManager("oracle");
        DateTime sendTime = DateTime.MinValue;

        public void prrRulseMailAlert(string prr_result)
        {
            try
            {
                using (OracleConnection conn = new OracleConnection(dmgr.ConnectionString))
                {
                    conn.Open();
                    OracleCommand ocmd;
                    PRR_RESULT value = JToken.Parse(prr_result).ToObject<PRR_RESULT>();
                    FileLog.WriteLog("prr_result:" + prr_result);
                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";

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

                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);
                    if (value.ISSTOP.ToString() == "1")
                    {
                        FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + value.EQPTID);
                        Hashtable pars = new Hashtable();
                        pars["key"] = Webkey();
                        pars["userId"] = "shuxi_newserver";
                        pars["eqptId"] = value.EQPTID;
                        pars["type"] = handle_name;
                        pars["lotId"] = "";
                        pars["Formname"] = "";
                        pars["Stepname"] = "";
                        FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + value.EQPTID);
                        stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                        FileLog.WriteLog("PRR停机返回值:" + stop);

                        //锁机成功插入数据，方便前台解锁
                        if (stop.Contains("Y"))
                        {
                            try
                            {
                                FileLog.WriteLog("---触发插表---");
                                ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{value.EQPTID}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate)", conn);
                                int res = ocmd.ExecuteNonQuery();
                                FileLog.WriteLog("插库反馈：" + res);
                            }
                            catch (Exception ex)
                            {
                                FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                            }
                        }
                    }

                    try
                    {
                        MailSender mail = new MailSender();
                        recevicer = "jiangtao@cj-elec.com;" + recevicer;
                        string[] recevicerList = recevicer.Split(';');
                        mail.AddTo(recevicerList);
                        switch (type)
                        {
                            case "BINCOUNTTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
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

                                mail.Send(mailTitle, mailBody);
                                break;

                            case "SITETOSITEYIELDTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "CONSECUTIVEBINTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "PARAMETRICTESTSTATISTICTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

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

                    FileLog.WriteLog("ptr_result:" + ptr_result);
                    string mailTitle = "";
                    string mailBody = "";
                    string type = "";
                    string recevicer = "";
                    string stop = "";
                    string unlockrole = "";
                    string unlockbm = "";
                    string handle_name = "";

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

                    FileLog.WriteLog("ISSTOP：" + value.ISSTOP + ",EQPNAME：" + value.EQPNAME);
                    if (value.ISSTOP.ToString() == "1")
                    {
                        //ocmd = new OracleCommand(@"select * from STOPHANDLER_EQPNAME where EQPNAME='" + value.EQPNAME.ToString() + "'", conn);

                        //OracleDataAdapter oda_StopHandler = new OracleDataAdapter(ocmd);
                        //DataSet ds_StopHandler = new DataSet();
                        //oda_StopHandler.Fill(ds_StopHandler);
                        //DataTable dt_StopHandler = ds_StopHandler.Tables[0];
                        //if (dt_StopHandler.Rows.Count > 0)
                        //{
                        FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + value.EQPTID);
                        Hashtable pars = new Hashtable();
                        pars["key"] = Webkey();
                        pars["userId"] = "shuxi_newserver";
                        pars["eqptId"] = value.EQPTID;
                        pars["type"] = handle_name;
                        pars["lotId"] = "";
                        pars["Formname"] = "";
                        pars["Stepname"] = "";
                        FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + value.EQPTID);
                        stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                        FileLog.WriteLog("PRR停机返回值:" + stop);

                        //锁机成功插入数据，方便前台解锁
                        if (stop.Contains("Y"))
                        {
                            try
                            {
                                FileLog.WriteLog("---触发插表---");
                                ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE) values ('{value.EQPNAME}','{Webkey()}','shuxi_newserver','{value.EQPTID}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate)", conn);
                                int res = ocmd.ExecuteNonQuery();
                                FileLog.WriteLog("插库反馈：" + res);
                            }
                            catch (Exception ex)
                            {
                                FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                            }
                        }

                    }

                    try
                    {
                        MailSender mail = new MailSender();
                        recevicer = "jiangtao@cj-elec.com;" + recevicer;
                        string[] recevicerList = recevicer.Split(';');
                        mail.AddTo(recevicerList);
                        switch (type)
                        {
                            case "BINCOUNTTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();
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

                                mail.Send(mailTitle, mailBody);
                                break;

                            case "SITETOSITEYIELDTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "CONSECUTIVEBINTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "PARAMETRICTESTSTATISTICTRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

                                break;

                            case "SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER":
                                ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.EQPNAME + @"'||'_'||'" + value.DATETIME + @"'
from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                                mailTitle = ocmd.ExecuteScalar()?.ToString();

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

                                mail.Send(mailTitle, mailBody);

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
                FileLog.WriteLog("ecid:" + result);
                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string type = "";
                string unlockrole = "";
                string unlockbm = "";
                string handle_name = "";
                var value = JToken.Parse(result).ToObject<dynamic>();

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

                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + value.EQPTID + "；EQPNAME：" + value.NODENAM);
                Hashtable pars = new Hashtable();
                pars["key"] = Webkey();
                pars["userId"] = "shuxi_newserver";
                pars["eqptId"] = value.EQPTID;
                pars["type"] = handle_name;
                pars["lotId"] = "";
                pars["Formname"] = "";
                pars["Stepname"] = "";
                string stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                FileLog.WriteLog("ECID停机返回值:" + stop);

                //锁机成功插入数据，方便前台解锁
                if (stop.Contains("Y"))
                {
                    try
                    {
                        FileLog.WriteLog("---触发插表---");
                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{value.EQPTID}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate)", conn);
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

                    MailSender mail = new MailSender();
                    recevicer = "jiangtao@cj-elec.com;" + recevicer;
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //                    ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from v_eq_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //                    mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECID_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.NODENAM + "_" + value.DATETIME;

                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECID|JCET_ECID_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.RESULT + "<br/>";
                    mailBody += value.REFS;
                    mail.Send(mailTitle, mailBody);

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
                FileLog.WriteLog("ecidwafer:" + result);
                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string unlockrole = "";
                string type = "";
                string unlockbm = "";
                string handle_name = "";
                var value = JToken.Parse(result).ToObject<dynamic>();

                if (value.STDFID.ToString() == "15161")
                {
                    return;
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

                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + value.EQPTID);
                Hashtable pars = new Hashtable();
                pars["key"] = Webkey();
                pars["userId"] = "shuxi_newserver";
                pars["eqptId"] = value.EQPTID;
                pars["type"] = handle_name;
                pars["lotId"] = "";
                pars["Formname"] = "";
                pars["Stepname"] = "";
                FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + value.EQPTID);
                string stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                FileLog.WriteLog("ECID停机返回值:" + stop);

                //锁机成功插入数据，方便前台解锁
                if (stop.Contains("Y"))
                {
                    try
                    {
                        FileLog.WriteLog("---触发插表---");
                        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{value.EQPTID}','{handle_name}','{unlockrole}','{unlockbm}','0',sysdate)", conn);
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

                    MailSender mail = new MailSender();
                    recevicer = "jiangtao@cj-elec.com;" + recevicer;
                    //recevicer = "kai.guo@shu-xi.com;zhengshuang.ding@shu-xi.com;jun.lai@cj-elec.com;jiangtao@cj-elec.com";
                    //recevicer = "zhengshuang.ding@shu-xi.com";
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from sys_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECIDWAFER_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.NODENAM + "_" + value.DATETIME;


                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECIDWAFER|JCET_ECIDWAFER_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.REFS + "<br/>";
                    mailBody += value.RESULT;
                    mail.Send(mailTitle, mailBody);

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
                FileLog.WriteLog("ecidwaferakj:" + result);
                string mailTitle = "";
                string mailBody = "";
                string recevicer = "";
                string unlockrole = "";
                string type = "";
                string unlockbm = "";
                var value = JToken.Parse(result).ToObject<dynamic>();

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
                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }

                FileLog.WriteLog("开始触发停机，key：" + Webkey() + "；eqptId：" + value.EQPTID);
                //Hashtable pars = new Hashtable();
                //pars["key"] = Webkey();
                //pars["userId"] = "shuxi_newserver";
                //pars["eqptId"] = value.EQPTID;
                //pars["type"] = type;
                //pars["lotId"] = "";
                //pars["Formname"] = "";
                //pars["Stepname"] = "";
                //FileLog.WriteLog("key:" + Webkey() + ",eqptid:" + value.EQPTID);
                //string stop = WebSvcHelper.QueryGetWebService("http://172.17.255.158:3344/mestocim/Service1.asmx/lockEqptByTypeWithkey", pars);
                //FileLog.WriteLog("ECID停机返回值:" + stop);

                ////锁机成功插入数据，方便前台解锁
                //if (stop.Contains("Y"))
                //{
                //    try
                //    {
                //        FileLog.WriteLog("---触发插表---");
                //        ocmd = new OracleCommand($"insert into UNLOCK_EQPT(EQPNAME,WEBKEY,USERID,EQPTID,TYPE,ULOCKROLE,UNLOCKBM,STATUS,CREATE_DATE) values ('{value.NODENAM}','{Webkey()}','shuxi_newserver','{value.EQPTID}','{type}','{unlockrole}','{unlockbm}','0',sysdate)", conn);
                //        int res = ocmd.ExecuteNonQuery();
                //        FileLog.WriteLog("插库反馈：" + res);
                //    }
                //    catch (Exception ex)
                //    {
                //        FileLog.WriteLog("插库反馈：" + ex.Message.ToString());
                //    }
                //}

                try
                {

                    MailSender mail = new MailSender();
                    recevicer = "jiangtao@cj-elec.com;" + recevicer;
                    string[] recevicerList = recevicer.Split(';');
                    mail.AddTo(recevicerList);

                    //                    ocmd = new OracleCommand(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||'" + value.NODENAM + @"'||'_'||'" + value.DATETIME + @"'
                    //from (select * from v_eq_rules_testrun where guid='" + value.GUID + "') t1,(select * from mir where stdfid='" + value.STDFID + "') t2 ", conn);
                    //                    mailTitle = ocmd.ExecuteScalar()?.ToString();
                    mailTitle = "Pause Production_ECIDWAFER-AKJ_" + value.PARTTYP + "_" + value.LOTID + "_" + value.TESTCOD + "_" + value.SBLOTID + "_" + value.NODENAM + "_" + value.DATETIME;


                    mailBody = value.DATETIME + ",Pause Production" + "<br/>";
                    mailBody += "[ECIDWAFER-AKJ|JCET_ECIDWAFER-AKJ_" + value.PARTTYP + "]" + "<br/>";
                    mailBody += value.REFS + "<br/>";
                    mailBody += value.RESULT;
                    mail.Send(mailTitle, mailBody);

                }
                catch (Exception ex)
                {
                    FileLog.WriteLog(ex.Message + ex.StackTrace);
                }
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

        }
    }
}


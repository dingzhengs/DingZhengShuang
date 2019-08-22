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

namespace MailConsole
{
    public class MailRules
    {
        MailHelper mailHelper = new MailHelper();
        SqlHelper sqlHelper = new SqlHelper();
        DateTime sendTime = DateTime.MinValue;



        public void prrRulseMailAlert(string prr_guid)
        {
            string mailTitle = "";
            string mailBody = "";
            string type = "";

            try
            {
                type = sqlHelper.ExecuteScalar(@"select t1.type from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'").ToString();
            }
            catch (Exception)
            {

            }
            MailSender mail = new MailSender();
            List<string> recevicerList = new List<string>();
            recevicerList.Add("zhengshuang.ding@shu-xi.com");
            mail.AddTo(recevicerList.ToArray());
            switch (type)
            {
                case "BINCOUNTTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + prr_guid + "'").ToString();

                    DataSet ds_BINCOUNTTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'");

                    mailBody = ds_BINCOUNTTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_BINCOUNTTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "SITETOSITEYIELD":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + prr_guid + "'").ToString();

                    DataSet ds_SITETOSITEYIELD = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,t.remark,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||
(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||'%',
'MAX_site='||substr(t.remark,instr(t.remark,'=',1,3)+1,instr(t.remark,',',1,3)-instr(t.remark,'=',1,3)-1)||
' MAX_value='||substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1)||'%'||
' MIN_site='||substr(t.remark,instr(t.remark,'=',1,4)+1,instr(t.remark,',',1,4)-instr(t.remark,'=',1,4)-1)||
' MIN_value='||substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1)||'%'||
' GAP='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||'%'
 from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'");

                    mailBody = ds_SITETOSITEYIELD.Tables[0].Rows[0][0].ToString() + "," + ds_SITETOSITEYIELD.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "CONSECUTIVEBINTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + prr_guid + "'").ToString();

                    DataSet ds_CONSECUTIVEBINTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'");

                    mailBody = ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "PARAMETRICTESTSTATISTICTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'").ToString();

                    DataSet ds_PARAMETRICTESTSTATISTICTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + prr_guid + "'");

                    mailBody = ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + prr_guid + "'").ToString();

                    DataSet ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||t.sitenum||' value='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||,
'MAX_site='||substr(t.remark,instr(t.remark,'=',1,3)+1,instr(t.remark,',',1,3)-instr(t.remark,'=',1,3)-1)||
' MAX_value='||substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1)||
' MIN_site='||substr(t.remark,instr(t.remark,'=',1,4)+1,instr(t.remark,',',1,4)-instr(t.remark,'=',1,4)-1)||
' MIN_value='||substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1)||
' GAP='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + prr_guid + "'");

                    mailBody = ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;
            }
        }

        public void ptrRulseMailAlert(string ptr_guid)
        {
            string mailTitle = "";
            string mailBody = "";
            string type = "";

            try
            {
                type = sqlHelper.ExecuteScalar(@"select t1.type from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'").ToString();
            }
            catch (Exception)
            {

            }
            MailSender mail = new MailSender();
            List<string> recevicerList = new List<string>();
            recevicerList.Add("zhengshuang.ding@shu-xi.com");
            mail.AddTo(recevicerList.ToArray());
            switch (type)
            {
                case "BINCOUNTTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + ptr_guid + "'").ToString();

                    DataSet ds_BINCOUNTTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'");

                    mailBody = ds_BINCOUNTTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_BINCOUNTTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_BINCOUNTTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "SITETOSITEYIELD":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + ptr_guid + "'").ToString();

                    DataSet ds_SITETOSITEYIELD = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,t.remark,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||
(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||'%',
'MAX_site='||substr(t.remark,instr(t.remark,'=',1,3)+1,instr(t.remark,',',1,3)-instr(t.remark,'=',1,3)-1)||
' MAX_value='||substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1)||'%'||
' MIN_site='||substr(t.remark,instr(t.remark,'=',1,4)+1,instr(t.remark,',',1,4)-instr(t.remark,'=',1,4)-1)||
' MIN_value='||substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1)||'%'||
' GAP='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||'%'
 from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'");

                    mailBody = ds_SITETOSITEYIELD.Tables[0].Rows[0][0].ToString() + "," + ds_SITETOSITEYIELD.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEYIELD.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "CONSECUTIVEBINTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + ptr_guid + "'").ToString();

                    DataSet ds_CONSECUTIVEBINTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit=null HighLimit='||case t1.counttype when '0' then t1.count when '1' then t1.count||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
from sys_rules_show t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'");

                    mailBody = ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_CONSECUTIVEBINTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "PARAMETRICTESTSTATISTICTRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'").ToString();

                    DataSet ds_PARAMETRICTESTSTATISTICTRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||t.sitenum||' value='||t.remark,
'MAX_site=null MAX_value=null MIN_site=null MIN_value=null GAP=null'
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + ptr_guid + "'");

                    mailBody = ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_PARAMETRICTESTSTATISTICTRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;

                case "SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER":
                    mailTitle = sqlHelper.ExecuteScalar(@"select t1.action||'_'||t1.type||'_'||t1.product||'_'||t2.lotid||'_'||t2.testcod||'_'||t2.sblotid||'_'||t.eqname||'_'||to_char(t.rules_time,'yyyyMMddHHmmss')
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid where t.guid='" + ptr_guid + "'").ToString();

                    DataSet ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER = sqlHelper.ExecuteDataSet(@"select t.rules_time,t1.action,
'['||t1.type||'/'||t1.name||']:',
'LowLimit='||case t1.minstatus when '0' then t1.minvalue when '1' then t1.minvalue||'%' else 'null' end  ||
' HighLimit='||case t1.maxstatus when '0' then t1.maxvalue when '1' then t1.maxvalue||'%' else 'null' end ||
' site='||t.sitenum||' value='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))||,
'MAX_site='||substr(t.remark,instr(t.remark,'=',1,3)+1,instr(t.remark,',',1,3)-instr(t.remark,'=',1,3)-1)||
' MAX_value='||substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1)||
' MIN_site='||substr(t.remark,instr(t.remark,'=',1,4)+1,instr(t.remark,',',1,4)-instr(t.remark,'=',1,4)-1)||
' MIN_value='||substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1)||
' GAP='||(to_number(NVL(substr(t.remark,instr(t.remark,'=',1,1)+1,instr(t.remark,',',1,1)-instr(t.remark,'=',1,1)-1),0))-
to_number(NVL(substr(t.remark,instr(t.remark,'=',1,2)+1,instr(t.remark,',',1,2)-instr(t.remark,'=',1,2)-1),0)))
 from sys_rules_show_ptr t
left join sys_rules_testrun t1 on t.rules_guid=t1.guid
left join mir t2 on t.stdfid=t2.stdfid  where t.guid='" + ptr_guid + "'");

                    mailBody = ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][0].ToString() + "," + ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][1].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][2].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][3].ToString() + "<br/>";
                    mailBody += ds_SITETOSITEPARAMETRICTESTSTATISTICDELTATRIGGER.Tables[0].Rows[0][4].ToString();

                    mail.Send(mailTitle, mailBody);

                    break;
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
    }
}


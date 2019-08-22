using System;
using System.Data;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data.Oracle;
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;
using System.Collections.Generic;

namespace MailClass
{
    public class SqlHelper
    {

        private DbTransaction _trans = null;
        private Database dbBase;
        public DbConnection conn = null;
        private bool _isDisposed;
        private bool _isTrans = false;

        /// <summary>
        /// 构造函数，创建数据库访问实例支持
        /// </summary>
        /// <param name="isCreateConnection">是否创建connection连接，默认否</param>
        /// <param name="dbName">数据库连接名称，不设置默认成web配置里的defaultConnection</param>
        public SqlHelper(bool isCreateConnection = false, string dbName = "")
        {
            dbBase = string.IsNullOrEmpty(dbName) ? DatabaseFactory.CreateDatabase() : DatabaseFactory.CreateDatabase(dbName);
            if (dbBase is SqlDatabase)
            {
                DbType = DataBaseType.SqlDataBase;
            }
            if (dbBase is OracleDatabase)
            {
                DbType = DataBaseType.OracleDataBase;
            }
            if (isCreateConnection)
            {
                conn = dbBase.CreateConnection();
            }
        }

        /// <summary>
        /// 开始启动事务，在遇到Commit或者Rollback之前的所有此实例执行的SQL语句都会绑定在同一个事务里
        /// </summary>
        public void BeginTransaction()
        {
            if (conn == null)
            {
                conn = dbBase.CreateConnection();
            }
            conn.Open();
            _isTrans = true;
            _trans = conn.BeginTransaction();
        }

        /// <summary>
        /// 返回数据库类型，是SQL Server还是Oracle
        /// </summary>
        public DataBaseType DbType { get; private set; }

        private int _timeOut = 0;
        /// <summary>
        /// 设置超时时间(单位秒)，默认30秒
        /// </summary>
        public int TimeOut
        {
            get
            {
                if (_timeOut == 0)
                {
                    if (System.Configuration.ConfigurationManager.AppSettings["TimeOut"] != null)
                    {
                        try
                        {
                            _timeOut = int.Parse(System.Configuration.ConfigurationManager.AppSettings["TimeOut"].ToString());
                        }
                        catch
                        {
                            _timeOut = 300;
                        }
                    }
                }
                return _timeOut;
            }
            set { _timeOut = value; }
        }

        /// <summary>
        /// 获取链接字符串
        /// </summary>
        /// <returns></returns>
        public string getConnection()
        {
            return dbBase.ConnectionString.ToString();
        }

        //创建Command信息
        private DbCommand CreateCommand(string commandText, CommandType commandtype, dbParameter[] dbParams)
        {
            DbCommand dbCommand = null;


            if (dbParams != null && dbParams.Length > 0)
            {
                string paramName = string.Empty;
                foreach (dbParameter param in dbParams)
                {
                    paramName = param.Name.Replace("@", "").Replace(":", "");
                    if (this.DbType == DataBaseType.OracleDataBase)
                    {
                        commandText = commandText.Replace("@" + paramName, ":" + paramName);
                    }
                    else if (this.DbType == DataBaseType.SqlDataBase)
                    {
                        commandText = commandText.Replace(":" + paramName, "@" + paramName);
                    }
                }
            }

            switch (commandtype)
            {
                case CommandType.Text:
                    dbCommand = dbBase.GetSqlStringCommand(commandText);
                    break;
                case CommandType.StoredProcedure:
                    dbCommand = dbBase.GetStoredProcCommand(commandText);
                    break;
                case CommandType.TableDirect:
                    dbCommand = dbBase.GetSqlStringCommand(string.Format("select * from {0}", commandText));
                    break;
            }

            if (dbParams != null)
            {
                foreach (dbParameter param in dbParams)
                {
                    if (param.Driection == ParamDirection.In)
                    {
                        dbBase.AddInParameter(dbCommand, param.Name, param.Type, param.Value);
                    }
                    else
                    {
                        dbBase.AddOutParameter(dbCommand, param.Name, param.Type, 2048);
                    }
                }
            }
            dbCommand.CommandTimeout = TimeOut;
            return dbCommand;
        }

        /// <summary>
        /// 返回一个DataSet结果集
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回一个结果集</returns>
        public DataSet ExecuteDataSet(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            DbCommand Command = CreateCommand(strSQL, commandtype, Param);
            if (_isTrans && _trans != null)
            {
                try
                {
                    return dbBase.ExecuteDataSet(Command, _trans);
                }
                catch (Exception ex)
                {
                    Rollback();
                    throw ex;
                }
            }
            else
            {
                return dbBase.ExecuteDataSet(Command);
            }
        }

        /// <summary>
        /// 返回结果集中的第一个DataTable
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回一个DataTable</returns>
        public DataTable ExecuteDataTable(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            DataSet ds = ExecuteDataSet(strSQL, Param, commandtype);
            if (ds == null || ds.Tables.Count == 0)
            {
                return null;
            }
            else
            {
                return ds.Tables[0];
            }
        }

        /// <summary>
        /// 返回第一行第一列单元格的值
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回一个单例</returns>
        public object ExecuteScalar(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            DbCommand Command = CreateCommand(strSQL, commandtype, Param);
            if (_isTrans && _trans != null)
            {
                try
                {
                    return dbBase.ExecuteScalar(Command, _trans);
                }
                catch (Exception ex)
                {
                    Rollback();
                    throw ex;
                }
            }
            else
            {
                return dbBase.ExecuteScalar(Command);
            }
        }

        /// <summary>
        /// 一个整数值
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回一个整数</returns>
        public int ExecuteCount(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            return (int)ExecuteScalar(strSQL, Param, commandtype);
        }

        /// <summary>
        /// 执行增删改SQL
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回小于-1 表示执行失败</returns>
        public int ExecuteNonQuery(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            DbCommand Command = CreateCommand(strSQL, commandtype, Param);
            //Command.Prepare();
            if (_isTrans && _trans != null)
            {
                try
                {
                    return dbBase.ExecuteNonQuery(Command, _trans);
                }
                catch (Exception ex)
                {
                    Rollback();
                    throw ex;
                }
            }
            else
            {
                return dbBase.ExecuteNonQuery(Command);
            }
        }

        /// <summary>
        /// ExecuteReader
        /// </summary>
        /// <param name="strSQL">查询sql语句</param>
        /// <param name="Param">参数，无则null</param>
        /// <param name="commandtype">Command类型 不写则默认CommandType.Text</param>
        /// <returns>返回一IDataReader</returns>
        public IDataReader ExecuteReader(string strSQL, dbParameter[] Param = null, CommandType commandtype = CommandType.Text)
        {
            DbCommand Command = CreateCommand(strSQL, commandtype, Param);
            if (_isTrans && _trans != null)
            {
                try
                {
                    return dbBase.ExecuteReader(Command, _trans);
                }
                catch (Exception ex)
                {
                    Rollback();
                    throw ex;
                }
            }
            else
            {
                try
                {
                    return dbBase.ExecuteReader(Command);
                }
                catch
                {

                    throw;
                }
            }
        }

        /// <summary>
        /// 提交所有事务
        /// </summary>
        public void Commit()
        {
            if (conn != null && conn.State == ConnectionState.Open)
            {
                if (_trans != null)
                {
                    _trans.Commit();
                    _trans.Dispose();
                    _trans = null;
                }
                conn.Close();
                _isTrans = false;
            }
        }

        /// <summary>
        /// 回滚所有事务
        /// </summary>
        public void Rollback()
        {
            if (conn != null && conn.State == ConnectionState.Open)
            {
                if (_trans != null)
                {
                    _trans.Rollback();
                    _trans.Dispose();
                    _trans = null;
                }
                conn.Close();
                _isTrans = false;
            }
        }

        /// <summary>
        /// 执行存储过程，并获取输出返回值
        /// </summary>
        /// <param name="strSQL">存储过程名称</param>
        /// <param name="Param">参数</param>
        /// <returns>根据out参数顺序返回值</returns>
        public object[] ExecuteProcedureWithOutParam(string strSQL, dbParameter[] Param)
        {
            DbCommand Command = CreateCommand(strSQL, CommandType.StoredProcedure, Param);
            List<Object> outList = new List<Object>();
            if (_isTrans && _trans != null)
            {
                try
                {
                    dbBase.ExecuteNonQuery(Command, _trans);

                    foreach (var item in Param)
                    {
                        if (item.Driection == ParamDirection.Out)
                        {
                            outList.Add(dbBase.GetParameterValue(Command, item.Name));
                        }
                    }
                    return outList.ToArray();
                }
                catch (Exception ex)
                {
                    Rollback();
                    throw ex;
                }
            }
            else
            {
                dbBase.ExecuteNonQuery(Command);
                foreach (var item in Param)
                {
                    if (item.Driection == ParamDirection.Out)
                    {
                        outList.Add(dbBase.GetParameterValue(Command, item.Name));
                    }
                }
                return outList.ToArray();
            }
        }
    }

    public class dbParameter
    {
        public string Name { get; set; }
        public DbType Type { get; set; }
        public object Value { get; set; }
        public ParamDirection Driection { get; set; }
        public dbParameter(string Name, object Value, DbType Type = DbType.String
                           , ParamDirection Driection = ParamDirection.In)
        {
            this.Name = Name;
            this.Value = Value;
            this.Type = Type;
            this.Driection = Driection;
        }
    }

    public enum DataBaseType
    {
        SqlDataBase = 0,

        OracleDataBase = 1
    }

    public enum ParamDirection
    {
        In,
        Out
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;

namespace MailClass
{
    public static class FileLog
    {

        private static readonly object _lockMethod = new object();
        private static DirectoryInfo WritePath = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + @"\Logs");
        public static void WriteLog(string Info)
        {
            lock (_lockMethod)
            {
                try
                {
                    string LogFormat = string.Format(">>>[{0}] {1}{2}", DateTime.Now.ToString("HH:mm:ss"), Info, Environment.NewLine+Environment.NewLine);
                    LogFormat = (Info == "" ? "" : LogFormat);
                    if (!WritePath.Exists)
                    {
                        WritePath.Create();
                    }
                    FileStream stream = new FileStream(string.Format(@"{0}\{1}.log", WritePath.FullName, DateTime.Now.ToString("yyyy-MM-dd")), FileMode.Append, FileAccess.Write);
                    StreamWriter writer = new StreamWriter(stream, Encoding.Default);
                    writer.Write(LogFormat);
                    writer.Close();
                    stream.Close();
                }
                catch
                {

                }
            }
        }

    }


}

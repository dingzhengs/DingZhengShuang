using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TDASTest
{
    public class TcpClient
    {
        Thread threadclient = null;
        Socket client;
        IPEndPoint point;
        public TcpClient(string ip, int port)
        {
            point = new IPEndPoint(IPAddress.Parse(ip), port);
        }

        public string Submit(string msg)
        {
            client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            client.Connect(point);

            client.Send(Encoding.UTF8.GetBytes(msg));

            byte[] buffer = new byte[1024];
            client.Receive(buffer);
            string result = Encoding.UTF8.GetString(buffer);
            client.Close();

            return result;
        }

        void recv()
        {
            //持续监听服务端发来的消息 
            while (true)
            {
                try
                {
                    Socket cc = client.Accept();
                    byte[] buffer = new byte[1024 * 1024];
                    //接收数据到缓冲区
                    client.Receive(buffer);

                }
                catch (Exception ex)
                {
                }
                finally
                {
                    client.Close();
                    threadclient.Abort();
                }
            }
        }

    }
}

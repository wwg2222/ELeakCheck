using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OPCAutomation;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string strTopic;
        OPCServer MyServer;
        OPCGroups MyGroups;
        OPCGroup MyGroup;
        OPCGroup MyGroup2;
        OPCItems MyItems;
        OPCItems MyItems2;
        OPCItem[] MyItem;
        OPCItem[] MyItem2;
        string strHostIP = "";
        string strHostName = "";
        bool opc_connected = false;
        int itmHandleClient = 0;
        int itmHandleServer = 0;

        private void GetLocalServer()
        {
            //获取本地计算机IP,计算机名称
            strHostName = System.Net.Dns.GetHostName();
            //或者通过局域网内计算机名称


            //获取本地计算机上的OPCServerName
            try
            {
                MyServer = new OPCServer();
                object serverList = MyServer.GetOPCServers(strHostName);

                foreach (string server in (Array)serverList)
                {
                    //cmbServerName.Items.Add(turn);
                    Console.WriteLine("本地OPC服务器：{0}", server);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("枚举本地OPC服务器出错：{0}", err.Message);
            }
        }

        private bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                MyServer.Connect(remoteServerName, remoteServerIP);//连接本地服务器：服务器名+主机名或IP

                if (MyServer.ServerState == (int)OPCServerState.OPCRunning)
                {
                    //                    string str1;

                    label6.Text = "已连接到 RSLinx OPC Server";
                //  timer1.Start();
                    // Console.WriteLine("已连接到：{0}", MyServer.ServerName);
                }
                else
                {
                    //这里你可以根据返回的状态来自定义显示信息，请查看自动化接口API文档
                    Console.WriteLine("状态：{0}", MyServer.ServerState.ToString());
                    timer1.Stop();
                }
                MyServer.ServerShutDown += ServerShutDown;//服务器断开事件
            }
            catch (Exception err)
            {
                Console.WriteLine("连接远程服务器出现错误：{0}" + err.Message);
                return false;
            }
            return true;
        }

        public void ServerShutDown(string Reason)//服务器先行断开
        {
            Console.WriteLine("服务器已经先行断开！");
        }

        private bool CreateGroup()
        {
            try
            {
                MyGroups = MyServer.OPCGroups;
                MyGroup = MyServer.OPCGroups.Add("测试");//添加组
                MyGroup2 = MyServer.OPCGroups.Add("测试2");

                //MyGroup2 = MyGroups.Add("测试2");
                //以下设置组属性
                {
                    MyServer.OPCGroups.DefaultGroupIsActive = true;//激活组。
                    MyServer.OPCGroups.DefaultGroupDeadband = 0;// 死区值，设为0时，服务器端该组内任何数据变化都通知组。
                    MyServer.OPCGroups.DefaultGroupUpdateRate = 200;//默认组群的刷新频率为200ms
                    MyGroup.UpdateRate = 100;//刷新频率为1秒。
                    MyGroup.IsSubscribed = true;//禁用订阅功能，即可以异步，默认false

                }

                MyGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(GroupDataChange);
                MyGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete);
                MyGroup.AsyncReadComplete += new DIOPCGroupEvent_AsyncReadCompleteEventHandler(GroupAsyncReadComplete);
                MyGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete);

                {
                    MyServer.OPCGroups.DefaultGroupIsActive = true;//激活组。
                    MyServer.OPCGroups.DefaultGroupDeadband = 0;// 死区值，设为0时，服务器端该组内任何数据变化都通知组。
                    MyServer.OPCGroups.DefaultGroupUpdateRate = 200;//默认组群的刷新频率为200ms
                    MyGroup2.UpdateRate = 100;//刷新频率为1秒。
                    MyGroup2.IsSubscribed = true;//使用订阅功能，即可以异步，默认false

                }

                MyGroup2.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(GroupDataChange);
                MyGroup2.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete);
                MyGroup2.AsyncReadComplete += new DIOPCGroupEvent_AsyncReadCompleteEventHandler(GroupAsyncReadComplete);
                MyGroup2.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(GroupAsyncWriteComplete);

                AddGroupItems();//设置组内items         
            }
            catch (Exception err)
            {
                Console.WriteLine("创建组出现错误：{0}", err.Message);
                return false;
            }
            return true;
        }

        private void AddGroupItems()//添加数据项
        {
            //itmHandleServer;
            MyItems = MyGroup.OPCItems;
            MyItems2 = MyGroup2.OPCItems;

            //添加item
            MyItem[0] = MyItems.AddItem("[" + strTopic + "]int1", 0);//[opc1]int1", 0);//int
            MyItem[1] = MyItems.AddItem("[" + strTopic + "]dint1", 1);//dint
            MyItem[2] = MyItems.AddItem("[" + strTopic + "]dint2", 1);//dint


            MyItem2[0] = MyItems2.AddItem("[" + strTopic + "]int2", 0);//int
            MyItem2[1] = MyItems2.AddItem("[" + strTopic + "]sint1", 1);//byte

        }

        void GroupDataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.WriteLine("++++++++++++++++DataChanged+++++++++++++++++++++++");

        }
        void GroupDataChange2(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            Console.WriteLine("----------------------DataChanged2------------------");

        }
        void GroupAsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {

            label3.Text = "成功写";
            /*for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("Tran：{0}   ClientHandles：{1}   Error：{2}", TransactionID.ToString(), ClientHandles.GetValue(i).ToString(), Errors.GetValue(i).ToString());
            }*/
        }
        void GroupAsyncReadComplete(int TransactionID, int NumItems, ref System.Array ClientHandles, ref System.Array ItemValues, ref System.Array Qualities, ref System.Array TimeStamps, ref System.Array Errors)
        {

            textBox4.Text = Convert.ToString(ItemValues.GetValue(1));
            textBox5.Text = Convert.ToString(ItemValues.GetValue(2));


        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MyItem = new OPCItem[3];
            MyItem2 = new OPCItem[2];
            strTopic = textBox3.Text;
            GetLocalServer();
            //ConnectRemoteServer("192.168.1.35", "KEPware.KEPServerEx.V4");//用IP的局域网
            if (ConnectRemoteServer("", "RSLinx OPC Server"))//本机
            {
                CreateGroup();

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            object ItemValues; object Qualities; object TimeStamps;//同步读的临时变量：值、质量、时间戳
            MyItem[0].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
            int q0 = Convert.ToInt16(ItemValues);//转换后获取item值
            textBox1.Text = q0 + "";
            MyItem[1].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
            int q1 = Convert.ToInt32(ItemValues);//转换后获取item值
            textBox2.Text = q1 + "";
            MyItem[2].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
            int q2 = Convert.ToInt32(ItemValues);//转换后获取item值
            textBox6.Text = q2 + "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int int1, dint1, dint2;
            int1 = Convert.ToInt16(textBox1.Text);
            dint1 = Convert.ToInt32(textBox2.Text);
            dint2 = Convert.ToInt32(textBox6.Text);

            MyItem[0].Write(int1);
            MyItem[1].Write(dint1);
            MyItem[2].Write(dint2);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int[] temp = new int[] { 0, MyItem2[0].ServerHandle, MyItem2[1].ServerHandle };

            Array serverHandles = (Array)temp;

            Array Errors;
            int cancelID;

            MyGroup.AsyncRead(2, ref serverHandles, out Errors, 1, out cancelID);//第一参数为item数量
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int[] temp = new int[] { 0, MyItem2[0].ServerHandle, MyItem2[1].ServerHandle };
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[3];
            int int1, sint1;
            int1 = Convert.ToInt16(textBox4.Text);
            sint1 = Convert.ToSByte(textBox5.Text);

            valueTemp[0] = "";
            valueTemp[1] = int1;
            valueTemp[2] = sint1;

            Array values = (Array)valueTemp;
            Array Errors;
            int cancelID;
            MyGroup.AsyncWrite(2, ref serverHandles, ref values, out Errors, 1, out cancelID);//第一参数为item数量
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
                object ItemValues; object Qualities; object TimeStamps;//同步读的临时变量：值、质量、时间戳
                MyItem[0].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                int q0 = Convert.ToInt16(ItemValues);//转换后获取item值
                textBox1.Text = q0 + "";
                MyItem[1].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                int q1 = Convert.ToInt32(ItemValues);//转换后获取item值
                textBox2.Text = q1 + "";
                MyItem[2].Read(1, out ItemValues, out Qualities, out TimeStamps);//同步读，第一个参数只能为1或2
                int q2 = Convert.ToInt32(ItemValues);//转换后获取item值
                textBox6.Text = q2 + "";
            
            
        }
    }
}

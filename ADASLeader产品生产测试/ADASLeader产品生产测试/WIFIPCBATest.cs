using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop;
using System.Diagnostics;
using System.ComponentModel;
namespace ADASLeader产品生产测试
{
    public partial class WIFIPCBATest : Form
    {
        public WIFIPCBATest(string barNum)
        {
            InitializeComponent();
            InitializeComponent();
            commBuf = new byte[4096];
            BarNum = barNum;
        }
        


        public int PASS_FLAG; //1：测试通过，其他：测试未通过

        public string mFilename;
        private string BarNum;
        public byte[] commBuf;
        public SerialPort comm;
        public enum timeF
        {
            selfTest = 0,

            TotalTimer = 16,
        };
        int[] TimerArry = new int[16];




        public string Cdirectory;
        Microsoft.Office.Interop.Excel.Application xApp;
        Microsoft.Office.Interop.Excel.Workbooks wbs;
        Microsoft.Office.Interop.Excel.Workbook wb;
        Microsoft.Office.Interop.Excel.Worksheets wss;
        Microsoft.Office.Interop.Excel.Worksheet ws;

        public enum column
        {
            BarNum = 1,             //条形码
            work_voltage = 2,
            work_current = 3,
            static_current = 4,
        }

        string[] columnName = { "条形码", "工作电压", "工作电流", "静态电流" };

        public int startColumn = (int)column.BarNum;
        public int endColumn = (int)column.static_current;

        public struct tlv_reply_t
        {
            public UInt32 ReplyType;
            public UInt32 CpuID0;
            public UInt32 CpuID1;
            public UInt32 CpuID2;
            public byte TestResult0;
            public byte TestResult1;
            public byte TestResult2;
            public byte TestResult3;
            public UInt16 Voltage0;     //IN_3.3V
            public UInt16 Voltage1;     //IN_5V
            public UInt16 Voltage2;     //IN_12V
            public UInt16 Voltage3;     //IN_3.3V
            public UInt16 Voltage4;     //IN_2.5V
            public UInt16 Voltage5;     //1.2V
            public UInt16 Voltage6;     //IN_POWER1
            public UInt16 Voltage7;     //IN_POWER1
            public UInt16 Voltage8;     //IN_POWER1
            public UInt16 Current0;     //workcurrent
            public UInt16 Current1;     //staticcurrent
            public byte Mode;           // 0:  none 1 : test   2:  work   3.  sleep 
        };

        public struct test_resault
        {
            public byte TestResult0;
            public byte TestResult1;
            public byte TestResult2;
            public byte TestResult3;


            public int WorkVoltage_OK_NOT;
            public int Static_Curent_OK_NOT;
            public int Work_Curent_OK_NOT;
            public byte Mode;           // 0:  none 1 : test   2:  work   3.  sleep 
        };
        tlv_reply_t tlv_reply;
        tlv_reply_t tlv_reply_min;
        tlv_reply_t tlv_reply_max;

        test_resault Test_Resault;
        public void Anisy_Reply(int rows, tlv_reply_t reply, tlv_reply_t replymax, tlv_reply_t replymin)  //如果reply1中
        {
            //静态电压
            Microsoft.Office.Interop.Excel.Range ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rows, (int)column.static_current];
            //if ((reply.Voltage0 <= replymin.Voltage0) | (reply.Voltage0 >= replymax.Voltage0))
            ran.Interior.ColorIndex = 3;   //红色
            //工作电流
            ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rows, (int)column.work_current];
            if ((reply.Voltage0 < replymin.Voltage0) | (reply.Voltage0 > replymax.Voltage0))
                ran.Interior.ColorIndex = 3;   //红色
            //工作电压
            ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rows, (int)column.work_voltage];
            if ((reply.Voltage0 < replymin.Voltage0) | (reply.Voltage0 > replymax.Voltage0))
                ran.Interior.ColorIndex = 3;   //红色
        }


        public void Compare_Reply(tlv_reply_t reply, tlv_reply_t replymax, tlv_reply_t replymin)
        {
            Test_Resault.WorkVoltage_OK_NOT = (reply.Voltage0 > replymin.Voltage0) ? 1 : 2;
            Test_Resault.Work_Curent_OK_NOT = (reply.Voltage1 > replymin.Voltage1) ? 1 : 2;
            Test_Resault.Static_Curent_OK_NOT = (reply.Voltage2 > replymin.Voltage2) ? 1 : 2;

            if ((Test_Resault.WorkVoltage_OK_NOT == 1) &&
                (Test_Resault.Work_Curent_OK_NOT == 1) &&
                (Test_Resault.Static_Curent_OK_NOT == 1)
                )
                PASS_FLAG = 1;
            else
                PASS_FLAG = 0;
        }
        public struct msg_head_t
        {
            public UInt32 start_bytes;
            public UInt16 msg_len;
            public byte msg_count;
            public byte resp;              // 0 failed, 1 ok
            public UInt16 seq;
            public byte service_type;
            public byte msg_type;
            public UInt32 crc32;
        };

        public static object BytesToStruct(byte[] bytes, Type structObj)
        {
            int size = Marshal.SizeOf(structObj);
            if (size > bytes.Length) return null;

            IntPtr structPtr = Marshal.AllocHGlobal(size);
            //将byte数组拷到分配好的内存空间
            Marshal.Copy(bytes, 0, structPtr, size);
            //将内存空间转换为目标结构体
            object obj = Marshal.PtrToStructure(structPtr, structObj);
            //释放内存空间
            Marshal.FreeHGlobal(structPtr);
            //返回结构体
            return obj;
        }
        public struct ACKINFO
        {
            public int ACK_File_Read;   //
            public int ACK_File_Write;
            public int ACK_CMD_Rest;
            public int ACK_CMD_Switch_Screen;
            public int ACK_CMD_Test_Resualt;
            public int ACK_CMD_Get_Mcu_ID;
            public int ACK_CMD_Start_Test;
            public int ACK_CMD_Warn;
            public int ACK_CMD_BOOT_IN_IAP_MODE;
            public int ACK_CMD_BOOT_IN_SLEEP_MODE;
            public int ACK_CMD_CRYPT;
            public int ACK_CMD_WORK_CURRENT;
            public int ACK_CMD_SLEEP_CURRENT;
            public int ACK_CMD_TEST_MCU_CODE;
            public int ACK_CMD_TEST_MCU_RESTART;
            public int ACK_CMD_BEEP;
            public int ACK_CMD_WIFI_CODE;
        };

        public ACKINFO ACKinfo;


        private void btn_openCOM_Click(object sender, EventArgs e)
        {
            comm = new SerialPort();
            //关闭时点击，则设置好端口，波特率后打开
            comm.PortName = comboBox_select_ComPort.Text.ToString();
            comm.BaudRate = 115200;
            comm.DataBits = 8;
            comm.Parity = Parity.None;
            comm.StopBits = StopBits.One;
            comm.WriteBufferSize = 1024 * 1024 * 5;
            comm.DataReceived += new SerialDataReceivedEventHandler(commDataReceivedHandler);
            try
            {
                comm.Open();
                textBox1.Text += "\r\n打开串口成功!";
                btn_write_bin.Enabled = true;

                //建立串口接收线程
            }
            catch (Exception ex)
            {
                //捕获到异常信息，创建一个新的comm对象，之前的不能用了。
                comm = new SerialPort();
                //现实异常信息给客户。
                textBox1.Text += "\r\n" + ex.Message;
            }


            //分析串口收到的数据，线程
            //Thread GetACKInfo = new Thread(GetACKStatus);
            //GetACKInfo.Start();
        }

        private void WIFIPCBATest_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void WIFIPCBATest_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            Array.Sort(ports);
            comboBox_select_ComPort.Items.AddRange(ports);
            timer1.Start();


            btn_submit.Enabled = false;
            //btn_submit.Enabled = false;
            btn_write_bin.Enabled = false;
            btn_function_test.Enabled = false;
            btn_performance_test.Enabled = false;

            ACKinfo = new ACKINFO();
            ACKinfo.ACK_CMD_BOOT_IN_IAP_MODE = 0;
            ACKinfo.ACK_CMD_BOOT_IN_SLEEP_MODE = 0;
            ACKinfo.ACK_CMD_CRYPT = 0;
            ACKinfo.ACK_CMD_Get_Mcu_ID = 0;
            ACKinfo.ACK_CMD_Rest = 0;
            ACKinfo.ACK_CMD_SLEEP_CURRENT = 0;
            ACKinfo.ACK_CMD_Start_Test = 0;
            ACKinfo.ACK_CMD_Switch_Screen = 0;
            ACKinfo.ACK_CMD_TEST_MCU_CODE = 0;
            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 0;
            ACKinfo.ACK_CMD_Test_Resualt = 0;
            ACKinfo.ACK_CMD_Warn = 0;
            ACKinfo.ACK_CMD_WORK_CURRENT = 0;
            ACKinfo.ACK_File_Read = 0;
            ACKinfo.ACK_File_Write = 0;


            tlv_reply = new tlv_reply_t();
            Test_Resault = new test_resault();

            tlv_reply_t tlv_reply_min;
            tlv_reply_t tlv_reply_max;

            tlv_reply_min.Voltage0 = 3;    //IN_3.3V
            tlv_reply_min.Voltage1 = 4;    //IN_5V
            tlv_reply_min.Voltage2 = 11;    //IN_12V
            tlv_reply_min.Voltage3 = 3;    //IN_3.3V
            tlv_reply_min.Voltage4 = 2;    //IN_2.5V
            tlv_reply_min.Voltage5 = 1;    //IN_1.2V
            tlv_reply_min.Voltage6 = 10;    //IN_POWER1
            tlv_reply_min.Voltage7 = 10;    //IN_POWER2
            tlv_reply_min.Voltage8 = 10;    //IN_POWER3
            tlv_reply_min.Current0 = 10;    //workCurrent
            tlv_reply_min.Current1 = 10;    //staticCurrent

            tlv_reply_max.Voltage0 = 20;
            tlv_reply_max.Voltage1 = 20;
            tlv_reply_max.Voltage2 = 20;
            tlv_reply_max.Voltage3 = 20;
            tlv_reply_max.Voltage4 = 20;
            tlv_reply_max.Voltage5 = 20;
            tlv_reply_max.Voltage6 = 20;
            tlv_reply_max.Voltage7 = 20;
            tlv_reply_max.Voltage8 = 20;
            tlv_reply_max.Current0 = 20;
            tlv_reply_max.Current1 = 20;


            Cdirectory = System.IO.Directory.GetCurrentDirectory();
            xApp = new Microsoft.Office.Interop.Excel.Application();
            xApp.DisplayAlerts = false;

            if (File.Exists(Cdirectory + "\\主板测试结果.xlsx"))
            {
                wbs = xApp.Workbooks;
                wb = wbs.Add(Cdirectory + "\\主板测试结果.xlsx");
                ws = wb.Worksheets["测试结果"];
            }
            else
            {
                wbs = xApp.Workbooks;
                wb = wbs.Add(true);
                ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ws.Name = "测试结果";
                for (int i = startColumn; i <= endColumn; i++)
                {
                    ws.Cells[1, i] = columnName[i - 1];
                }
            }
        }

        private void btn_write_bin_Click(object sender, EventArgs e)
        {
            //开始测试    线程
            Thread startTestThred = new Thread(startTest);
            startTestThred.Start();

            //下载MCU  BIN文件  线程
            Thread DownLoadMCUBINThred = new Thread(downLoadWIFIBIN);
            DownLoadMCUBINThred.Start();

            //超时等待  线程
            Thread WaitDownloadACKThred = new Thread(waitDownloadACK);
            WaitDownloadACKThred.Start();
        }
    }
}

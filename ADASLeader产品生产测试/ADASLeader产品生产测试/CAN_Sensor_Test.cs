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
using System.Text.RegularExpressions;

namespace ADASLeader产品生产测试
{
    public partial class CAN_Sensor_Test : Form
    {
        test_resault Test_Resault;

        public const byte SERVICE_TYPE_FILE = 0x01;
        public const byte SERVICE_TYPE_CMD = 0x02;


        //Msg type

        // SERVICE_TYPE_FILE
        public const byte MSG_TYPE_FILE_READ_REQ = 0x01;
        public const byte MSG_TYPE_FILE_READ_RESP = 0x01;

        public const byte MSG_TYPE_FILE_WRITE_REQ = 0x02;
        public const byte MSG_TYPE_FILE_WRITE_RESP = 0x02;


        // SERVICE_TYPE_CMD
        public const byte MSG_TYPE_CMD_RESET_REQ = 0x01;
        public const byte MSG_TYPE_CMD_RESET_RESP = 0x01;

        public const byte MSG_TYPE_CMD_SWITCH_SCREEN_REQ = 0x02;
        public const byte MSG_TYPE_CMD_SWITCH_SCREEN_RESP = 0x02;

        public const byte MSG_TYPE_CMD_TEST_RESULT_REQ = 0x03;
        public const byte MSG_TYPE_CMD_TEST_RESULT_RESP = 0x03;

        public const byte MSG_TYPE_CMD_GET_MCU_ID_REQ = 0x04;
        public const byte MSG_TYPE_CMD_GET_MCU_ID_RESP = 0x04;

        public const byte MSG_TYPE_CMD_START_TEST_REQ = 0x05;
        public const byte MSG_TYPE_CMD_START_TEST_RESP = 0x05;

        public const byte MSG_TYPE_CMD_WARN_REQ = 0x06;
        public const byte MSG_TYPE_CMD_WARN_RESP = 0x06;

        public const byte MSG_TYPE_CMD_BOOT_IN_IAP_MODE_REQ = 0x07;
        public const byte MSG_TYPE_CMD_BOOT_IN_IAP_MODE_RESP = 0x07;

        public const byte MSG_TYPE_CMD_BOOT_IN_SLEEP_MODE_REQ = 0x08;
        public const byte MSG_TYPE_CMD_BOOT_IN_SLEEP_MODE_RESP = 0x08;

        public const byte MSG_TYPE_CMD_CRYPT_REQ = 0x09;
        public const byte MSG_TYPE_CMD_CRYPT_RESP = 0x09;

        public const byte MSG_TYPE_CMD_WORK_CURRENT_REQ = 0x0A;
        public const byte MSG_TYPE_CMD_WORK_CURRENT_RESP = 0x0A;

        public const byte MSG_TYPE_CMD_SLEEP_CURRENT_REQ = 0x0B;
        public const byte MSG_TYPE_CMD_SLEEP_CURRENT_RESP = 0x0B;

        public const byte MSG_TYPE_CMD_TEST_MCU_CODE_REQ = 0x0D;
        public const byte MSG_TYPE_CMD_TEST_MCU_CODE_RESP = 0x0D;

        public const byte MSG_TYPE_CMD_TEST_MCU_CODE_RESTART = 0x0E;
        public const byte MSG_TYPE_CMD_TEST_MCU_CODE_RESTART_RESP = 0x0E;

        public const byte MSG_TYPE_CMD_BEEP_REQ = 0x0F;
        public const byte MSG_TYPE_CMD_BEEP_RESP = 0x0F;

        public const byte MSG_TYPE_CMD_CAN_SENSOR_TEST_REQ = 0x10;
        public const byte MSG_TYPE_CMD_CAN_SENSOR_TEST_RESP = 0x10;

        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ = 0x18;
        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_RESP = 0x18;

        public int commBusy = 0;
        public int CommBuffWriteIndex = 0;
        public int CommBuffReadIndex = 0;


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
            time = 1,
            BarNum = 2,             //条形码
            CAN_SENSOR = 3,
        }


        string[] columnName = { "时间" , "条形码", "CAN SENSOR 状态" };

        public int startColumn = (int)column.time;
        public int endColumn = (int)column.CAN_SENSOR;

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


            public int StaticVoltage_OK_NOT;
            public int Static_Curent_OK_NOT;
            public int Work_Curent_OK_NOT;

            public int Voltage1_OK_NOT;  //0：没结果   1：正常    2：过小   3：过大
            public int Voltage2_OK_NOT;
            public int Voltage3_OK_NOT;
            public int Voltage4_OK_NOT;
            public int Voltage5_OK_NOT;
            public int Voltage6_OK_NOT;
            public int Voltage7_OK_NOT;
            public int Voltage8_OK_NOT;
            public int Voltage9_OK_NOT;


            public int CAN_SENSOR_OK_NOT;

            public int WIFI_OK_NOT;
            public int CAN_OK_NOT;
            public int SWITCH_SCREEN_OK_NOT;

            public byte Mode;           // 0:  none 1 : test   2:  work   3.  sleep 
        };
        tlv_reply_t tlv_reply;
        tlv_reply_t tlv_reply_min;
        tlv_reply_t tlv_reply_max;


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

        public struct TestResault
        {
            public byte[] data;
            public int returnF;
        };
        TestResault testresault;
        public CAN_Sensor_Test()
        {
            InitializeComponent();
            commBuf = new byte[4096];
            //commWifiTestBuf = new byte[4096 * 10];
            testresault.data = new byte[0x15];
        }


        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }

        public int test_run = 0;
        private void btn_Test_self_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }
            if (test_run == 1)
            {
                btn_Test_self.Text = "开始测试";
                test_run = 0;
            }
            else
            {
                btn_Test_self.Text = "停止测试";
                test_run = 1;
                BarNum = textBox2.Text.ToString();
            }
        }


        Microsoft.Office.Interop.Excel.Range ran;
        public void saveTestResault()
        {
            int blankRow = 0;
            //excel 增加一行
            try
            {
                for (int i = 1; i < 65536; i++)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, 1];
                    if (ran.Value == null)
                    {
                        blankRow = i;
                        break;
                    }
                }
                ws.Cells[blankRow, (int)column.time] = DateTime.Now.ToLocalTime().ToString();// time
                ws.Cells[blankRow, (int)column.BarNum] = BarNum;   // barNum
                if (commBuf[0x07] == 1)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.CAN_SENSOR];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.CAN_SENSOR] = "N0";
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.CAN_SENSOR];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.CAN_SENSOR] = "OK";
                }

                xApp.DisplayAlerts = false;
                ws.SaveAs("CAN SENSOR模块测试结果");
                if (wb != null)
                {
                    //wb.Save();
                    wb.SaveAs(Cdirectory + "\\CAN SENSOR测试结果.xlsx");
                    wb.Saved = true;
                }
            }
            catch (Exception Ex2)
            {
                MessageBox.Show(Ex2.Message + " , 请重新测试");
                Process[] localAll = Process.GetProcesses();
                foreach (Process A in localAll)
                {
                    if (A.ProcessName.ToString() == "EXCEL")
                        A.Kill();
                    //textBox1.Text += "\r\n" + A.ProcessName.ToString();
                }
                return;
            }
            
        }
        public void can_sendor_test()
        {
            while (true)
            {
                if (test_run == 1)
                {
                    ACKinfo.ACK_CMD_CAN_SENSOR_TEST = 0;
                    sendCMD(MSG_TYPE_CMD_CAN_SENSOR_TEST_REQ);
                    for (int i = 0; i < 5; i++)
                    {
                        if (ACKinfo.ACK_CMD_CAN_SENSOR_TEST == 1)
                            break;
                        else
                            Thread.Sleep(1000);
                    }
                    if (ACKinfo.ACK_CMD_CAN_SENSOR_TEST == 1)
                    {
                        if (commBuf[0x07] == 1)    //CAN SENSOR 不正常
                        {
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Red;
                                textBox1.Text = "NO";
                            }));
                            this.Invoke((EventHandler)(delegate { checkBox5.Checked = false; checkBox6.Checked = true; }));
                            //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n" + DateTime.Now.ToLocalTime().ToString() + " : CAN SENSOR 不正常\r\n"; }));
                        }
                        else     //CAN SENSOR 正常
                        {
                      
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Green;
                                textBox1.Text = "PASS";
                            }));
                            this.Invoke((EventHandler)(delegate { checkBox5.Checked = true; checkBox6.Checked = false; }));
                            //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n" + DateTime.Now.ToLocalTime().ToString() + " : CAN SENSOR 正常\r\n"; }));
                        }
                        //保存结果
                        saveTestResault();
                        test_run = 0;
                        this.Invoke((EventHandler)(delegate { btn_Test_self.Text = "开始测试"; }));
                    }
                    else
                    {
                        this.Invoke((EventHandler)(delegate{ textBox1.Text += "\r\n" + DateTime.Now.ToLocalTime().ToString() + " : CAN SENSOR 测试超时\r\n"; }));
                    }
                }
                else
                    Thread.Sleep(1000);
            }
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
            public int ACK_CMD_CAN_SENSOR_TEST;
            public int ACK_CMD_LOOK_FOR_COM;
        };
        public ACKINFO ACKinfo;

        public void commDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            //搬移combuf，先整理一下combuf
            if (((CommBuffWriteIndex - CommBuffReadIndex) == 0))
            {
                CommBuffWriteIndex = 0;
                CommBuffReadIndex = 0;

                for (int i = 0; i < CommBuffWriteIndex; i++)
                {
                    commBuf[i] = 0;
                }
            }
            else if (((CommBuffWriteIndex - CommBuffReadIndex) > 0) & (CommBuffReadIndex > 0))
            {
                for (int i = 0; i < CommBuffWriteIndex - CommBuffReadIndex; i++)
                {
                    commBuf[i] = commBuf[i + CommBuffReadIndex];
                    commBuf[i + CommBuffReadIndex] = 0;
                }

                CommBuffReadIndex = 0;
                CommBuffWriteIndex = CommBuffWriteIndex - CommBuffReadIndex;
            }

            //将接收到的数据写入commbuf
            int len = comm.BytesToRead;
            comm.Read(commBuf, CommBuffWriteIndex, len);
            CommBuffWriteIndex += len;

            //获取ACK状态
            DealWithACK();

        }

        public bool completedFrame()
        {
            if ((CommBuffWriteIndex - CommBuffReadIndex) >= 16)
            {
                if ((commBuf[CommBuffReadIndex] == 0xAA) & (commBuf[CommBuffReadIndex + 1] == 0xAA) & (commBuf[CommBuffReadIndex + 2] == 0xAA) & (commBuf[CommBuffReadIndex + 3] == 0xAA))
                {
                    if ((CommBuffWriteIndex - CommBuffReadIndex) >= commBuf[CommBuffReadIndex + 4] + commBuf[CommBuffReadIndex + 5] * 256)
                        return true;
                    else
                        return false;
                }
                else
                {
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Text += "\r\n从测试板 发送 给PC的数据 可能存在丢失！";
                    }));
                    for (int start = 0; start < CommBuffWriteIndex; start++)
                    {
                        if ((commBuf[start] == 0xAA) & (commBuf[start + 1] == 0xAA) & (commBuf[start + 2] == 0xAA) & (commBuf[start + 3] == 0xAA))
                        {
                            CommBuffReadIndex = start;
                        }
                    }
                    return false;
                }
            }
            else
                return false;

        }

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
        public void DealWithACK()
        {
            byte serviceType; byte message; int valid = 1;
            //分析commbuf
            if (completedFrame())
            {
                serviceType = commBuf[0x0a];
                message = commBuf[0x0b];

                if (serviceType == SERVICE_TYPE_FILE)
                {
                    switch (message)
                    {
                        case MSG_TYPE_FILE_READ_RESP:
                            ACKinfo.ACK_File_Read = 1;
                            break;
                        case MSG_TYPE_FILE_WRITE_RESP:
                            ACKinfo.ACK_File_Write = 1;
                            break;
                        default:
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Text += "\r\n无效消息类型";
                            }));
                            break;
                    }
                }
                else if (serviceType == SERVICE_TYPE_CMD)
                {
                    switch (message)
                    {
                        case MSG_TYPE_CMD_RESET_RESP:
                            ACKinfo.ACK_CMD_Rest = 1;
                            break;
                        case MSG_TYPE_CMD_SWITCH_SCREEN_RESP:
                            ACKinfo.ACK_CMD_Switch_Screen = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_RESULT_RESP:
                            {
                                msg_head_t msg_head = new msg_head_t();
                                msg_head = (msg_head_t)BytesToStruct(commBuf, typeof(msg_head_t));
                                byte[] tmpbuff = new byte[1088];
                                for (int i = 0; i < msg_head.msg_len - 16; i++)
                                {
                                    tmpbuff[i] = commBuf[i + 16];
                                }
                                tlv_reply = (tlv_reply_t)BytesToStruct(tmpbuff, typeof(tlv_reply_t));
                                ACKinfo.ACK_CMD_Test_Resualt = 1;
                            }
                            break;
                        case MSG_TYPE_CMD_GET_MCU_ID_RESP:
                            ACKinfo.ACK_CMD_Get_Mcu_ID = 1;
                            break;
                        case MSG_TYPE_CMD_START_TEST_RESP:
                            ACKinfo.ACK_CMD_Start_Test = 1;
                            break;
                        case MSG_TYPE_CMD_WARN_RESP:
                            ACKinfo.ACK_CMD_Warn = 1;
                            break;
                        case MSG_TYPE_CMD_BOOT_IN_IAP_MODE_RESP:
                            ACKinfo.ACK_CMD_BOOT_IN_IAP_MODE = 1;
                            break;
                        case MSG_TYPE_CMD_BOOT_IN_SLEEP_MODE_RESP:
                            ACKinfo.ACK_CMD_BOOT_IN_SLEEP_MODE = 1;
                            break;
                        case MSG_TYPE_CMD_CRYPT_RESP:
                            ACKinfo.ACK_CMD_CRYPT = 1;
                            break;
                        case MSG_TYPE_CMD_WORK_CURRENT_RESP:
                            ACKinfo.ACK_CMD_WORK_CURRENT = 1;
                            break;
                        case MSG_TYPE_CMD_SLEEP_CURRENT_RESP:
                            ACKinfo.ACK_CMD_SLEEP_CURRENT = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_MCU_CODE_RESP:
                            ACKinfo.ACK_CMD_TEST_MCU_CODE = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_MCU_CODE_RESTART_RESP:
                            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 1;
                            break;
                        case MSG_TYPE_CMD_BEEP_RESP:
                            ACKinfo.ACK_CMD_BEEP = 1;
                            break;
                        case MSG_TYPE_CMD_CAN_SENSOR_TEST_RESP:
                            ACKinfo.ACK_CMD_CAN_SENSOR_TEST = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ:
                            ACKinfo.ACK_CMD_LOOK_FOR_COM = 1;
                            break;
                        default:
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Text += "\r\n无效消息类型";
                            }));
                            break;
                    }
                }
                else
                {
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Text += "\r\n无效服务类型！";
                    }));
                }

                CommBuffReadIndex += commBuf[4] + commBuf[5] * 256;
            }
        }

        private void CAN_Sensor_Test_Load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            Array.Sort(ports);
            timer1.Start();


            //btn_Test_self.Enabled = false;

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
            ACKinfo.ACK_CMD_CAN_SENSOR_TEST = 0;
            ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;

            //打开测试板串口
            try
            {
                foreach (string port in ports)
                {
                    comm = new SerialPort();
                    openCOM(port);
                    if (comm.IsOpen)
                        sendCMD(MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ, 0x00);
                    else
                        continue;
                    for (int i = 0; i < 3; i++)
                    {
                        Thread.Sleep(500);
                        if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)  //open com success ，no more wait
                            break;
                    }

                    if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)      //open com success ，no more looking for
                    {
                        ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;
                        //btn_write_bin.Enabled = true;
                        textBox1.Text = "OPEN TEST BOARD COM SUCCESS";
                        //COMMPORT = port;
                        break;
                    }
                    else                                        //no right com
                    {
                        comm.Close();
                    }
                }
                if (!comm.IsOpen)
                {
                    MessageBox.Show("请检查PC机与测试板之间的连接！");
                    return;
                }

                Thread CAN_SENOR_TEST = new Thread(can_sendor_test);
                CAN_SENOR_TEST.Start();
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }
            finally
            {
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

                if (File.Exists(Cdirectory + "\\CAN SENSOR测试结果.xlsx"))
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(Cdirectory + "\\CAN SENSOR测试结果.xlsx");
                    ws = wb.Worksheets["CAN SENSOR测试结果"];
                }
                else
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(true);
                    ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = "CAN SENSOR测试结果";
                    for (int i = startColumn; i <= endColumn; i++)
                    {
                        ws.Cells[1, i] = columnName[i - 1];
                    }
                }
            }
        }

        public void openCOM(string portName)
        {
            //关闭时点击，则设置好端口，波特率后打开
            comm.PortName = portName;
            comm.BaudRate = 115200;
            comm.DataBits = 8;
            comm.Parity = Parity.None;
            comm.StopBits = StopBits.One;
            comm.WriteBufferSize = 1024 * 1024 * 5;
            comm.DataReceived += new SerialDataReceivedEventHandler(commDataReceivedHandler);
            try
            {
                if (comm.IsOpen)
                    return;
                comm.Open();

                //建立串口接收线程
            }
            catch (Exception ex)
            {
                //捕获到异常信息，创建一个新的comm对象，之前的不能用了。
                comm = new SerialPort();
                //现实异常信息给客户。
                textBox1.Text += "\r\n" + ex.Message;
            }
        }
        public void sendCMD(byte REQ, byte Opration)
        {
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE_CMD;
            msg_head.msg_type = REQ;
            msg_head.resp = Opration;
            msg_head.msg_len =
            (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head));
            msg_head.crc32 = 0;// GetCrc32(framebuff, msg_len);
            int len = System.Runtime.InteropServices.Marshal.SizeOf(msg_head);
            byte[] tmpbuff = StructToBytes(msg_head);
            //发送测试指令
            while ((commBusy == 1) | (!comm.IsOpen)) ;
            try
            {
                commBusy = 1;
                if (comm.IsOpen)
                {
                    comm.Write(tmpbuff, 0, len);
                }
                else
                {
                    comm.Open();
                    comm.Write(tmpbuff, 0, len);
                }
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;
                }));

                MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
                comm.Close();
            }
            finally
            {
                commBusy = 0;
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < (int)timeF.TotalTimer; i++)
                if (TimerArry[i] > 0)
                    TimerArry[i]--;
        }

        public static byte[] StructToBytes(object structObj)
        {
            //得到结构体的大小
            int size = Marshal.SizeOf(structObj);
            //创建byte数组
            byte[] bytes = new byte[size];
            //分配结构体大小的内存空间
            IntPtr structPtr = Marshal.AllocHGlobal(size);
            //将结构体拷到分配好的内存空间
            Marshal.StructureToPtr(structObj, structPtr, false);
            //从内存空间拷到byte数组
            Marshal.Copy(structPtr, bytes, 0, size);
            //释放内存空间
            Marshal.FreeHGlobal(structPtr);
            //返回byte数组
            return bytes;
        }
        public void sendCMD(byte REQ)
        {
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE_CMD;
            msg_head.msg_type = REQ;
            msg_head.msg_len =
            (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head));
            msg_head.crc32 = 0;// GetCrc32(framebuff, msg_len);

            int len = System.Runtime.InteropServices.Marshal.SizeOf(msg_head);
            byte[] tmpbuff = StructToBytes(msg_head);
            //发送测试指令
            while ((commBusy == 1) | (!comm.IsOpen)) ;
            try
            {
                commBusy = 1;
                if (comm.IsOpen)
                {
                    comm.Write(tmpbuff, 0, len);
                }
                else
                {
                    comm.Open();
                    comm.Write(tmpbuff, 0, len);
                }
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;
                }));

                MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
                comm.Close();
            }
            finally
            {
                commBusy = 0;
            }
        }

        private void CAN_Sensor_Test_FormClosing(object sender, FormClosingEventArgs e)
        {
            
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            xApp.Quit();
            wb = null;
            wbs = null;
            xApp = null;

            if (comm != null)
                if (comm.IsOpen)
                    comm.Close();


            Process[] localAll = Process.GetProcesses();
            foreach (Process A in localAll)
            {
                if (A.ProcessName.ToString() == "EXCEL")
                    A.Kill();
                //textBox1.Text += "\r\n" + A.ProcessName.ToString();
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                if (checkBox6.Checked == true)
                {
                    checkBox5.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                if (checkBox5.Checked == true)
                {
                    checkBox6.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }
    }
}

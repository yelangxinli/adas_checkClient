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
using System.Text.RegularExpressions;
using System.Management;

namespace ADASLeader产品生产测试
{
    public partial class WIFI_Board_Test : Form
    {
        public WIFI_Board_Test()
        {
            InitializeComponent();
            commBuf = new byte[4096];
            commWifiTestBuf = new byte[4096 * 10];
            testresault.data = new byte[0x15];
            testresault.returnF = 0;
            timer2.Interval = 3000;
            timer2.Start();
            timer1.Interval = 1000;
            timer1.Start();

            //测试
            Threa_V_I_Test = new Thread(selfTest);
            Threa_V_I_Test.Start();
            netEntries = new SlWlanNetworkEntry_t[30];
            for (int i = 0; i < 30; i++)
            {
                netEntries[i].Ssid = new byte[32];
                netEntries[i].Bssid = new byte[8];
            }
        }

        public Thread Threa_V_I_Test;
        public int PASS_FLAG; //1：测试通过，其他：测试未通过

        public string mFilename;
        private string BarNum;
        public byte[] commBuf;
        public byte[] commWifiTestBuf;
        public SerialPort comm;
        public SerialPort commWifiTest;
        public int commWifiTest_returnF;
        public string commWifiWriteBIN = null;
        public int WIFITest_RUN = 0;
        public enum timeF
        {
            selfTest = 0,
            wifiTest = 1,
            lookforwifi = 2,
            TotalTimer = 16,
        };
        int[] TimerArry = new int[16];
        int heartBeat = 0;



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
            func = 3,
            rssi = 4,
            wifiName = 5
        }

        string[] columnName = { "时间", "条形码", "功能", "RRSI" ,"WIFI名称"};

        public int startColumn = (int)column.time;
        public int endColumn = (int)column.wifiName;

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
            public int ACK_CMD_WIFI_Test_Resualt;
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
            public int ACK_CMD_CANSENSOR;
            public int ACK_CMD_WIFI_CODE_CONNECT;
            public int ACK_CMD_WIFI_CODE_ERASE_SRAM;
            public int ACK_CMD_WIFI_CODE_WRITE_SRAM_PATCH;
            public int ACK_CMD_WIFI_CODE_ERASE_FLASH;
            public int ACK_CMD_WIFI_CODE_WRITE_FLASH_PATCH;
            public int ACK_CMD_WIFI_CODE_WRITE_FS;
            public int ACK_CMD_WIFI_CODE_WRITE_CHUNK;
            public int ACK_CMD_LOOK_FOR_COM;
        };

        public ACKINFO ACKinfo;
        public string COMMPORT = null;


        public enum HardwareEnum
        {
            // 硬件
            Win32_Processor, // CPU 处理器
            Win32_PhysicalMemory, // 物理内存条
            Win32_Keyboard, // 键盘
            Win32_PointingDevice, // 点输入设备，包括鼠标。
            Win32_FloppyDrive, // 软盘驱动器
            Win32_DiskDrive, // 硬盘驱动器
            Win32_CDROMDrive, // 光盘驱动器
            Win32_BaseBoard, // 主板
            Win32_BIOS, // BIOS 芯片
            Win32_ParallelPort, // 并口
            Win32_SerialPort, // 串口
            Win32_SerialPortConfiguration, // 串口配置
            Win32_SoundDevice, // 多媒体设置，一般指声卡。
            Win32_SystemSlot, // 主板插槽 (ISA & PCI & AGP)
            Win32_USBController, // USB 控制器
            Win32_NetworkAdapter, // 网络适配器
            Win32_NetworkAdapterConfiguration, // 网络适配器设置
            Win32_Printer, // 打印机
            Win32_PrinterConfiguration, // 打印机设置
            Win32_PrintJob, // 打印机任务
            Win32_TCPIPPrinterPort, // 打印机端口
            Win32_POTSModem, // MODEM
            Win32_POTSModemToSerialPort, // MODEM 端口
            Win32_DesktopMonitor, // 显示器
            Win32_DisplayConfiguration, // 显卡
            Win32_DisplayControllerConfiguration, // 显卡设置
            Win32_VideoController, // 显卡细节。
            Win32_VideoSettings, // 显卡支持的显示模式。

            // 操作系统
            Win32_TimeZone, // 时区
            Win32_SystemDriver, // 驱动程序
            Win32_DiskPartition, // 磁盘分区
            Win32_LogicalDisk, // 逻辑磁盘
            Win32_LogicalDiskToPartition, // 逻辑磁盘所在分区及始末位置。
            Win32_LogicalMemoryConfiguration, // 逻辑内存配置
            Win32_PageFile, // 系统页文件信息
            Win32_PageFileSetting, // 页文件设置
            Win32_BootConfiguration, // 系统启动配置
            Win32_ComputerSystem, // 计算机信息简要
            Win32_OperatingSystem, // 操作系统信息
            Win32_StartupCommand, // 系统自动启动程序
            Win32_Service, // 系统安装的服务
            Win32_Group, // 系统管理组
            Win32_GroupUser, // 系统组帐号
            Win32_UserAccount, // 用户帐号
            Win32_Process, // 系统进程
            Win32_Thread, // 系统线程
            Win32_Share, // 共享
            Win32_NetworkClient, // 已安装的网络客户端
            Win32_NetworkProtocol, // 已安装的网络协议
            Win32_PnPEntity,//all device
        }

        public static string[] MulGetHardwareInfo(HardwareEnum hardType, string propKey)
        {

            List<string> strs = new List<string>();
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("select * from " + hardType))
                {
                    var hardInfos = searcher.Get();
                    foreach (var hardInfo in hardInfos)
                    {
                        if (hardInfo.Properties[propKey].Value != null)
                        {
                            if (hardInfo.Properties[propKey].Value.ToString().Contains("COM"))
                            {
                                strs.Add(hardInfo.Properties[propKey].Value.ToString());
                            }
                        }
                    }
                    searcher.Dispose();
                }
                return strs.ToArray();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return strs.ToArray();
            }
            finally
            {
                strs = null;
            }
        }
        private void FRM_load(object sender, EventArgs e)
        {
            commWifiTest = new SerialPort();
            string[] ports = SerialPort.GetPortNames();
            string[] port1;
            Array.Sort(ports);
            timer1.Start();

            string[] ss = MulGetHardwareInfo(HardwareEnum.Win32_PnPEntity, "Name");

            btn_write_bin.Enabled = true;
            btn_Test_self.Enabled = false;
            btn_submit.Enabled = false;
            btn_look_for_wifi.Enabled = false;

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
            ACKinfo.ACK_CMD_BEEP = 0;
            ACKinfo.ACK_CMD_CANSENSOR = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_CONNECT = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_FLASH = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_SRAM = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FLASH_PATCH = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FS = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_SRAM_PATCH = 0;
            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_CHUNK = 0;
            ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;
            ACKinfo.ACK_CMD_WIFI_Test_Resualt = 0;

            //打开测试板串口
            try
            {
                //打开测试板串口
                string commport = null;
                foreach (string a in ss)
                {
                    if (a.Contains("STMicroelectronics Virtual COM Port"))
                    {
                        commport = a.Substring(a.IndexOf("(COM") + 1);
                        commport = commport.Substring(0, commport.IndexOf(")"));
                    }
                }
                comm = new SerialPort();
                openCOM(commport);
                if (comm.IsOpen)
                    sendCMD(MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ, 0x00);
                for (int i = 0; i < 3; i++)
                {
                    Thread.Sleep(500);
                    if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)  //open com success ，no more wait
                        break;
                }

                if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)      //open com success ，no more looking for
                {
                    ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;
                    textBox1.Text = "OPEN VIRTUAL COM PORT SUCCESS";
                }
                else                                        //no right com
                {
                    comm.Close();
                }
                if (!comm.IsOpen)
                {
                    textBox1.Text += "\r\n请检查PC主机与测试板之间的连接，并保证测试板已经上电!\r\n\r\n\r\n请重启软件！";
                    return;
                }
                else
                {
                    if (ports.Contains("COM1"))
                        port1 = new string[ports.Length - 2];
                    else
                        port1 = new string[ports.Length - 1];

                    for (int i = 0, m = 0; i < ports.Length; i++)
                    {
                        if ((ports[i] != "COM1") && (ports[i] != comm.PortName.ToString()))
                        {
                            port1[m] = ports[i];
                            m++;
                        }
                    }
                    if (port1 != null)
                    {
                        comboBox1.Items.AddRange(port1);
                        comboBox2.Items.AddRange(port1);
                    }



                    //commWifiTest = new SerialPort();
                    foreach (string port in port1)
                    {
                        //textBox1.Text += "\r\n" + port;
                        if (!commWifiTest.IsOpen)
                            openWifiTestCOM(port);
                        if (commWifiTest.IsOpen)
                            sendCMD(MSG_TYPE_CMD_TEST_WIFI_TEST_REQ, 0x00);//重启wifi 测试板
                        else
                            continue;
                        commWifiTest_returnF = 0;
                        for (int i = 0; i < 100; i++)
                        {
                            Thread.Sleep(100);
                            if (commWifiTest_returnF == 1)  //open com success ，no more wait
                                break;
                        }

                        if (commWifiTest_returnF == 1)      //open com success ，no more looking for
                        {
                            textBox1.Text += "\r\nOPEN WIFI TEST UART COM SUCCESS";
                            comboBox1.Text = commWifiTest.PortName.ToString();
                            break;
                        }
                        else                                        //no right com
                        {
                            commWifiTest.Close();
                        }
                    }
                    if (!commWifiTest.IsOpen)
                    {
                        textBox1.Text += "\r\nOPEN WIFI TEST UART COM FAILED";
                    }

                }


                foreach (string port in ports)
                {
                    if (ports.Contains("COM1"))
                    {
                        if (ports.Length == 4)
                        {
                            if (comm.IsOpen && commWifiTest.IsOpen)
                            {
                                if ((port != "COM1") && (port != comm.PortName) && (port != commWifiTest.PortName))
                                {
                                    commWifiWriteBIN = port;
                                    comboBox2.Text = commWifiWriteBIN;
                                }
                            }
                        }
                        else
                        {
                            textBox1.Text += "\r\n请拔掉其余的串口工具";
                        }
                    }
                    else
                    {
                        if (ports.Length == 3)
                        {
                            if (comm.IsOpen && commWifiTest.IsOpen)
                            {
                                if ((port != comm.PortName) && (port != commWifiTest.PortName))
                                {
                                    commWifiWriteBIN = port;
                                    comboBox2.Text = commWifiWriteBIN;
                                }
                                    
                            }
                        }
                        else
                        {
                            textBox1.Text += "\r\n请拔掉其余的串口工具";
                        }
                    }
                }
                
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

                if (File.Exists(Cdirectory + "\\WIFI模块测试结果.xlsx"))
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Open(Cdirectory + "\\WIFI模块测试结果.xlsx");
                    ws = wb.Worksheets["WIFI模块测试结果"];
                }
                else
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(true);
                    ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = "WIFI模块测试结果";
                    for (int i = startColumn; i <= endColumn; i++)
                    {
                        ws.Cells[1, i] = columnName[i - 1];
                    }
                }

                Thread.Sleep(100);
                EXCEL_PROCESS = Process.GetProcesses();
            }     
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }
        

        private void timer1_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < (int)timeF.TotalTimer; i++)
                if (TimerArry[i] > 0)
                    TimerArry[i]--;
        }

        

        public void openCOM(string portName)
        {
            //关闭时点击，则设置好端口，波特率后打开
            comm.ReadTimeout = 1000;
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
                comm.DiscardInBuffer();
                comm.DiscardOutBuffer();

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


        public bool openWifiTestCOM(string portName)
        {
            //关闭时点击，则设置好端口，波特率后打开
            commWifiTest.PortName = portName;
            commWifiTest.BaudRate = 115200;
            commWifiTest.DataBits = 8;
            commWifiTest.Parity = Parity.None;
            commWifiTest.StopBits = StopBits.One;
            commWifiTest.WriteBufferSize = 1024 * 1024 * 5;
            commWifiTest.ReadBufferSize = 1024 * 1024 * 5;
            commWifiTest.DataReceived += new SerialDataReceivedEventHandler(commWifiTestDataReceivedHandler);
            try
            {
                if (commWifiTest.IsOpen)
                    return false;
                commWifiTest.Open();
                return true;
                //建立串口接收线程
            }
            catch (Exception ex)
            {
                //捕获到异常信息，创建一个新的comm对象，之前的不能用了。
                commWifiTest = new SerialPort();
                //现实异常信息给客户。
                textBox1.Text += "\r\n" + ex.Message;
            }
            return false;
        }


        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }

        private void btn_write_bin_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }

            this.Invoke((EventHandler)(delegate
            {
                textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                textBox1.ForeColor = Color.Black;
            }));

            BarNum = textBox2.Text.ToString();
            this.Invoke((EventHandler)(delegate
            {
                btn_Test_self.Enabled = false;
                btn_submit.Enabled = false;
                btn_look_for_wifi.Enabled = false;
                btn_Test_self.Text = "开始测试";
            }));
            WIFITest_RUN = 0;
            textBox1.Text = "烧录中。。。";
            btn_write_bin.Enabled = false;
            //下载MCU  BIN文件  线程
            Thread DownLoadMCUBINThred = new Thread(downLoadWIFIBIN);
            DownLoadMCUBINThred.Start();

            this.Invoke((EventHandler)(delegate {
                int count = listView1.Items.Count;
                for (int i = 0; i < count; i++)
                {
                    ADASLeader_device.RemoveAt(0);
                    listView1.Items.Remove(listView1.Items[0]);
                }
            }));
        }

        private void btn_Test_self_Click(object sender, EventArgs e)
        {
            if (btn_Test_self.Text.ToString() == "开始测试")
            {
                btn_submit.Enabled = false;
                WIFITest_RUN = 1;
                btn_Test_self.Text = "停止测试";
                btn_write_bin.Enabled = false;
                btn_look_for_wifi.Enabled = false;
            }
            else
            {
                btn_submit.Enabled = true;
                btn_Test_self.Text = "开始测试";
                WIFITest_RUN = 0;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            heartBeat = 0;
        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            TimerArry[(int)timeF.selfTest] = TimerArry[(int)timeF.selfTest] - 1;
            TimerArry[(int)timeF.wifiTest] = TimerArry[(int)timeF.wifiTest] - 1;
            TimerArry[(int)timeF.lookforwifi] = TimerArry[(int)timeF.lookforwifi] - 1;

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                ListViewItem item = listView1.Items[i];
                if (item.Selected == true)
                {
                    btn_Test_self.Enabled = true;
                    break;
                }
            }
        }

        private void btn_openCOM_Click(object sender, EventArgs e)
        {
            if (commWifiTest.IsOpen)
                commWifiTest.Close();
            if ((comboBox1.Text == null) || (comboBox2.Text == null))
            {
                MessageBox.Show("请选择串口");
                return;
            }

            if ((comboBox1.Text == "COM1") || (comboBox2.Text == "COM1"))
            {
                MessageBox.Show("请不要选择COM1");
                return;
            }

            if ((comboBox1.Text == comm.PortName.ToString()) || (comboBox2.Text == comm.PortName.ToString()))
            {
                MessageBox.Show(comm.PortName.ToString() + "是测试主板虚拟串口，请选择其余串口");
                return;
            }

            //重启WIFI测试模块
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_TEST_REQ, 0x00);

            
            openWifiTestCOM(comboBox1.Text);
            commWifiWriteBIN = comboBox2.Text;


            //重启WIFI测试模块
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_TEST_REQ, 0x00);
            
            commWifiTest_returnF = 0;
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(1000);
                if (commWifiTest_returnF == 1)  //open com success ，no more wait
                    break;
            }

            if((commWifiTest_returnF == 1))
            if (commWifiTest.IsOpen)
            {
                textBox1.Text += "\r\n打开串口";
            }
        }

        Microsoft.Office.Interop.Excel.Range ran;
        private void btn_submit_Click(object sender, EventArgs e)
        {
            int blankRow = 0; 
            //excel 增加一行
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
            if ((testresault.data[16] == 0x01) && (testresault.data[17] == 0x01))   //func
            {
                ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.func];
                ran.Interior.Color = Color.Green;
                ws.Cells[blankRow, (int)column.func] = "OK";  
            }
            else
            {
                ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.func];
                ran.Interior.Color = Color.Red;
                ws.Cells[blankRow, (int)column.func] = "NO";   
            }

            ws.Cells[blankRow, (int)column.rssi] = textBox_rssi.Text.ToString();//rssi
            if (testresault.data[20] == 0xff)   
            {
                ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.rssi];
                ws.Cells[blankRow, (int)column.rssi] = "NO";
                ran.Interior.Color = Color.Red;
            }
            else
            {
                ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.rssi];
                ws.Cells[blankRow, (int)column.rssi] = "OK";
                ran.Interior.Color = Color.Green;
            }

            ListViewItem item = new ListViewItem();
            for (int i = 0; i < ADASLeader_device.Count; i++)
            {
                item = listView1.Items[i];
                if(item.Selected == true)
                    ws.Cells[blankRow, (int)column.wifiName] = ADASLeader_device[i].name.ToString();//wifi name
            }

            //ws.SaveAs("eee");
            //wb.SaveAs("eeee");

            xApp.DisplayAlerts = false;
            ws.SaveAs("WIFI模块测试结果");
            if (wb != null)
            {
                wb.SaveAs(Cdirectory + "\\WIFI模块测试结果.xlsx");
                wb.Saved = true;
                //wb.Close();
                //wb = null;
            }
            
            btn_Test_self.Text = "开始测试";
            WIFITest_RUN = 0;

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                item = listView1.Items[i];
                if (item.Selected == true)
                {
                    item.Selected = false;
                }
            }

            this.Invoke((EventHandler)(delegate {
                int count = listView1.Items.Count;
                for (int i = 0; i < count; i++)
                {
                    ADASLeader_device.RemoveAt(0);
                    listView1.Items.Remove(listView1.Items[0]);
                }
            }));
            btn_submit.Enabled = false;
            btn_look_for_wifi.Enabled = false;
            btn_Test_self.Enabled = false;
            //xApp.Quit();
        }

        public Process[] EXCEL_PROCESS;
        private void WIFI_Board_Test_FormClosing(object sender, FormClosingEventArgs e)
        {
            //wb.Close();
            //xApp.Quit();
            Threa_V_I_Test.Abort();
            Thread.Sleep(20);
            try
            {
                foreach (Process A in EXCEL_PROCESS)
                {
                    if (A.ProcessName.ToString() == "EXCEL")
                        A.Kill();
                }
                
                if (comm.IsOpen) { comm.Close(); };
                if (commWifiTest.IsOpen) { commWifiTest.Close(); };
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void btn_look_for_wifi_Click(object sender, EventArgs e)
        {
            Thread look_for_wifi_thread = new Thread(LOOK_FOR_WIFI);
            look_for_wifi_thread.Start();
            btn_write_bin.Enabled = false;
            btn_Test_self.Enabled = false;
            btn_submit.Enabled = false;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            ListViewItem item = new ListViewItem();
            for (int i = 0; i < ADASLeader_device.Count; i++)
            {
                item = listView1.Items[i];
                if (item.Selected == true)
                    item.BackColor = Color.DeepSkyBlue;
                else
                {
                    if (ADASLeader_device[i].is_tested == "否")
                    {
                        item.BackColor = Color.Yellow;
                    }
                    else
                    {
                        if(ADASLeader_device[i].is_pass == "是")
                            item.BackColor = Color.Green;
                        else
                            item.BackColor = Color.Red;
                    }
                }
            }
        }
    }
}
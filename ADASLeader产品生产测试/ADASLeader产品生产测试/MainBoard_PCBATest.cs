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
    public partial class MainBoard_PCBATest : Form
    {

        public int PASS_FLAG; //1：测试通过，其他：测试未通过

        public string mFilename;
        private string BarNum; 
        public byte[] commBuf;
        public byte[] commWifiTestBuf;
        public int mainboardBinWrite = 0;
        public MainBoard_PCBATest()
        {
            InitializeComponent();
            commBuf = new byte[4096];
            commWifiTestBuf = new byte[4096 * 10];
            testresault.data = new byte[0x15];
            testresault.returnF = 0;
            //测试
            Threa_WIFI_Test = new Thread(wifiselfTest);
            Threa_WIFI_Test.Start();
            netEntries = new SlWlanNetworkEntry_t[30];
            for (int i = 0; i < 30; i++)
            {
                netEntries[i].Ssid = new byte[32];
                netEntries[i].Bssid = new byte[8];
            }
        }
        public SerialPort comm;
        public enum timeF
        {
            selfTest = 0,
            lookforwifi = 1,
            wifiTest = 2,
            TotalTimer = 16,
        };
        int[] TimerArry = new int[16];

        public int SL_WLAN_SSID_MAX_LENGTH = 32;
        public int SL_WLAN_BSSID_LENGTH = 6;
        public struct SlWlanNetworkEntry_t
        {
            public byte[] Ssid;
            public byte[] Bssid;
            public byte SsidLen;
            public byte Rssi;
            public Int16 SecurityInfo;
            public byte Channel;
            public byte Reserved;
        };
        SlWlanNetworkEntry_t[] netEntries;


        public string Cdirectory;
        Microsoft.Office.Interop.Excel.Application xApp;
        Microsoft.Office.Interop.Excel.Workbooks wbs;
        Microsoft.Office.Interop.Excel.Workbook wb;
        Microsoft.Office.Interop.Excel.Worksheets wss;
        Microsoft.Office.Interop.Excel.Worksheet ws;

        public enum column {
            time = 1 ,               //时间
            BarNum = 2,             //条形码
            _12_V = 3,
            _5_V = 4,
            _3_3_V = 5,
            _2_5_V = 6,
            _1_8_V = 7,
            _1_2_V = 8,
            INPOWER1_V = 9,
            INPOWER2_V = 10,
            INPOWER3_V = 11,

            workcurrent = 12,
            staticcurrent = 13,

            SWITCH_VIDEO = 14,
            BEEP = 15,
            CAN = 16,
            WIFI_RSSI = 17,
            WIFI_FUC = 18,
            wifiName = 19
        }

        public enum Test_opration{
            beep = 0,
            Car_Video = 1,
            Mobileye_Video = 2,
            DVR_Video = 3,
            CAN_Test = 4,
        }

        string[] columnName = {"时间" , "条形码",  "12V", "IN_5V", "3.3V", "IN_2.5V", "IN_1.8V", "IN_1.2V", "IN_POWER1", "IN_POWER2", "IN_POWER3", "工作电流", "静态电流", "视频切换", "蜂鸣器", "CAN" , "RSSI" , "WIFI功能" , "WIFI名称"};

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
            

            public int WIFI_OK_NOT;
            public int CAN_OK_NOT;
            public int SWITCH_SCREEN_OK_NOT;

            public byte Mode;           // 0:  none 1 : test   2:  work   3.  sleep 
        };
        tlv_reply_t tlv_reply;
        tlv_reply_t tlv_reply_min;
        tlv_reply_t tlv_reply_max;
        

        test_resault Test_Resault;

        public int CommWifiTestBuffWriteIndex = 0;
        public int CommWifiTestBuffReadIndex = 0;
        public int commWifiTestBusy = 0;

        public SerialPort commWifiTest;
        public int commWifiTest_returnF;
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


        public void Compare_Reply(tlv_reply_t reply, tlv_reply_t replymax, tlv_reply_t replymin)
        {
            Test_Resault.Work_Curent_OK_NOT = ((reply.Current0 > replymin.Current0) && ((reply.Current0 < replymin.Current0))) ? 1 : 2;
            Test_Resault.Static_Curent_OK_NOT = ((reply.Current1 > replymin.Current1) && (reply.Current1 < replymin.Current1)) ? 1 : 2;

            Test_Resault.Voltage1_OK_NOT = (((reply.Voltage1 > replymin.Voltage1)) && (reply.Voltage1 < replymax.Voltage1)) ? 1 : 2;   //12V
            Test_Resault.Voltage2_OK_NOT = ((reply.Voltage2 > replymin.Voltage2) && (reply.Voltage2 < replymax.Voltage2)) ? 1 : 2;     //5V
            Test_Resault.Voltage3_OK_NOT = ((reply.Voltage0 > replymin.Voltage0) && (reply.Voltage0 < replymax.Voltage0)) ? 1 : 2;     //3.3V
            Test_Resault.Voltage4_OK_NOT = ((reply.Voltage3 > replymin.Voltage3) && (reply.Voltage3 < replymax.Voltage3)) ? 1 : 2;
            Test_Resault.Voltage5_OK_NOT = ((reply.Voltage4 > replymin.Voltage4) && (reply.Voltage4 < replymax.Voltage4)) ? 1 : 2;
            Test_Resault.Voltage6_OK_NOT = ((reply.Voltage5 > replymin.Voltage5) && (reply.Voltage5 < replymax.Voltage5)) ? 1 : 2;
            Test_Resault.Voltage7_OK_NOT = ((reply.Voltage6 > replymin.Voltage6) && (reply.Voltage6 < replymax.Voltage6)) ? 1 : 2;
            Test_Resault.Voltage8_OK_NOT = ((reply.Voltage7 > replymin.Voltage7) && (reply.Voltage7 < replymax.Voltage7)) ? 1 : 2;
            Test_Resault.Voltage9_OK_NOT = ((reply.Voltage8 > replymin.Voltage8) && (reply.Voltage8 < replymax.Voltage8)) ? 1 : 2;
            

            //Test_Resault.WIFI_OK_NOT = ((reply.TestResult0 & 0xff) == 0x01) ? 1 : 0;
            Test_Resault.CAN_OK_NOT = ((reply.TestResult0 & 0xff) == 0x02) ? 1 : 0;
            //Test_Resault.SWITCH_SCREEN_OK_NOT = ((reply.TestResult0 & 0xff) == 0x04) ? 1 : 0;
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
            public int ACK_CMD_LOOK_FOR_COM;
            public int ACK_DEVICE_FILE_Write;
            public int ACK_DEVICE_FILE_MORE_Write;
        } ;

        public ACKINFO ACKinfo;

        public Thread Threa_V_I_Test;
        public Thread Threa_WIFI_Test;

        private void FRM_load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            Array.Sort(ports);
            timer1.Start();

            string[] ss = MulGetHardwareInfo(HardwareEnum.Win32_PnPEntity, "Name");


            btn_beep.Enabled = false;
            btn_submit.Enabled = false;
            btn_switch_video.Enabled = false;
            btn_Test_self.Enabled = false;
            btn_search_wifi.Enabled = false;
            btn_test_wifi.Enabled = false;
            //btn_Write_MCUBIN.Enabled = false;

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
            ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;
            ACKinfo.ACK_DEVICE_FILE_MORE_Write = 0;
            ACKinfo.ACK_DEVICE_FILE_Write = 0;

            try
            {
                //打开测试板串口
                string port = null;
                foreach (string a in ss)
                {
                    if (a.Contains("STMicroelectronics Virtual COM Port"))
                    {
                        port = a.Substring(a.IndexOf("(COM") + 1);
                        port = port.Substring(0, port.IndexOf(")"));
                    }
                }
                comm = new SerialPort();
                openCOM(port);
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
                    textBox1.Text = "\r\n请检查PC主机与测试板之间的连接，并保证测试板已经上电!\r\n\r\n\r\n请重启软件！";
                    btn_Write_MCUBIN.Enabled = false;
                    textBox2.Enabled = false;
                }
                else
                {
                    string[] port1 = null;
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

                    commWifiTest = new SerialPort();
                    foreach (string portnow in port1)
                    {
                        //textBox1.Text += "\r\n" + port;
                        if (!commWifiTest.IsOpen)
                            openWifiTestCOM(portnow);
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


                    tlv_reply = new tlv_reply_t();
                    Test_Resault = new test_resault();


                    tlv_reply_min.Voltage0 = 316;    //3.3V
                    tlv_reply_min.Voltage1 = 1140;   //12V
                    tlv_reply_min.Voltage2 = 475;    //5V
                    tlv_reply_min.Voltage3 = 238;    //2.5
                    tlv_reply_min.Voltage4 = 171;    //1.8
                    tlv_reply_min.Voltage5 = 114;    //IN_1.2V
                    tlv_reply_min.Voltage6 = 316;    //IN_POWER1
                    tlv_reply_min.Voltage7 = 316;    //IN_POWER2
                    tlv_reply_min.Voltage8 = 316;    //IN_POWER3
                    tlv_reply_min.Current0 = 10;    //workCurrent
                    tlv_reply_min.Current1 = 10;    //staticCurrent

                    tlv_reply_max.Voltage0 = 350;    //3.3V
                    tlv_reply_max.Voltage1 = 1260;   //12V
                    tlv_reply_max.Voltage2 = 525;    //5V
                    tlv_reply_max.Voltage3 = 350;    //2.5V
                    tlv_reply_max.Voltage4 = 189;    //1.8V
                    tlv_reply_max.Voltage5 = 126;    //1.2V
                    tlv_reply_max.Voltage6 = 350;    //POWER1
                    tlv_reply_max.Voltage7 = 350;    //POWER2
                    tlv_reply_max.Voltage8 = 350;    //POWER3
                    tlv_reply_max.Current0 = 20;
                    tlv_reply_max.Current1 = 20;


                    Cdirectory = System.IO.Directory.GetCurrentDirectory();
                    xApp = new Microsoft.Office.Interop.Excel.Application();
                    xApp.DisplayAlerts = false;

                    if (File.Exists(Cdirectory + "\\主板测试结果.xlsx"))
                    {
                        wbs = xApp.Workbooks;
                        wb = wbs.Add(Cdirectory + "\\主板测试结果.xlsx");
                        ws = wb.Worksheets["主板测试结果"];
                    }
                    else
                    {
                        wbs = xApp.Workbooks;
                        wb = wbs.Add(true);
                        ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        ws.Name = "主板测试结果";
                        for (int i = startColumn; i <= endColumn; i++)
                        {
                            ws.Cells[1, i] = columnName[i - 1];
                        }
                    }

                    textBox1.Text = "请扫码\r\n" + textBox1.Text;
                    
                }
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }
            finally
            {
                checkBox_Switch_ok.Enabled = false; checkBox_Switch_failed.Enabled = false;
                checkBox_beep_ok.Enabled = false; checkBox_beep_failed.Enabled = false;
            }

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void comboBox1_DragDrop(object sender, DragEventArgs e)
        {
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_static_Curent_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_7_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_7_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_8_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Curent_8_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void comboBox_selectComPort_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = "0000000001";
        }
        

        private void btn_openCOM_Click(object sender, EventArgs e)
        {
            
        }


        public delegate int MyDelegate();
        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }
        private void btn_Write_MCUBIN_Click(object sender, EventArgs e)
        {
            //开始测试    线程
            //Thread startTestThred = new Thread(startTest);
            //startTestThred.Start();

            //下载MCU  BIN文件  线程

            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }
            btn_Write_MCUBIN.Enabled = false;
            textBox2.Enabled = false;
            checkBox_Switch_ok.Enabled = false; checkBox_Switch_failed.Enabled = false;
            checkBox_beep_ok.Enabled = false;checkBox_beep_failed.Enabled = false;
            BarNum = textBox2.Text.ToString();
            Thread DownLoadMCUBINThred = new Thread(downLoadBIN);
            DownLoadMCUBINThred.Start();

            //超时等待  线程
            //Thread WaitDownloadACKThred = new Thread(waitDownloadACK);
            //WaitDownloadACKThred.Start();
        }

        private void btn_Test_self_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }

            //开始自测   线程
            Thread selfTestThred = new Thread(selfTest);
            selfTestThred.Start();
        }
        
        public byte OP_BEEP = 0x00;
        public byte OP_CAR_Video = 0x01;
        public byte OP_Mobileye_Video = 0x02;
        public byte OP_DVR_Video = 0x03;
        public byte OP_CAN_Test_ok = 0x04;

        public byte video_source = 0x01;
        private void btn_switch_video_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }

            checkBox_Switch_ok.Enabled = true;checkBox_Switch_failed.Enabled = true;
            sendCMD(SERVICE_TYPE_CMD , MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ , video_source);
            textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
            textBox1.ForeColor = Color.Black;
            textBox1.Text = "切换视频\r\n" + textBox1.Text;
            //发送指令
            if (video_source == OP_DVR_Video)
                video_source = OP_CAR_Video;
            else
                video_source++;
        }


        public void sendCMD(byte SERVICE_TYPE , byte MSG_TYPE , byte OP_CODE)
        {
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE;
            msg_head.msg_type = MSG_TYPE;
            msg_head.msg_len = 16;
            msg_head.crc32 = 0;
            msg_head.msg_count = 0;
            msg_head.resp = OP_CODE;
            msg_head.seq = 0;

            byte[] tmpbuff = StructToBytes(msg_head);
            comm.Write(tmpbuff, 0, msg_head.msg_len);
        }

        private void btn_beep_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }
            checkBox_beep_ok.Enabled = true;checkBox_beep_failed.Enabled = true;
            //发送指令  
            sendCMD(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, 0x00);
            textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
            textBox1.ForeColor = Color.Black;
            textBox1.Text = "响蜂鸣器\r\n" + textBox1.Text;
        }

        Microsoft.Office.Interop.Excel.Range ran;
        private void btn_submit_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }
            int blankRow = 0;
            xApp.DisplayAlerts = false;
            //判断是否已手动确认蜂鸣器 和 视频切换
            if ((checkBox_CAN_ok.Checked != true) && (checkBox_CAN_failed.Checked != true))
            {
                MessageBox.Show("尚未完成软件自测，请进行软件自测");
                return;
            }
            if ((checkBox_Switch_ok.Checked != true) && (checkBox_Switch_failed.Checked != true))
            {
                MessageBox.Show("请手动确认切换视频功能是否正常！如果正常勾选正常，否则勾选不正常！");
                return;
            }
            if ((checkBox_beep_ok.Checked != true) && (checkBox_beep_failed.Checked != true))
            {
                MessageBox.Show("请手动确认蜂鸣器是否正常！如果正常否选正常，否则勾选不正常！");
                return;
            }

            try
            {
                //保存数据
                for (int i = 1; i < 65536; i++)
                {
                    Microsoft.Office.Interop.Excel.Range ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, 1];
                    if (ran.Value == null)
                    {
                        blankRow = i;
                        break;
                    }
                }
                ws.Cells[blankRow, (int)column.time] = DateTime.Now.ToLocalTime().ToString();// time;
                ws.Cells[blankRow, (int)column.BarNum] = textBox2.Text.ToString();   //条码

                ws.Cells[blankRow, (int)column._12_V] = tlv_reply.Voltage1.ToString();
                ws.Cells[blankRow, (int)column._5_V] = tlv_reply.Voltage2.ToString();
                ws.Cells[blankRow, (int)column._3_3_V] = tlv_reply.Voltage0.ToString();
                ws.Cells[blankRow, (int)column._2_5_V] = tlv_reply.Voltage3.ToString();
                ws.Cells[blankRow, (int)column._1_8_V] = tlv_reply.Voltage4.ToString();
                ws.Cells[blankRow, (int)column._1_2_V] = tlv_reply.Voltage5.ToString();
                ws.Cells[blankRow, (int)column.INPOWER1_V] = tlv_reply.Voltage6.ToString();
                ws.Cells[blankRow, (int)column.INPOWER2_V] = tlv_reply.Voltage7.ToString();
                ws.Cells[blankRow, (int)column.INPOWER3_V] = tlv_reply.Voltage8.ToString();

                ws.Cells[blankRow, (int)column.staticcurrent] = tlv_reply.Current1.ToString();
                ws.Cells[blankRow, (int)column.workcurrent] = tlv_reply.Current0.ToString();

                if (checkBox_CAN_ok.Checked == true)
                    ws.Cells[blankRow, (int)column.CAN] = "OK";
                else
                    ws.Cells[blankRow, (int)column.CAN] = "NO";
                if (checkBox_beep_ok.Enabled == true)
                    ws.Cells[blankRow, (int)column.BEEP] = "OK";
                else
                    ws.Cells[blankRow, (int)column.BEEP] = "NO";
                if (checkBox_Switch_ok.Enabled == true)
                    ws.Cells[blankRow, (int)column.SWITCH_VIDEO] = "OK";
                else
                    ws.Cells[blankRow, (int)column.SWITCH_VIDEO] = "NO";


                if (textBox_Voltage_12V.BackColor == Color.Red)   //12V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._12_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._12_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_5V.BackColor == Color.Red)   //5V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._5_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._5_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_3_3V.BackColor == Color.Red)   //3.3V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._3_3_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._3_3_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_2_5V.BackColor == Color.Red)   //2.5V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._2_5_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._2_5_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_1_8V.BackColor == Color.Red)   //1.8V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._1_8_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._1_8_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_1_2V.BackColor == Color.Red)   //1.2V
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._1_2_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column._1_2_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_INPOWER1.BackColor == Color.Red)   //INPOWER1
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER1_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER1_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_INPOWER2.BackColor == Color.Red)   //INPOWER2
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER2_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER2_V];
                    ran.Interior.Color = Color.Green;
                }

                if (textBox_Voltage_INPOWER3.BackColor == Color.Red)   //INPOWER3
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER3_V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.INPOWER3_V];
                    ran.Interior.Color = Color.Green;
                }


                xApp.DisplayAlerts = false;
                ws.SaveAs("主板测试结果");
                if (wb != null)
                {
                    wb.SaveAs(Cdirectory + "\\主板测试结果.xlsx");
                    wb.Saved = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                textBox1.Text = "关闭EXCEL应用！\r\n" + textBox1.Text;
                MessageBox.Show("请重新测试！！！\r\n\r\n\r\n 即将关闭所有EXCEL文档！！！");
                this.Close();
            }
            finally
            {
                btn_beep.Enabled = false;
                btn_submit.Enabled = false;
                btn_switch_video.Enabled = false;
                btn_Test_self.Enabled = false;
                btn_search_wifi.Enabled = true;
                btn_test_wifi.Enabled = true;
                btn_Write_MCUBIN.Enabled = true;

                textBox2.Text = "";

                checkBox_CAN_ok.Checked = false;checkBox_CAN_failed.Checked = false; checkBox_CAN_ok.Enabled = false;checkBox_CAN_failed.Enabled = false;
                checkBox_beep_ok.Checked = false;checkBox_beep_failed.Checked = false;checkBox_beep_ok.Enabled = false;checkBox_beep_failed.Enabled = false;
                checkBox_Switch_ok.Checked = false;checkBox_Switch_failed.Checked = false;
                textBox_Voltage_12V.BackColor = Color.White;textBox_Voltage_12V.Text = null;
                textBox_Voltage_5V.BackColor = Color.White; textBox_Voltage_5V.Text = null;
                textBox_Voltage_3_3V.BackColor = Color.White; textBox_Voltage_3_3V.Text = null;
                textBox_Voltage_2_5V.BackColor = Color.White; textBox_Voltage_2_5V.Text = null;
                textBox_Voltage_1_8V.BackColor = Color.White; textBox_Voltage_1_8V.Text = null;
                textBox_Voltage_1_2V.BackColor = Color.White; textBox_Voltage_1_2V.Text = null;
                textBox_Voltage_INPOWER1.BackColor = Color.White; textBox_Voltage_INPOWER1.Text = null;
                textBox_Voltage_INPOWER2.BackColor = Color.White; textBox_Voltage_INPOWER2.Text = null;
                textBox_Voltage_INPOWER3.BackColor = Color.White; textBox_Voltage_INPOWER3.Text = null;

                textBox_static_Curent.Text = null;
                textBox_work_Curent.Text = null;

                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Black;
                    textBox1.Text = null;

                    textBox2.Enabled = true;
                    textBox1.Text = "请扫码";

                    textBox_WIFI_FUC.Text = "";
                    textBox_WIFI_FUC.BackColor = Color.White;

                    textBox_rssi.Text = "";
                    textBox_rssi.BackColor = Color.White;
                }));
                //btn_Write_MCUBIN.Enabled = false;
                progressBar1.Value = 0;
            }
        }

        private void MainBoard_PCBATest_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (wb != null)
                {
                    wb.Close(Type.Missing, Type.Missing, Type.Missing);
                    wbs.Close();
                    xApp.Quit();
                    wb = null;
                    wbs = null;
                    xApp = null;
                }

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
            catch (Exception ex)
            {

            }
            
        }

        private void MainBoard_PCBATest_Leave(object sender, EventArgs e)
        {
            xApp.DisplayAlerts = false;
            xApp.Quit();
            wb = null;
            wbs = null;
            xApp = null;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //btn_switch_video.Enabled = true;
            //btn_beep.Enabled = true;


            for (int i = 0; i < (int)timeF.TotalTimer; i++)
                if (TimerArry[i] > 0)
                    TimerArry[i]--;
            if (((checkBox_CAN_ok.Checked == true) || (checkBox_CAN_failed.Checked == true)) && ((checkBox_beep_ok.Checked == true) || (checkBox_beep_failed.Checked == true)) && ((checkBox_Switch_ok.Checked == true) || (checkBox_Switch_failed.Checked == true)))
                btn_submit.Enabled = true;
            else
                btn_submit.Enabled = false;


            if (((checkBox_beep_ok.Checked == true) && (checkBox_beep_failed.Checked == true)) || ((checkBox_Switch_ok.Checked == true) && (checkBox_Switch_failed.Checked == true)))
            {

            }
            if ((Test_Resault.CAN_OK_NOT == 1) &&
                //(Test_Resault.WIFI_OK_NOT == 1) &&
                //(Test_Resault.SWITCH_SCREEN_OK_NOT == 1) &&
                (Test_Resault.Voltage1_OK_NOT == 1) &&
                (Test_Resault.Voltage2_OK_NOT == 1) &&
                (Test_Resault.Voltage3_OK_NOT == 1) &&
                (Test_Resault.Voltage4_OK_NOT == 1) &&
                (Test_Resault.Voltage5_OK_NOT == 1) &&
                (Test_Resault.Voltage6_OK_NOT == 1) &&
                (Test_Resault.Voltage7_OK_NOT == 1) &&
                (Test_Resault.Voltage8_OK_NOT == 1) &&
                (Test_Resault.Voltage9_OK_NOT == 1) &&
                (Test_Resault.Work_Curent_OK_NOT == 1) &&
                (Test_Resault.Static_Curent_OK_NOT == 1)&&
                (checkBox_beep_ok.Checked==true)&&
                (checkBox_Switch_ok.Checked==true)
                )
                PASS_FLAG = 1;
            else if ((Test_Resault.CAN_OK_NOT == 1) &&
                //(Test_Resault.WIFI_OK_NOT == 1) &&
                //(Test_Resault.SWITCH_SCREEN_OK_NOT == 1) &&
                (Test_Resault.Voltage1_OK_NOT == 1) &&
                (Test_Resault.Voltage2_OK_NOT == 1) &&
                (Test_Resault.Voltage3_OK_NOT == 1) &&
                (Test_Resault.Voltage4_OK_NOT == 1) &&
                (Test_Resault.Voltage5_OK_NOT == 1) &&
                (Test_Resault.Voltage6_OK_NOT == 1) &&
                (Test_Resault.Voltage7_OK_NOT == 1) &&
                (Test_Resault.Voltage8_OK_NOT == 1) &&
                (Test_Resault.Voltage9_OK_NOT == 1) &&
                (Test_Resault.Work_Curent_OK_NOT == 1) &&
                (Test_Resault.Static_Curent_OK_NOT == 1) &&
                ((checkBox_beep_failed.Checked == true) ||
                (checkBox_Switch_failed.Checked == true))
                )
                PASS_FLAG = -1;
            else
                PASS_FLAG = 0;


            if (PASS_FLAG == 1)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Green;
                    textBox1.Text = "PASS";
                }));
                return;
            }
            else if (PASS_FLAG == -1)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Red;
                    textBox1.Text = "NO";
                }));
                return;
            }
            else
            {

            }

            if ((textBox2.Text == "") || (textBox2.Text == null) || (textBox2.Enabled == false))
                btn_Write_MCUBIN.Enabled = false;
            else
                btn_Write_MCUBIN.Enabled = true;

            if (((checkBox_beep_ok.Checked == true) || (checkBox_beep_failed.Checked == true)) && ((checkBox_Switch_ok.Checked = true) || (checkBox_Switch_failed.Checked = true)))
            {
                btn_submit.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DeviceERest();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Invoke((EventHandler)(delegate { progressBar1.Minimum = 0; }));
            this.Invoke((EventHandler)(delegate { progressBar1.Value = 0; }));
            long lSize = new FileInfo("FPGA_01.BIN").Length;
            if (lSize % 1024 == 0)
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024); }));
            else
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024) + 1; }));
            lSize = new FileInfo("PIC_20180322.BIN").Length;
            if (lSize % 1024 == 0)
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024); }));
            else
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024) + 1; }));
            DownLoadFileToMCU("PIC_20180322.BIN");
            Thread.Sleep(5000);
            DownLoadFileToMCU("FPGA_01.BIN");
        }

        private void checkBox_Switch_ok_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Switch_ok.Checked == true)
            {
                if (checkBox_Switch_failed.Checked == true)
                {
                    checkBox_Switch_ok.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void checkBox_Switch_failed_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Switch_failed.Checked == true)
            {
                if (checkBox_Switch_ok.Checked == true)
                {
                    checkBox_Switch_failed.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void checkBox_beep_ok_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_beep_ok.Checked == true)
            {
                if (checkBox_beep_failed.Checked == true)
                {
                    checkBox_beep_ok.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void checkBox_beep_failed_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_beep_failed.Checked == true)
            {
                if (checkBox_beep_ok.Checked == true)
                {
                    checkBox_beep_failed.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DeviceERest();
        }

        private void btn_search_wifi_Click(object sender, EventArgs e)
        {
            CommWifiTestBuffWriteIndex = 0;
            CommWifiTestBuffReadIndex = 0;
            Thread look_for_wifi_thread = new Thread(LOOK_FOR_WIFI);
            look_for_wifi_thread.Start();
        }

        private void btn_connect_wifi_Click(object sender, EventArgs e)
        {
            
        }

        private void btn_test_wifi_Click(object sender, EventArgs e)
        {
            if (WIFITest_RUN == 0)
            {
                WIFITest_RUN = 1;
            }
            else
            {
                WIFITest_RUN = 0;
            }
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
                        if (ADASLeader_device[i].is_pass == "是")
                            item.BackColor = Color.Green;
                        else
                            item.BackColor = Color.Red;
                    }
                }
            }
        }
    }
}

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
using System.Management;
namespace ADASLeader产品生产测试
{
    public partial class CameraMoudle_EndProduct_PCBATest : Form
    {
        public int PASS_FLAG; //1：测试通过，其他：测试未通过
        
        public void Anisy_Reply(int rows, tlv_reply_t reply, tlv_reply_t replymax, tlv_reply_t replymin)  //如果reply1中
        {
            // BEEP
            Microsoft.Office.Interop.Excel.Range ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rows, (int)column.BEEP];
            if ((reply.Voltage0 < replymin.Current0) | (reply.Voltage0 > replymax.Current0))
                ran.Interior.ColorIndex = 3;   //红色
            // SWITCH
            ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rows, (int)column.SWITCH];
            if ((reply.Voltage0 < replymin.Current0) | (reply.Voltage0 > replymax.Current0))
                ran.Interior.ColorIndex = 3;   //红色

        }

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

        public struct TEST_RESAULT
        {
            public byte SWITCH_VIDEO;
            public byte BEEP;
            public byte TEST_FINISHED;  //bit 0 : switch viceo , bit 1 : beep 
        };
        TEST_RESAULT test_resault;

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
        };
        public ACKINFO ACKinfo;

        public string mFilename;
        private string BarNum;
        public byte[] commBuf;
        public SerialPort comm;
        public enum timeF
        {
            selfTest = 0,
            BEEP = 1,
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
            BEEP = 3,
            SWITCH = 4,
        }

        string[] columnName = { "时间" , "条形码", "蜂鸣器", "视频切换" };

        public int startColumn = (int)column.time;
        public int endColumn = (int)column.SWITCH;


        public CameraMoudle_EndProduct_PCBATest(string barNum)
        {
            InitializeComponent();
            commBuf = new byte[4096];
            BarNum = barNum;
            timer1.Interval = 1000;
            timer1.Start();
        }


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
            string[] ports = SerialPort.GetPortNames();
            Array.Sort(ports);


            string[] ss = MulGetHardwareInfo(HardwareEnum.Win32_PnPEntity, "Name");


            btn_submit.Enabled = false;
            btn_Beep.Enabled = true;

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

            

            tlv_reply = new tlv_reply_t();

            tlv_reply_min.Voltage4 = (UInt16)(180 * 0.95);     //1.8V
            tlv_reply_min.Voltage2 = (UInt16)(500 * 0.95);    //5V
            tlv_reply_min.Voltage6 = (UInt16)(350 * 0.95);    //3.5V
            tlv_reply_min.Voltage7 = (UInt16)(330 * 0.95);    //3.3V

            tlv_reply_max.Voltage4 = (UInt16)(180 * 1.05);
            tlv_reply_max.Voltage2 = (UInt16)(500 * 1.05);
            tlv_reply_max.Voltage6 = (UInt16)(350 * 1.05);
            tlv_reply_max.Voltage7 = (UInt16)(330 * 1.05);



            try
            {
                Cdirectory = System.IO.Directory.GetCurrentDirectory();
                xApp = new Microsoft.Office.Interop.Excel.Application();
                xApp.DisplayAlerts = false;

                if (File.Exists(Cdirectory + "\\摄像头模块测试结果.xlsx"))
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(Cdirectory + "\\摄像头模块测试结果.xlsx");
                    ws = wb.Worksheets["摄像头模块测试结果"];
                }
                else
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(true);
                    ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = "摄像头模块测试结果";
                    for (int i = startColumn; i <= endColumn; i++)
                    {
                        ws.Cells[1, i] = columnName[i - 1];
                    }
                    ws.SaveAs("摄像头模块测试结果");
                    if (wb != null)
                    {
                        wb.SaveAs(Cdirectory + "\\摄像头模块测试结果.xlsx");
                        wb.Saved = true;
                    }
                }

                Thread.Sleep(1000);
                EXCEL_PROCESS = Process.GetProcesses();

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
                for (int i = 0; i < 10; i++)
                {
                    Thread.Sleep(500);
                    if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)  //open com success ，no more wait
                        break;
                }

                if (ACKinfo.ACK_CMD_LOOK_FOR_COM == 1)      //open com success ，no more looking for
                {
                    ACKinfo.ACK_CMD_LOOK_FOR_COM = 0;
                    textBox1.Text = "OPEN TEST BOARD COM SUCCESS";
                }
                else                                        //no right com
                {
                    comm.Close();
                }
                if (!comm.IsOpen)
                {
                    textBox1.Text += "\r\n请检查PC主机与测试板之间的连接，并保证测试板已经上电!\r\n\r\n\r\n请重启软件！";
                }
                else
                    textBox1.Text += "\r\n请扫码";
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);


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
                return;
            }
            finally
            {
                checkBox_Switch_ok.Enabled = false; checkBox_Switch_failed.Enabled = false;
                checkBox_beep_ok.Enabled = false; checkBox_beep_failed.Enabled = false;
                btn_Beep.Enabled = false;
                btn_Switch_Screen.Enabled = false;
                btn_submit.Enabled = false;
            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        /*
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
                btn_Switch_Screen.Enabled = true;

                //建立串口接收线程
            }
            catch (Exception ex)
            {
                //捕获到异常信息，创建一个新的comm对象，之前的不能用了。
                comm = new SerialPort();
                //现实异常信息给客户。
                textBox1.Text += "\r\n" + ex.Message;
            }
        }*/
        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }
        public byte DVR_VIDEO = 1;
        public byte CAR_VIDEO = 2;
        public byte CAMERA_VIDEO = 3;
        public byte BEEP = 4;

        public void sendCMD(byte SERVICE_TYPE, byte MSG_TYPE, byte OP_CODE)
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

        public byte OP_BEEP = 0x00;
        public byte OP_CAR_Video = 0x01;
        public byte OP_Mobileye_Video = 0x02;
        public byte OP_DVR_Video = 0x03;
        public byte OP_CAN_Test_ok = 0x04;

        public byte video_source = 0x01;
        private void btn_Switch_Screen_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }

            checkBox_Switch_ok.Enabled = true; checkBox_Switch_failed.Enabled = true;
            sendCMD(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, video_source);
            if(video_source == OP_CAR_Video)
                textBox1.Text = "车机视频\r\n" + textBox1.Text;
            if (video_source == OP_Mobileye_Video)
                textBox1.Text = "天眼视频\r\n" + textBox1.Text;
            if (video_source == OP_DVR_Video)
                textBox1.Text = "DVR视频\r\n" + textBox1.Text;
            //发送指令
            if (video_source == OP_DVR_Video)
                video_source = OP_CAR_Video;
            else
                video_source++;
        }

        private void btn_Test_self_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请扫码");
                return;
            }
            checkBox_beep_ok.Enabled = true; checkBox_beep_failed.Enabled = true;
            //发送指令  
            sendCMD(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, 0x00);
            textBox1.Text = "响蜂鸣器\r\n" + textBox1.Text;
        }

        Microsoft.Office.Interop.Excel.Range ran;
        private void btn_submit_Click(object sender, EventArgs e)
        {
            int blankRow = 0;
            xApp.DisplayAlerts = false;
            //判断是否已手动确认蜂鸣器 和 视频切换
            if ((checkBox_Switch_ok.Checked != true) && (checkBox_Switch_failed.Checked != true) || (checkBox_Switch_ok.Checked == true) && (checkBox_Switch_failed.Checked == true))
            {
                MessageBox.Show("请手动确认切换视频功能是否正常！如果正常勾选正常，否则勾选不正常！");
                return;
            }
            if ((checkBox_beep_ok.Checked != true) && (checkBox_beep_failed.Checked != true) || (checkBox_beep_ok.Checked == true) && (checkBox_beep_failed.Checked == true))
            {
                MessageBox.Show("请手动确认蜂鸣器是否正常！如果正常否选正常，否则勾选不正常！");
                return;
            }
            
            try
            {
                //保存数据
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
                ws.Cells[blankRow, (int)column.BarNum] = textBox2.Text.ToString();   //条码
                if (checkBox_Switch_ok.Checked == true)    //switch
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.SWITCH];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.SWITCH] = "OK";
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.SWITCH];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.SWITCH] = "N0";
                }

                if (checkBox_beep_ok.Checked == true)  //BEEP
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.BEEP];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.BEEP] = "OK";
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.BEEP];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.BEEP] = "N0";
                }

                xApp.DisplayAlerts = false;
                ws.SaveAs("摄像头模块测试结果");
                if (wb != null)
                {
                    //wb.Save();
                    wb.SaveAs(Cdirectory + "\\摄像头模块测试结果.xlsx");
                    wb.Saved = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("请重新测试！！！\r\n\r\n\r\n 即将关闭所有EXCEL文档！！！" , ex.Message);
            }
            finally
            {
                btn_submit.Enabled = false;
                btn_Switch_Screen.Enabled = false;
                btn_Beep.Enabled = false;
                textBox2.Enabled = true;textBox2.Text = null;

                checkBox_beep_ok.Enabled = false;checkBox_beep_failed.Enabled = false;checkBox_beep_ok.Checked = false;checkBox_beep_failed.Checked = false;
                checkBox_Switch_ok.Enabled = false;checkBox_Switch_failed.Enabled = false;checkBox_Switch_ok.Checked = false;checkBox_Switch_failed.Checked = false;
                textBox1.Text = "测试结果保存成功!\r\n请扫码";
            }
            return;
        }

        public Process[] EXCEL_PROCESS;
        private void CameraMoudle_EndProduct_PCBATest_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {

                foreach (Process A in EXCEL_PROCESS)
                {
                    if (A.ProcessName.ToString() == "EXCEL")
                        A.Kill();
                    //textBox1.Text += "\r\n" + A.ProcessName.ToString();
                }

                if (comm.IsOpen) { comm.Close(); };

                //this.Close();

            }
            catch (StackOverflowException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (((checkBox_beep_failed.Checked != false) || (checkBox_beep_ok.Checked != false)) && ((checkBox_Switch_failed.Checked != false) || (checkBox_Switch_ok.Checked != false)))
                btn_submit.Enabled = true;
            else
                btn_submit.Enabled = false;

            if ((textBox2.Text != null) && (IsNumber(textBox2.Text.ToString())))
            {
                btn_Beep.Enabled = true;
                btn_Switch_Screen.Enabled = true;
            }
            else
            {
                btn_Beep.Enabled = false;
                btn_Switch_Screen.Enabled = false;
                btn_submit.Enabled = false;
            }
        }
    }
}

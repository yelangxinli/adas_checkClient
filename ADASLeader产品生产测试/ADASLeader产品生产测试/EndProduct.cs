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
    public partial class EndProduct : Form
    {
        public EndProduct(string option)    //"end product"——成品 ， "mainboard moudle"——主板模块
        {
            InitializeComponent();

            commBuf = new byte[4096];
            commWifiTestBuf = new byte[4096 * 10];
            testresault.data = new byte[0x15];
            testresault.returnF = 0;

            Threa_V_I_Test = new Thread(selfTest);
            timer1.Interval = 1000;
            timer1.Start();
            timer2.Interval = 1000;
            timer2.Start();

            if (option == "end product")
            {
                this.Text = "ADASLeader成品测试";
                saveFileName = "成品测试结果";
            }
            else if (option == "mainboard moudle")
            {
                this.Text = "ADASLeader主机模块测试";
                saveFileName = "主板模块测试结果";
            }
            else
            {

            }
        }
        Thread Threa_V_I_Test;
        public string saveFileName = null;
        public struct TEST_RESAULT
        {
            public byte SWITCH_VIDEO;
            public byte WIFI_FUC;
            public byte WIFI_RSSI;
            public byte CAN_SENSOR;
            public byte TEST_FINISHED;  //bit 0 : wifi test , bit 1 : can test  , bit4 : 车机视频  , bit5 : 摄像头视频 , bit6 : DVR视频
        };
        TEST_RESAULT test_resault;

        public int PASS_FLAG; //1：测试通过，其他：测试未通过

        public string mFilename;
        private string BarNum;
        public byte[] commBuf;
        public byte[] commWifiTestBuf;
        public int mainboardBinWrite = 0;
        public SerialPort comm;
        public SerialPort commWifiTest;
        public int commWifiTest_returnF;
        public enum timeF
        {
            selfTest = 0,
            lookforwifi = 1,
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
            wifiName = 3,
            switch_video = 4,
            WIFI_FUC = 5,
            WIFI_RSSI = 6,
        }

        public enum Test_opration
        {
            beep = 0,
            Car_Video = 1,
            Mobileye_Video = 2,
            DVR_Video = 3,
            CAN_Test = 4,
        }

        string[] columnName = { "时间", "条形码", "WIFI名称", "视频切换", "WIFI功能", "WIFI RSSI" };

        public int startColumn = (int)column.time;
        public int endColumn = (int)column.WIFI_RSSI;






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
            public int ACK_CMD_CAN_SENSOR_TEST;
        };

        public ACKINFO ACKinfo;

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
        tlv_reply_t tlv_reply;
        Microsoft.Office.Interop.Excel.Range ran;

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
            string[] port1;
            Array.Sort(ports);
            timer1.Start();

            string[] ss = MulGetHardwareInfo(HardwareEnum.Win32_PnPEntity, "Name");

            //btn_beep.Enabled = false;
            //btn_submit.Enabled = false;
            //btn_switch_video.Enabled = false;
            //btn_Test_self.Enabled = false;
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
                    if (a.Contains("XDS110 Class Application/User UART"))
                    {
                        port = a.Substring(a.IndexOf("(COM") + 1);
                        port = port.Substring(0, port.IndexOf(")"));
                        break;
                    }
                }
                commWifiTest = new SerialPort();
                if ((!commWifiTest.IsOpen) && (port != null))
                    openWifiTestCOM(port);


                Cdirectory = System.IO.Directory.GetCurrentDirectory();
                xApp = new Microsoft.Office.Interop.Excel.Application();
                xApp.DisplayAlerts = false;

                if (File.Exists(Cdirectory + "\\" + saveFileName+ ".xlsx"))
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(Cdirectory + "\\" + saveFileName + ".xlsx");
                    ws = wb.Worksheets["saveFileName"];
                }
                else
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(true);
                    ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = "saveFileName";
                    for (int i = startColumn; i <= endColumn; i++)
                    {
                        ws.Cells[1, i] = columnName[i - 1];
                    }
                }

                //Thread CAN_SENOR_TEST = new Thread(can_sendor_test);
                //CAN_SENOR_TEST.Start();
            }
            catch (Exception ex2)
            {
                MessageBox.Show(ex2.Message);
            }

            Threa_V_I_Test.Start();

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button6.Enabled = false;
            btn_connect.Enabled = false;
        }

        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }


        public int WIFITest_RUN = 0;

        public byte OP_BEEP = 0x00;
        public byte OP_CAR_Video = 0x01;
        public byte OP_Mobileye_Video = 0x02;
        public byte OP_DVR_Video = 0x03;
        public byte OP_CAN_Test_ok = 0x04;
        private void btn_search_wifi_Click(object sender, EventArgs e)
        {
            CommWifiTestBuffWriteIndex = 0;
            CommWifiTestBuffReadIndex = 0;
            Thread look_for_wifi_thread = new Thread(LOOK_FOR_WIFI);
            look_for_wifi_thread.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }
            byte[] data = { 0x0C, 0x00, 0x06, 0x00, 0x01, 0x00 };
            sendCMDToWIFI(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, data, (UInt16)data.Length);
            test_resault.TEST_FINISHED = (byte)(test_resault.TEST_FINISHED | (byte)0x10);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }
            byte[] data = { 0x0C, 0x00, 0x06, 0x00, 0x03, 0x00 };
            sendCMDToWIFI(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, data, (UInt16)data.Length);
            test_resault.TEST_FINISHED = (byte)(test_resault.TEST_FINISHED | (byte)0x40);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }
            byte[] data = { 0x0C, 0x00, 0x06, 0x00, 0x02, 0x00 };
            sendCMDToWIFI(SERVICE_TYPE_CMD, MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ, data, (UInt16)data.Length);
            test_resault.TEST_FINISHED = (byte)(test_resault.TEST_FINISHED | (byte)0x20);
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            TimerArry[(int)timeF.selfTest] = TimerArry[(int)timeF.selfTest] - 1;
            TimerArry[(int)timeF.lookforwifi] = TimerArry[(int)timeF.lookforwifi] - 1;

            if ((test_resault.TEST_FINISHED & 0x70) == 0x70)
            {
                button6.Enabled = true;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            heartBeat = 0;
        }

        private void btn_connect_Click(object sender, EventArgs e)
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
        public string wifi_name = null;

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
                ws.Cells[blankRow, (int)column.BarNum] = textBox2.Text.ToString();   // barNum
                ws.Cells[blankRow, (int)column.wifiName] = wifi_name;   // barNum
                if ((checkBox2.Checked == true) && (checkBox1.Checked == false))     //切换视频
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.switch_video];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.switch_video] = "N0";
                }
                else if ((checkBox2.Checked == false) && (checkBox1.Checked == true))
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.switch_video];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.switch_video] = "OK";
                }
                else
                {
                    MessageBox.Show("请正确勾选测试结果");
                    return;
                }
                if (((test_resault.TEST_FINISHED & 0x01) != 0) && (test_resault.WIFI_FUC == 0))     //WIFI 功能
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.WIFI_FUC];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.WIFI_FUC] = "N0";
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.WIFI_FUC];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.WIFI_FUC] = "OK";
                }
                if (((test_resault.TEST_FINISHED & 0x01) != 0) && (test_resault.WIFI_RSSI == 0))     //WIFI RISSI
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.WIFI_RSSI];
                    ran.Interior.Color = Color.Red;
                    ws.Cells[blankRow, (int)column.WIFI_RSSI] = "N0";
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.WIFI_RSSI];
                    ran.Interior.Color = Color.Green;
                    ws.Cells[blankRow, (int)column.WIFI_RSSI] = textBox_rssi.Text.ToString();
                }
                xApp.DisplayAlerts = false;
                ws.SaveAs("saveFileName");
                if (wb != null)
                {
                    //wb.Save();
                    wb.SaveAs(Cdirectory + "\\" + saveFileName + ".xlsx");
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
        private void button6_Click(object sender, EventArgs e)
        {
            if ((test_resault.TEST_FINISHED & 0x10) == 0)
            {
                MessageBox.Show("未测试车机视频");
                return;
            }
            if ((test_resault.TEST_FINISHED & 0x20) == 0)
            {
                MessageBox.Show("未测试摄像头视频");
                return;
            }
            if ((test_resault.TEST_FINISHED & 0x40) == 0)
            {
                MessageBox.Show("未测试DVR视频");
                return;
            }
            if ((checkBox1.Checked == false) && (checkBox2.Checked == false))
            {
                MessageBox.Show("请完勾选视频切换测试结果");
                return;
            }

            //保存测试结果
            saveTestResault();
            button6.Enabled = false;
            test_resault.TEST_FINISHED = 0;
        }

        private void EndProduct_FormClosing(object sender, FormClosingEventArgs e)
        {
            WIFITest_RUN = 0;
            Threa_V_I_Test.Abort();
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                if (checkBox2.Checked == true)
                {
                    checkBox1.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                if (checkBox1.Checked == true)
                {
                    checkBox2.Checked = false;
                    MessageBox.Show("正常 和 不正常状态不能同时被勾选");
                }
            }
        }
    }
}

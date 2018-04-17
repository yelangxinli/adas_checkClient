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
    public partial class CameraBoard_PCBATest : Form
    {
        public int PASS_FLAG; //1：测试通过，其他：测试未通过
        public SerialPort comm;
        public string mFilename;
        private string BarNum;
        public byte[] commBuf;
        public enum timeF
        {
            write_mcu = 0,
            selfTest = 1,
            TotalTimer = 16,
        };
        int[] TimerArry = new int[16];
        public int mainboardBinWrite = 0;



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
            work_current = 3,           //工作电流
            static_current = 4,           //静态电流
            voltage_5V = 5,           //5V
            voltage_1_8V = 6,         //1.8V
            voltage_3_5V = 7,           //3.5V
            voltage_3_3V = 8,           //3.3V
        }

        string[] columnName = { "时间" , "条形码", "工作电流" , "静态电流", "5V", "1.8V", "3.5V", "3.3V"};

        public int startColumn = (int)column.BarNum;
        public int endColumn = (int)column.voltage_3_3V;

        tlv_reply_t tlv_reply_min;
        tlv_reply_t tlv_reply_max;
        private void FRM_load(object sender, EventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            Array.Sort(ports);
            timer1.Interval = 1000;
            timer1.Start();


            btn_Write_MCUBIN.Enabled = false;
            btn_submit.Enabled = false;
            btn_Test_self.Enabled = false;
            btn_Write_MCUBIN.Enabled = false;

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
            Test_Resault = new test_resault();
            
            tlv_reply_min.Voltage4 = (UInt16)(180*0.95);     //1.8V
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

                if (File.Exists(Cdirectory + "\\摄像头板测试结果.xlsx"))
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(Cdirectory + "\\摄像头板测试结果.xlsx");
                    ws = wb.Worksheets["摄像头板测试结果"];
                }
                else
                {
                    wbs = xApp.Workbooks;
                    wb = wbs.Add(true);
                    ws = wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = "摄像头板测试结果";
                    for (int i = startColumn; i <= endColumn; i++)
                    {
                        ws.Cells[1, i] = columnName[i - 1];
                    }
                    ws.SaveAs("摄像头板测试结果");
                    if (wb != null)
                    {
                        wb.SaveAs(Cdirectory + "\\摄像头板测试结果.xlsx");
                        wb.Saved = true;
                    }
                }

                Thread.Sleep(1000);
                EXCEL_PROCESS = Process.GetProcesses();

                //打开测试板串口
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
                        btn_Write_MCUBIN.Enabled = true;
                        textBox1.Text = "OPEN TEST BOARD COM SUCCESS";
                        break;
                    }
                    else                                        //no right com
                    {
                        comm.Close();
                    }
                }
                if (!comm.IsOpen)
                {
                    textBox1.Text += "\r\n请检查PC主机与测试板之间的连接，并保证测试板已经上电!\r\n\r\n\r\n请重启软件！";
                }
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
        }

        public bool sendCMD(byte REQ, byte Opration)
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
            int i = 0;
            while ((commBusy == 1) | (!comm.IsOpen))
            {
                i++;
                Thread.Sleep(100);
                if (i == 50)
                {
                    MessageBox.Show("串口已关闭，请重启软件");
                    return false;
                }
            }
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
                commBusy = 0;
                return false;
            }
            commBusy = 0;
            return true;
        }
        public CameraBoard_PCBATest(string barNum)
        {
            InitializeComponent();
            commBuf = new byte[4096];
            BarNum = barNum;
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
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_static_Curent_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_work_Curent_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_Voltage_1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
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

            public int Voltage_5V_OK_NOT;  //0：没结果   1：正常    2：过小   3：过大
            public int Voltage_3_5V_OK_NOT;
            public int Voltage_3_3V_OK_NOT;
            public int Voltage_1_8V_OK_NOT;

            public byte Mode;           // 0:  none 1 : test   2:  work   3.  sleep 
        };
        tlv_reply_t tlv_reply;

        test_resault Test_Resault;
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


        public void Compare_Reply(tlv_reply_t reply, tlv_reply_t replymax, tlv_reply_t replymin)
        {
            Test_Resault.Voltage_5V_OK_NOT = ((reply.Voltage2 > tlv_reply_min.Voltage2) && (reply.Voltage2 < tlv_reply_max.Voltage2)) ? 1 : 2;
            Test_Resault.Voltage_3_5V_OK_NOT = ((reply.Voltage6 > tlv_reply_min.Voltage6) && (reply.Voltage6 < tlv_reply_max.Voltage6)) ? 1 : 2;
            Test_Resault.Voltage_3_3V_OK_NOT = ((reply.Voltage7 > tlv_reply_min.Voltage7) && (reply.Voltage7 < tlv_reply_max.Voltage7)) ? 1 : 2;
            Test_Resault.Voltage_1_8V_OK_NOT = ((reply.Voltage4 > tlv_reply_min.Voltage4) && (reply.Voltage4 < tlv_reply_max.Voltage4)) ? 1 : 2;

            if ((Test_Resault.Voltage_5V_OK_NOT == 1) &&
                (Test_Resault.Voltage_3_5V_OK_NOT == 1) &&
                (Test_Resault.Voltage_3_3V_OK_NOT == 1) &&
                (Test_Resault.Voltage_1_8V_OK_NOT == 1)
                )
                PASS_FLAG = 1;
            else
                PASS_FLAG = 0;
            if (Test_Resault.Voltage_5V_OK_NOT == 1){ this.Invoke((EventHandler)(delegate { textBox_5V.BackColor = Color.Green; })); }
            else{ this.Invoke((EventHandler)(delegate { textBox_5V.BackColor = Color.Red; })); }

            if (Test_Resault.Voltage_3_5V_OK_NOT == 1) { this.Invoke((EventHandler)(delegate { textBox_3_5V.BackColor = Color.Green; })); }
            else { this.Invoke((EventHandler)(delegate { textBox_3_5V.BackColor = Color.Red; })); }

            if (Test_Resault.Voltage_3_3V_OK_NOT == 1) { this.Invoke((EventHandler)(delegate { textBox_3_3V.BackColor = Color.Green; })); }
            else { this.Invoke((EventHandler)(delegate { textBox_3_3V.BackColor = Color.Red; })); }

            if (Test_Resault.Voltage_1_8V_OK_NOT == 1) { this.Invoke((EventHandler)(delegate { textBox_1_8V.BackColor = Color.Green; })); }
            else { this.Invoke((EventHandler)(delegate { textBox_1_8V.BackColor = Color.Red; })); }

        }


        public bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (s.Length != 10)
                return false;

            for (int i = 1; i < 65536; i++)
            {
                ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, (int)column.BarNum];
                if (ran.Text == s)
                {
                    DialogResult resault = MessageBox.Show("此设备已经测试过，请检查是否有重复条码！\r\n若需要重新测试，点击确定！若不需要重新测试，请剪辑取消","提示", MessageBoxButtons.YesNo, MessageBoxIcon.None);
                    if (resault == DialogResult.No)
                        return false;
                    break;
                }
                if (ran.Value == null)
                    break;
            }
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }

        private void btn_Write_MCUBIN_Click(object sender, EventArgs e)
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }

            btn_submit.Enabled = false;
            btn_Test_self.Enabled = false;
            textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
            textBox1.ForeColor = Color.Black;

            BarNum = textBox2.Text.ToString();
            CommBuffWriteIndex = 0;
            CommBuffReadIndex = 0;
            for (int i = 0; i < 4096; i++)
                commBuf[i] = 0;
            while (iRecive == 1) ;
            //下载MCU  BIN文件  线程
            Thread DownLoadMCUBINThred = new Thread(downLoadMCUBIN);
            DownLoadMCUBINThred.Start();
        }

        private void btn_Test_self_Click(object sender, EventArgs e)
        {
            //开始自测   线程
            Thread selfTestThred = new Thread(selfTest);
            selfTestThred.Start();
        }

        Microsoft.Office.Interop.Excel.Range ran;
        private void btn_submit_Click(object sender, EventArgs e)
        {
            int blankRow = 0;
            xApp.DisplayAlerts = false;
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
                ws.Cells[blankRow, (int)column.time] = DateTime.Now.ToLocalTime().ToString();// time;
                ws.Cells[blankRow, (int)column.BarNum] = textBox2.Text.ToString();   //条码
                ws.Cells[blankRow, (int)column.work_current] = textBox_work_current.Text.ToString();
                ws.Cells[blankRow, (int)column.static_current] = textBox_static_Curent.Text.ToString();
                ws.Cells[blankRow, (int)column.voltage_5V] = textBox_5V.Text.ToString();
                ws.Cells[blankRow, (int)column.voltage_3_5V] = textBox_3_5V.Text.ToString(); 
                ws.Cells[blankRow, (int)column.voltage_3_3V] = textBox_3_3V.Text.ToString();
                ws.Cells[blankRow, (int)column.voltage_1_8V] = textBox_1_8V.Text.ToString();
                if (textBox_5V.BackColor == Color.Red)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_5V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_5V];
                    ran.Interior.Color = Color.Green;
                }
                //3.5V
                if (textBox_3_5V.BackColor == Color.Red)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_3_5V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_3_5V];
                    ran.Interior.Color = Color.Green;
                }
                //3.3V
                if (textBox_5V.BackColor == Color.Red)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_3_3V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_3_3V];
                    ran.Interior.Color = Color.Green;
                }
                //1.8V
                if (textBox_5V.BackColor == Color.Red)
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_1_8V];
                    ran.Interior.Color = Color.Red;
                }
                else
                {
                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[blankRow, (int)column.voltage_1_8V];
                    ran.Interior.Color = Color.Green;
                }

                xApp.DisplayAlerts = false;
                ws.SaveAs("摄像头板测试结果");
                if (wb != null)
                {
                    wb.SaveAs(Cdirectory + "\\摄像头板测试结果.xlsx");
                    wb.Saved = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("请重新测试！！！\r\n\r\n\r\n 即将关闭所有EXCEL文档！！！");
                return;
            }
            //提交至web 服务器
        }

        public Process[] EXCEL_PROCESS;
        private void CameraBoard_PCBATest_FormClosing(object sender, FormClosingEventArgs e)
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < (int)timeF.TotalTimer; i++)
                if (TimerArry[i] > 0)
                    TimerArry[i]--;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DeviceERest();
        }
    }
}

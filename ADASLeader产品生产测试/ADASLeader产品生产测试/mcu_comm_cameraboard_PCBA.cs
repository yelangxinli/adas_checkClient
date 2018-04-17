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
namespace ADASLeader产品生产测试
{
    public partial class CameraBoard_PCBATest : Form
    {
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

        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ = 0x18;
        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_RESP = 0x18;

        public int commBusy = 0;
        public int CommBuffWriteIndex = 0;
        public int CommBuffReadIndex = 0;

        public struct tlv_firm_file_t
        {
            public UInt16 type;
            public UInt16 len;
            public UInt32 start;
        };
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
        private bool BufAllFF(ref byte[] buf, int len)
        {
            int i;
            for (i = 0; i < len; i++)
                if (buf[i] != 0xff) return false;
            return true;
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

        public void DeviceERest()
        {
            sendCMD(MSG_TYPE_CMD_TEST_MCU_CODE_RESTART);
            //while (ACKinfo.ACK_CMD_TEST_MCU_RESTART == 0) ;
            //ACKinfo.ACK_CMD_TEST_MCU_RESTART = 0;
            /*
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\n被测设备启动成功！";
            }));*/
        }


        public void CommFuncTestDownload(UInt32 start, byte[] buff, UInt16 len, bool NeedAck)
        {
            UInt32 i, pos;
            msg_head_t msg_head = new msg_head_t();
            tlv_firm_file_t tlv_firm_file = new tlv_firm_file_t();



            byte[] framebuff = new byte[1088];

            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE_CMD;
            msg_head.msg_type = MSG_TYPE_CMD_TEST_MCU_CODE_RESP;
            msg_head.msg_len =
            (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head) + System.Runtime.InteropServices.Marshal.SizeOf(tlv_firm_file) + len);
            msg_head.crc32 = 0;// GetCrc32(framebuff, msg_len);
            byte[] tmpbuff = StructToBytes(msg_head);

            /*
            this.Invoke((EventHandler)(delegate
            {
                BartextBox1.Text += "长度：" + msg_head.msg_len.ToString() + "\r\n";
            }));
            */
            this.Invoke((EventHandler)(delegate { progressBar1.Value++; }));


            for (i = 0; i < System.Runtime.InteropServices.Marshal.SizeOf(msg_head); i++)
                framebuff[i] = tmpbuff[i];
            pos = i;

            tlv_firm_file.type = 0;
            tlv_firm_file.len = len;
            tlv_firm_file.start = start;

            tmpbuff = StructToBytes(tlv_firm_file);
            for (i = 0; i < System.Runtime.InteropServices.Marshal.SizeOf(tlv_firm_file); i++)
                framebuff[pos + i] = tmpbuff[i];
            pos = pos + i;

            for (i = 0; i < len; i++)
                framebuff[pos + i] = buff[i];
            pos = pos + i;
            comm.Write(framebuff, 0, msg_head.msg_len);


            //等待回复
            while (ACKinfo.ACK_CMD_TEST_MCU_CODE == 0) ;
            ACKinfo.ACK_CMD_TEST_MCU_CODE = 0;
        }

        public void EnterBootMod()
        {
            sendCMD(MSG_TYPE_CMD_BOOT_IN_IAP_MODE_REQ);
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
                        //textBox1.Text += "\r\n从测试板 发送 给PC的数据 可能存在丢失！";
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

        public int iRecive = 0;
        public void commDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            iRecive = 1;
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

            iRecive = 0;
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
                                //textBox1.Text += "\r\n无效消息类型";
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
                            if (commBuf[0x07] == 0x01)
                                mainboardBinWrite = 0;
                            else if (commBuf[0x07] == 0x03)
                                mainboardBinWrite = 3;
                            else
                                mainboardBinWrite = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_MCU_CODE_RESTART_RESP:
                            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 1;
                            break;
                        case MSG_TYPE_CMD_BEEP_RESP:
                            ACKinfo.ACK_CMD_BEEP = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ:
                            ACKinfo.ACK_CMD_LOOK_FOR_COM = 1;
                            break;
                        default:
                            this.Invoke((EventHandler)(delegate
                            {
                                //textBox1.Text += "\r\n无效消息类型";
                            }));
                            break;
                    }
                }
                else
                {
                    this.Invoke((EventHandler)(delegate
                    {
                        //textBox1.Text += "\r\n无效服务类型！";
                    }));
                }

                CommBuffReadIndex += commBuf[4] + commBuf[5] * 256;
            }
        }

        public void startTest()   //发送开始测试指令
        {
            //生成测试指令byte[]数组
            byte[] frambuf = new byte[250];
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE_CMD;
            msg_head.msg_type = MSG_TYPE_CMD_START_TEST_REQ;
            msg_head.msg_len = (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head));

            byte[] tmpbuff = StructToBytes(msg_head);
            for (int i = 0; i < msg_head.msg_len; i++)
                frambuf[i] = tmpbuff[i];


            //发送测试指令
            while ((commBusy == 1) | (!comm.IsOpen)) ;
            try
            {
                commBusy = 1;
                if (comm.IsOpen)
                {
                    comm.Write(frambuf, 0, 16);
                }
                else
                {
                    comm.Open();
                    comm.Write(frambuf, 0, 16);
                }
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\nSTART";
                }));
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


        public bool writeSTM32MCU(string FileName)
        {
            UInt32 addr;
            BinaryReader br;
            bool err = false;
            UInt32 count = 0;

            startTest();
            long lSize = new FileInfo(FileName).Length;
            this.Invoke((EventHandler)(delegate { progressBar1.Minimum = 0; }));
            this.Invoke((EventHandler)(delegate { progressBar1.Value = 0; }));
            if (lSize % 1024 == 0)
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024) + 10; }));
            else
                this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024) + 1 + 10; }));

            while (ACKinfo.ACK_CMD_Start_Test == 0) ;
            ACKinfo.ACK_CMD_Start_Test = 0;

            if (commBuf[7] == 0x01)
            {
                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n烧录程序握手失败，请重试！"; }));
                this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true; }));
                return false;
            }
            else
            {
                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n烧录程序握手成功"; }));
                this.Invoke((EventHandler)(delegate { progressBar1.Value += 5; }));
                //this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = false; }));
            }

            //进入boot模式
            EnterBootMod();
            while (ACKinfo.ACK_CMD_BOOT_IN_IAP_MODE == 0) ;
            ACKinfo.ACK_CMD_BOOT_IN_IAP_MODE = 0;
            if (commBuf[7] == 0x01)
            {
                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n进入boot模式，失败！"; }));
                this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true; }));
                this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.PerformClick(); }));
                return false;
            }
            else
            {
                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n进入boot模式，成功"; }));
                this.Invoke((EventHandler)(delegate { progressBar1.Value += 5; }));
                //this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = false; }));
            }

            //打开文件  并发送BIN文件
            if (!File.Exists(FileName))
            {
                //MessageBox.Show("CAN NOT", "MH680产线测试工具");
                this.Invoke((EventHandler)(delegate { textBox1.Text += FileName + "\r\nIS NOT EXIST"; }));
                return false;
            }
            else
            {
                br = new BinaryReader(new FileStream(FileName, FileMode.Open));
                addr = (UInt32)0x08000000;
                byte[] binbuf = new byte[1024];

                count = 0;

                //this.Invoke((EventHandler)(delegate{textBox1.Text += "\r\n下载MCU BIN文件开始\r\n";}));
                mainboardBinWrite = 1;
                while (mainboardBinWrite == 1)
                {
                    binbuf = br.ReadBytes(1024);

                    if (binbuf.Length <= 0) break;
                    /* If buff all ff, we do not write */
                    if (BufAllFF(ref binbuf, binbuf.Length) == false)
                        CommFuncTestDownload(addr, binbuf, (UInt16)binbuf.Length, true);
                    else
                        this.Invoke((EventHandler)(delegate { progressBar1.Value++; }));
                    addr = addr + (UInt32)binbuf.Length;
                    //if (binbuf.Length < 1024) break;

                }
                if (mainboardBinWrite == 0)
                {
                    this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nWRITE BIN FAILED"; }));
                    br.Close();
                    return false;
                }
                else if (mainboardBinWrite == 3)
                {
                    this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nWRITE BIN INTERRUPTED , PLEASE TRY AGAIN."; }));
                    br.Close();
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.PerformClick(); }));
                    return false;
                }
                else
                    this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nWRITE BIN SUCCESS"; }));
            }
            br.Close();

            //Thread.Sleep(3000);

            //重启设备
            DeviceERest();
            Thread.Sleep(1500);
            //烧录WIFI CODE文件

            //使能 软件自测 和 人工测试 两个功能
            this.Invoke((EventHandler)(delegate {
                //btn_Write_MCUBIN.Enabled = false;
                btn_Test_self.Enabled = true;
            }));

            return true;
        }
        public void downLoadMCUBIN()  //等待 开始测试指令回复  后开始下载MCU 的BIN文件
        {
            writeSTM32MCU("CAMERA_MCU_BIN.BIN");
        }


        public void selfTest()
        {

            TimerArry[(int)timeF.selfTest] = 2;
            while (TimerArry[(int)timeF.selfTest] > 0) ;  //等待2S

            //获取自测结果(下发自测指令、等待ACK)
            sendCMD(MSG_TYPE_CMD_TEST_RESULT_REQ);
            for (int i = 10; i > 0; i--)
            {
                //发送 请求测试结果指令
                TimerArry[(int)timeF.selfTest] = 10;
                while (TimerArry[(int)timeF.selfTest] > 0)  //等待10S钟
                {
                    if (ACKinfo.ACK_CMD_Test_Resualt == 1)
                    {
                        //显示测试结果
                        this.Invoke((EventHandler)(delegate { textBox_5V.Text = tlv_reply.Voltage2.ToString(); }));//5V
                        this.Invoke((EventHandler)(delegate { textBox_3_5V.Text = tlv_reply.Voltage6.ToString(); }));//3.5V
                        this.Invoke((EventHandler)(delegate { textBox_3_3V.Text = tlv_reply.Voltage7.ToString(); }));//3.3V
                        this.Invoke((EventHandler)(delegate { textBox_1_8V.Text = tlv_reply.Voltage4.ToString(); }));//1.8V

                        this.Invoke((EventHandler)(delegate { textBox_static_Curent.Text = tlv_reply.Current0.ToString(); }));  //静态电流
                        this.Invoke((EventHandler)(delegate { textBox_work_current.Text = tlv_reply.Current1.ToString(); }));    //工作电流

                        Compare_Reply(tlv_reply, tlv_reply_max, tlv_reply_min);

                        if (PASS_FLAG == 1)
                        {
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Green;
                                textBox1.Text = "PASS";
                                btn_submit.Enabled = true;
                            }));
                            return;
                        }
                        else
                        {
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Red;
                                textBox1.Text = "NO";
                                btn_submit.Enabled = true;
                            }));
                            return;
                        }
                    }
                }
            }
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                textBox1.ForeColor = Color.Red;
                textBox1.Text = "TIME OUT";
            }));
        }


    }
}

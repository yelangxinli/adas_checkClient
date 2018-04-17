﻿using System;
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
namespace ADASLeader产品生产测试
{
    public partial class DVRBoard_PCBATest : Form
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
            while (ACKinfo.ACK_CMD_TEST_MCU_RESTART == 0) ;
            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 0;

            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\n被测设备启动成功！";
            }));
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
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "@";
            }));


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
                    textBox1.Text += "\r\nStart Test指令 通过串口发送！";
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

        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9]*$";
            Regex rx = new Regex(pattern);
            return rx.IsMatch(s);
        }

        public void selfTest()
        {
            if ((textBox2.Text == null) || (!IsNumber(textBox2.Text.ToString())))
            {
                MessageBox.Show("请先输入正确的条形码");
                return;
            }
            if (!comm.IsOpen)
            {
                MessageBox.Show("请打开串口！");
                return;
            }

            DialogResult resault = MessageBox.Show("程序是否已经通过DirectUSB烧录成功","提示",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (resault == DialogResult.No)
                return;
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text = "开始测试";
            }));

            TimerArry[(int)timeF.selfTest] = 2;
            while (TimerArry[(int)timeF.selfTest] > 0) ;  //等待2S

            //获取自测结果(下发自测指令、等待ACK)
            //for (int i = 5; i > 0; i--)
            {
                //发送 请求测试结果指令
                sendCMD(MSG_TYPE_CMD_TEST_RESULT_REQ);
                TimerArry[(int)timeF.selfTest] = 10;
                while (TimerArry[(int)timeF.selfTest] > 0)  //等待10S钟
                {
                    Thread.Sleep(1000);
                    this.Invoke((EventHandler)(delegate { textBox1.Text += ((int)timeF.selfTest).ToString(); }));   //1.1V
                    if (ACKinfo.ACK_CMD_Test_Resualt == 1)
                    {
                        //显示测试结果
                        //this.Invoke((EventHandler)(delegate { textBox_static_Curent.Text = tlv_reply.Current1.ToString(); }));  //静态电流

                        //this.Invoke((EventHandler)(delegate { textBox_work_current.Text = tlv_reply.Current0.ToString(); }));   //工作电流

                        this.Invoke((EventHandler)(delegate { textBox_5V.Text = tlv_reply.Voltage2.ToString(); }));  //5V
                        this.Invoke((EventHandler)(delegate { textBox_1_8V.Text = tlv_reply.Voltage3.ToString(); }));   //1.8V
                        this.Invoke((EventHandler)(delegate { textBox_3_3V.Text = tlv_reply.Voltage0.ToString(); }));  //3.3V
                        this.Invoke((EventHandler)(delegate { textBox_1_1V.Text = tlv_reply.Voltage7.ToString(); }));   //1.1V
                        
                        Compare_Reply(tlv_reply, tlv_reply_max, tlv_reply_min);

                        if (Test_Resault._5V_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_5V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_5V.BackColor = Color.Red; }));

                        if (Test_Resault._3_3_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_3_3V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_3_3V.BackColor = Color.Red; }));

                        if (Test_Resault._1_1_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_1_1V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_1_1V.BackColor = Color.Red; }));

                        if (Test_Resault._1_8_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_1_8V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_1_8V.BackColor = Color.Red; }));
                        this.Invoke((EventHandler)(delegate { btn_submit.Enabled = true; }));
                        break;
                    }
                }
            }
            if (ACKinfo.ACK_CMD_Test_Resualt == 0)
                this.Invoke((EventHandler)(delegate {
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Red;
                    textBox1.Text = "TIME OUT";
                }));
        }

    }
}

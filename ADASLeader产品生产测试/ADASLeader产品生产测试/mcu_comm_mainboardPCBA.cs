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
    public partial class MainBoard_PCBATest : Form
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

        public const byte MSG_TYPE_CMD_BEEP_SWITCH_CAN_REQ = 0x0F;
        public const byte MSG_TYPE_CMD_BEEP_SWITCH_CAN_RESP = 0x0F;

        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ = 0x18;
        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_RESP = 0x18;

        public const byte MSG_TYPE_WRITE_DEVICE_FILE_REQ = 0x20;
        public const byte MSG_TYPE_WRITE_DEVICE_FILE_RESP = 0x20;

        public const byte MSG_TYPE_WRITE_DEVICE_FILE_MORE_REQ = 0x21;
        public const byte MSG_TYPE_WRITE_DEVICE_FILE_MORE_RESP = 0x21;

        public int commBusy = 0;
        public int CommBuffWriteIndex = 0;
        public int CommBuffReadIndex = 0;

        private static uint[] lstCRC = new uint[256]  {
            0x00000000,0x04C11DB7,0x09823B6E,0x0D4326D9,0x130476DC,0x17C56B6B,0x1A864DB2,0x1E475005,
            0x2608EDB8,0x22C9F00F,0x2F8AD6D6,0x2B4BCB61,0x350C9B64,0x31CD86D3,0x3C8EA00A,0x384FBDBD,
            0x4C11DB70,0x48D0C6C7,0x4593E01E,0x4152FDA9,0x5F15ADAC,0x5BD4B01B,0x569796C2,0x52568B75,
            0x6A1936C8,0x6ED82B7F,0x639B0DA6,0x675A1011,0x791D4014,0x7DDC5DA3,0x709F7B7A,0x745E66CD,
            0x9823B6E0,0x9CE2AB57,0x91A18D8E,0x95609039,0x8B27C03C,0x8FE6DD8B,0x82A5FB52,0x8664E6E5,
            0xBE2B5B58,0xBAEA46EF,0xB7A96036,0xB3687D81,0xAD2F2D84,0xA9EE3033,0xA4AD16EA,0xA06C0B5D,
            0xD4326D90,0xD0F37027,0xDDB056FE,0xD9714B49,0xC7361B4C,0xC3F706FB,0xCEB42022,0xCA753D95,
            0xF23A8028,0xF6FB9D9F,0xFBB8BB46,0xFF79A6F1,0xE13EF6F4,0xE5FFEB43,0xE8BCCD9A,0xEC7DD02D,
            0x34867077,0x30476DC0,0x3D044B19,0x39C556AE,0x278206AB,0x23431B1C,0x2E003DC5,0x2AC12072,
            0x128E9DCF,0x164F8078,0x1B0CA6A1,0x1FCDBB16,0x018AEB13,0x054BF6A4,0x0808D07D,0x0CC9CDCA,
            0x7897AB07,0x7C56B6B0,0x71159069,0x75D48DDE,0x6B93DDDB,0x6F52C06C,0x6211E6B5,0x66D0FB02,
            0x5E9F46BF,0x5A5E5B08,0x571D7DD1,0x53DC6066,0x4D9B3063,0x495A2DD4,0x44190B0D,0x40D816BA,
            0xACA5C697,0xA864DB20,0xA527FDF9,0xA1E6E04E,0xBFA1B04B,0xBB60ADFC,0xB6238B25,0xB2E29692,
            0x8AAD2B2F,0x8E6C3698,0x832F1041,0x87EE0DF6,0x99A95DF3,0x9D684044,0x902B669D,0x94EA7B2A,
            0xE0B41DE7,0xE4750050,0xE9362689,0xEDF73B3E,0xF3B06B3B,0xF771768C,0xFA325055,0xFEF34DE2,
            0xC6BCF05F,0xC27DEDE8,0xCF3ECB31,0xCBFFD686,0xD5B88683,0xD1799B34,0xDC3ABDED,0xD8FBA05A,
            0x690CE0EE,0x6DCDFD59,0x608EDB80,0x644FC637,0x7A089632,0x7EC98B85,0x738AAD5C,0x774BB0EB,
            0x4F040D56,0x4BC510E1,0x46863638,0x42472B8F,0x5C007B8A,0x58C1663D,0x558240E4,0x51435D53,
            0x251D3B9E,0x21DC2629,0x2C9F00F0,0x285E1D47,0x36194D42,0x32D850F5,0x3F9B762C,0x3B5A6B9B,
            0x0315D626,0x07D4CB91,0x0A97ED48,0x0E56F0FF,0x1011A0FA,0x14D0BD4D,0x19939B94,0x1D528623,
            0xF12F560E,0xF5EE4BB9,0xF8AD6D60,0xFC6C70D7,0xE22B20D2,0xE6EA3D65,0xEBA91BBC,0xEF68060B,
            0xD727BBB6,0xD3E6A601,0xDEA580D8,0xDA649D6F,0xC423CD6A,0xC0E2D0DD,0xCDA1F604,0xC960EBB3,
            0xBD3E8D7E,0xB9FF90C9,0xB4BCB610,0xB07DABA7,0xAE3AFBA2,0xAAFBE615,0xA7B8C0CC,0xA379DD7B,
            0x9B3660C6,0x9FF77D71,0x92B45BA8,0x9675461F,0x8832161A,0x8CF30BAD,0x81B02D74,0x857130C3,
            0x5D8A9099,0x594B8D2E,0x5408ABF7,0x50C9B640,0x4E8EE645,0x4A4FFBF2,0x470CDD2B,0x43CDC09C,
            0x7B827D21,0x7F436096,0x7200464F,0x76C15BF8,0x68860BFD,0x6C47164A,0x61043093,0x65C52D24,
            0x119B4BE9,0x155A565E,0x18197087,0x1CD86D30,0x029F3D35,0x065E2082,0x0B1D065B,0x0FDC1BEC,
            0x3793A651,0x3352BBE6,0x3E119D3F,0x3AD08088,0x2497D08D,0x2056CD3A,0x2D15EBE3,0x29D4F654,
            0xC5A92679,0xC1683BCE,0xCC2B1D17,0xC8EA00A0,0xD6AD50A5,0xD26C4D12,0xDF2F6BCB,0xDBEE767C,
            0xE3A1CBC1,0xE760D676,0xEA23F0AF,0xEEE2ED18,0xF0A5BD1D,0xF464A0AA,0xF9278673,0xFDE69BC4,
            0x89B8FD09,0x8D79E0BE,0x803AC667,0x84FBDBD0,0x9ABC8BD5,0x9E7D9662,0x933EB0BB,0x97FFAD0C,
            0xAFB010B1,0xAB710D06,0xA6322BDF,0xA2F33668,0xBCB4666D,0xB8757BDA,0xB5365D03,0xB1F740B4};
        public static uint GetCrc32(byte[] buf, int startIndex = 0, int len = 0)
        {
            if (buf == null || buf.Length == 0)
                return 0;
            if (len == 0)
                len = buf.Length - startIndex;
            int v1, v3;
            uint a1 = 0;
            for (int i = startIndex; i < (startIndex + len); i++)
            {
                v1 = buf[i] ^ ((byte)(i - startIndex));
                v3 = (int)(v1 ^ ((a1 >> 24) & 0xFF));
                a1 = lstCRC[v3] ^ (a1 << 8);
            }
            return a1;
        }

        public struct TP_FILE_PARA
        {

            public UInt16 TLV_TYPE;
            public UInt16 TLV_LEN;
            public byte[] NAME;
            public UInt32 FILE_LENGTH;
            public UInt32 FILE_CRC;
        };
        public struct HEAD
        {
            public UInt32 START;
            public UInt32 CRC;
            public UInt16 MSG_LEN;
            public byte SERVICE_TYPE;
            public byte MSG_TYPE;
            public UInt32 UNDEFINE;
        };

        public struct TP_FILE_PACKET
        {
            public UInt16 TLV_TYPE;
            public UInt16 TLV_LEN;
            public UInt16 TotalPacks;
            public UInt16 CurPack;
            public UInt16 CurDataLength;
        };

        HEAD head;
        TP_FILE_PACKET tp_file_pack;
        TP_FILE_PARA tp_file_para;

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
            
            while (ACKinfo.ACK_CMD_TEST_MCU_RESTART == 0);
            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 0;

            //this.Invoke((EventHandler)(delegate{textBox1.Text += "\r\nMAINBOARD RESTARTED , PLEASE GO ON TEST";}));
            
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
            //this.Invoke((EventHandler)(delegate { textBox1.Text += "@"; }));
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
                        textBox1.Font = new Font(textBox1.Font.FontFamily, 50, textBox1.Font.Style);
                        textBox1.ForeColor = Color.Black;
                        textBox1.Text = null;

                        textBox2.Enabled = true;
                        textBox1.Text = "通讯异常，即将重启!";

                        DialogResult resault = MessageBox.Show("通讯异常，是否重启软件?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (resault == DialogResult.Yes)
                            System.Environment.Exit(0);
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
            byte serviceType; byte message;int valid = 1;
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
                        case MSG_TYPE_WRITE_DEVICE_FILE_MORE_REQ:
                            ACKinfo.ACK_DEVICE_FILE_MORE_Write = 1;
                            break;
                        case MSG_TYPE_WRITE_DEVICE_FILE_REQ:
                            ACKinfo.ACK_DEVICE_FILE_Write = 1;
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
                            if (commBuf[0x07] == 0x01)
                                mainboardBinWrite = 0;
                            else if(commBuf[0x07] == 0x03)
                                mainboardBinWrite = 3;
                            else
                                mainboardBinWrite = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_MCU_CODE_RESTART_RESP:
                            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 1;
                            break;
                        case MSG_TYPE_CMD_BEEP_SWITCH_CAN_RESP:
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
                //this.Invoke((EventHandler)(delegate{textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;}));

                //MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
                comm.Close();
            }
            finally
            {
                commBusy = 0;
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
                //this.Invoke((EventHandler)(delegate{textBox1.Text += "\r\nWRITE MCU";}));
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate{textBox1.Text += "CAN NOT OPEN COM PORT \r\n" + e.Message; }));
                //MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
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
            this.Invoke((EventHandler)(delegate { textBox1.Text = "\r\n\r\n" + FileName + "\r\n" + textBox1.Text; }));
            long lSize = new FileInfo(FileName).Length;
            this.Invoke((EventHandler)(delegate { textBox1.Text = textBox1.Text + "\r\n"; }));
            while (ACKinfo.ACK_CMD_Start_Test == 0)
            {
                this.Invoke((EventHandler)(delegate { textBox1.Text = "." + textBox1.Text; }));
                Thread.Sleep(500);
            }
            //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n擦除FLASH成功"; }));
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
                //this.Invoke((EventHandler)(delegate { progressBar1.Value += 5; }));
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
                //this.Invoke((EventHandler)(delegate { progressBar1.Value += 5; }));
                //this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = false; }));
            }

            //打开文件  并发送BIN文件
            if (!File.Exists(FileName))
            {
                //MessageBox.Show("CAN NOT", "MH680产线测试工具");
                this.Invoke((EventHandler)(delegate { textBox1.Text = FileName + "IS NOT EXIST\r\n" + textBox1.Text; }));
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
                    this.Invoke((EventHandler)(delegate { textBox1.Text = "WRITE BIN FAILED\r\n" + textBox1.Text; }));
                    br.Close();
                    return false;
                }
                else if (mainboardBinWrite == 3)
                {
                    this.Invoke((EventHandler)(delegate { textBox1.Text = "WRITE BIN INTERRUPTED , PLEASE TRY AGAIN.\r\n" + textBox1.Text; }));
                    br.Close();
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.PerformClick(); }));
                    return false;
                }
                //else
                    //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nWRITE BIN SUCCESS"; }));
            }
            br.Close();

            //Thread.Sleep(3000);

            //重启设备
            DeviceERest();

            this.Invoke((EventHandler)(delegate { textBox1.Text = FileName + " OK." + "\r\n"  + textBox1.Text; }));
            //烧录WIFI CODE文件

            //使能 软件自测 和 人工测试 两个功能
            return true;
        }


        public bool DownLoadFileToMCU(string FileName)
        {
            UInt32 addr;
            BinaryReader br;
            bool err = false;
            UInt32 count = 0;


            long lSize = 0;
            if (!File.Exists(FileName))
            {
                this.Invoke((EventHandler)(delegate { textBox1.Text = "can not founf file -- " + FileName + "\r\n" + textBox1.Text; }));
                return false;
            }
            else
            {
                this.Invoke((EventHandler)(delegate { textBox1.Text =  FileName + "\r\n\r\n" + textBox1.Text; }));
                lSize = new FileInfo(FileName).Length;

                //打开文件
                br = new BinaryReader(new FileStream(FileName, FileMode.Open));
                byte[] wholeFileBin = new byte[lSize];
                wholeFileBin = br.ReadBytes((int)lSize);

                //first packet
                //head
                head.START = 0xAAAAAAAA;
                head.SERVICE_TYPE = SERVICE_TYPE_FILE;
                head.MSG_TYPE = MSG_TYPE_WRITE_DEVICE_FILE_REQ; // write file
                head.UNDEFINE = 0;
                //TP_FILE_PARA
                tp_file_para.TLV_TYPE = 0x02;
                tp_file_para.TLV_LEN = 76;
                tp_file_para.FILE_LENGTH = (UInt32)lSize;
                tp_file_para.FILE_CRC = GetCrc32(wholeFileBin, 0, (int)lSize);
                tp_file_para.NAME = new byte[64];
                byte[] byteArray = System.Text.Encoding.ASCII.GetBytes(FileName);
                for (int i = 0; i < byteArray.Length; i++)
                    tp_file_para.NAME[i] = byteArray[i];
                br.Close();
                //TP_FILE_PACKET
                tp_file_pack.TLV_TYPE = 0x04;
                if(tp_file_para.FILE_LENGTH % 1024 == 0)
                    tp_file_pack.TotalPacks = (UInt16)(tp_file_para.FILE_LENGTH / 1024);
                else
                    tp_file_pack.TotalPacks = (UInt16)(tp_file_para.FILE_LENGTH / 1024 + 1);

                //this.Invoke((EventHandler)(delegate { progressBar1.Minimum = 0; }));
                //this.Invoke((EventHandler)(delegate { progressBar1.Maximum = tp_file_pack.TotalPacks; }));
                //this.Invoke((EventHandler)(delegate { progressBar1.Value = 0; }));
                

                if (tp_file_pack.TotalPacks > 1)
                {
                    tp_file_pack.TLV_LEN = (UInt16)(10);
                    tp_file_pack.CurPack = 0x00;
                    tp_file_pack.CurDataLength = 1024;

                    head.MSG_LEN = 16 + 76 + 10 + 1024;

                    br = new BinaryReader(new FileStream(FileName, FileMode.Open));
                    byte[] firstFileBinPACK = new byte[1024];
                    firstFileBinPACK = br.ReadBytes(1024);

                    byte[] firstPackBIN = new byte[16 + 76 + 10 + 1024];

                    byte[] tmp = StructToBytes(head);
                    for (int i = 0; i < 16; i++)
                        firstPackBIN[i] = tmp[i];
                    tmp = StructToBytes(tp_file_para);
                    for (int i = 0; i < 4; i++)
                        firstPackBIN[i + 16] = tmp[i];
                    for (int i = 0; i < 64; i++)
                        firstPackBIN[i + 16 + 4] = tp_file_para.NAME[i];
                    for (int i = 0; i < 8; i++)
                        firstPackBIN[i + 16 + 4 + 64 ] = tmp[i + 8];
                    tmp = StructToBytes(tp_file_pack);
                    for (int i = 0; i < 10; i++)
                        firstPackBIN[i + 16 + tp_file_para.TLV_LEN] = tmp[i];
                    //data
                    for (int i = 0; i < 1024; i++)
                        firstPackBIN[i + 16 + tp_file_para.TLV_LEN + 10] = firstFileBinPACK[i];


                    head.CRC = GetCrc32(firstPackBIN, 8 , 16 + tp_file_para.TLV_LEN + tp_file_pack.TLV_LEN + 1024 - 8);
                    tmp = StructToBytes(head);
                    for (int i = 0; i < 16; i++)
                        firstPackBIN[i] = tmp[i];

                    
                    ACKinfo.ACK_DEVICE_FILE_Write = 0;
                    //send data
                    //br.Close();
                    //return true;
                    //wait for ack
                    //this.Invoke((EventHandler)(delegate { textBox1.Text += "@"; }));
                    this.Invoke((EventHandler)(delegate { progressBar1.Value++; }));
                    Thread.Sleep(2);
                    comm.Write(firstPackBIN, 0, firstPackBIN.Length);
                    for (int i = 0; i < 12; i++)
                    {
                        if (ACKinfo.ACK_DEVICE_FILE_Write == 1)
                            break;
                        Thread.Sleep(1000);
                    }
                    if ((commBuf[7] == 0) || (ACKinfo.ACK_DEVICE_FILE_Write == 0))
                    {
                        this.Invoke((EventHandler)(delegate { textBox1.Text = "烧录失败\r\n" + textBox1.Text; }));
                        br.Close();
                        return false;
                    }
                    tp_file_pack.CurPack = (UInt16)(tp_file_pack.CurPack + 1);
                    //for (int i = 0; i < 16 + tp_file_para.TLV_LEN + tp_file_pack.TLV_LEN; i++)
                        //this.Invoke((EventHandler)(delegate { textBox1.Text += string.Format("{0:X000}", firstPackBIN[i]) + " "; }));
                    //return true;
                    
                    //other packets
                    byte[] otherPackBIN = new byte[16 + 10 + 1024];
                    byte[] otherFileBinPACK = new byte[1024];
                    head.MSG_TYPE = MSG_TYPE_WRITE_DEVICE_FILE_MORE_REQ; // write file
                    while (tp_file_pack.CurPack < tp_file_pack.TotalPacks)
                    {
                        otherFileBinPACK = br.ReadBytes(1024);
                        tp_file_pack.CurDataLength = (UInt16)otherFileBinPACK.Length;
                        tp_file_pack.TLV_LEN = (UInt16)(10);

                        //fill head
                        head.MSG_LEN = (UInt16)(16 + tp_file_pack.TLV_LEN + (UInt16)otherFileBinPACK.Length);
                        tmp = StructToBytes(head);
                        for (int i = 0; i < 16; i++)
                            otherPackBIN[i] = tmp[i];
                        //fill tp_file_pack
                        tmp = StructToBytes(tp_file_pack);
                        for (int i = 0; i < 10; i++)
                            otherPackBIN[i + 16] = tmp[i];
                        //fill data
                        for (int i = 0; i < otherFileBinPACK.Length; i++)
                            otherPackBIN[i + 16 + 10] = otherFileBinPACK[i];

                        tmp = StructToBytes(head);
                        head.CRC = GetCrc32(otherPackBIN, 8, 16 + tp_file_pack.TLV_LEN + (UInt16)otherFileBinPACK.Length - 8);
                        tmp = StructToBytes(head);
                        for (int i = 0; i < 16; i++)
                            otherPackBIN[i] = tmp[i];

                        this.Invoke((EventHandler)(delegate { progressBar1.Value++; }));

                        ACKinfo.ACK_DEVICE_FILE_MORE_Write = 0;
                        //send data
                        comm.Write(otherPackBIN, 0, 16 + 10 + otherFileBinPACK.Length);
                        //wait for ack
                        for (int i = 0; i < 11000; i++)
                        {
                            if (ACKinfo.ACK_DEVICE_FILE_MORE_Write == 1)
                                break;
                            Thread.Sleep(10);
                        }
                        if (ACKinfo.ACK_DEVICE_FILE_MORE_Write == 0)
                        {
                            this.Invoke((EventHandler)(delegate { textBox1.Text = "DOWNLOAD HAD BEEN INTERRUPT , PLEASE TRY AGAIN.\r\n" + textBox1.Text; }));
                            br.Close();
                            return false;
                        }
                        /*for (int i= 0; i < 50; i++)
                        {
                            if (ACKinfo.ACK_DEVICE_FILE_MORE_Write == 1)
                                break;
                            else
                                Thread.Sleep(10);
                        }
                        if (ACKinfo.ACK_DEVICE_FILE_MORE_Write == 0)
                        {
                            this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n传输超时，请重新下载！"; }));
                            return false;
                        }*/
                        if (commBuf[7] == 0)
                        {
                            this.Invoke((EventHandler)(delegate { textBox1.Text = "DOWNLOAD HAD BEEN INTERRUPT , PLEASE TRY AGAIN.\r\n" + textBox1.Text; }));
                            br.Close();
                            return false;
                        }
                        tp_file_pack.CurPack = (UInt16)(tp_file_pack.CurPack + 1);
                    }
                    if(tp_file_pack.CurPack == tp_file_pack.TotalPacks)
                        this.Invoke((EventHandler)(delegate { textBox1.Text =  FileName + " OK." + "\r\n" + textBox1.Text; }));
                    else
                        this.Invoke((EventHandler)(delegate { textBox1.Text =  FileName + " FAILED.\r\n" + textBox1.Text; }));
                    br.Close();
                    //this.Invoke((EventHandler)(delegate { progressBar1.Value = 0; }));
                }
                else
                {
                    tp_file_pack.TLV_LEN = (UInt16)(10 );
                    tp_file_pack.CurPack = 0x00;
                    tp_file_pack.CurDataLength = (UInt16)lSize;

                    head.MSG_LEN = (UInt16)(16 + 76 + 10 + lSize);

                    br = new BinaryReader(new FileStream(FileName, FileMode.Open));
                    byte[] firstFileBinPACK = new byte[lSize];
                    firstFileBinPACK = br.ReadBytes((int)lSize);

                    byte[] firstPackBIN = new byte[16 + 76 + 10 + lSize];

                    byte[] tmp = StructToBytes(head);
                    for (int i = 0; i < 16; i++)
                        firstPackBIN[i] = tmp[i];

                    tmp = StructToBytes(tp_file_para);
                    for (int i = 0; i < 4; i++)
                        firstPackBIN[i + 16] = tmp[i];
                    for (int i = 0; i < 64; i++)
                        firstPackBIN[i + 16 + 4] = tp_file_para.NAME[i];
                    for (int i = 0; i < 8; i++)
                        firstPackBIN[i + 16 + 4 + 64] = tmp[i + 8];

                    tmp = StructToBytes(tp_file_pack);
                    for (int i = 0; i < 10; i++)
                        firstPackBIN[i + 16 + tp_file_para.TLV_LEN] = tmp[i];
                    //data 
                    for (int i = 0; i < lSize; i++)
                        firstPackBIN[i + 16 + tp_file_para.TLV_LEN + 10] = firstFileBinPACK[i];

                    head.CRC = GetCrc32(firstPackBIN, 4, 16 + tp_file_para.TLV_LEN + tp_file_pack.TLV_LEN + (int)lSize + (int)lSize - 4);
                    tmp = StructToBytes(head);
                    for (int i = 0; i < 16; i++)
                        firstPackBIN[i] = tmp[i];

                    ACKinfo.ACK_DEVICE_FILE_Write = 0;
                    //send data
                    comm.Write(firstPackBIN, 0, firstPackBIN.Length);
                    //wait for ack
                    //this.Invoke((EventHandler)(delegate { textBox1.Text += "@"; }));
                    this.Invoke((EventHandler)(delegate { progressBar1.Value++; }));
                    while (ACKinfo.ACK_DEVICE_FILE_Write == 0) ;
                    ACKinfo.ACK_DEVICE_FILE_Write = 0;
                    if (commBuf[7] == 0)
                    {
                        this.Invoke((EventHandler)(delegate { textBox1.Text = "DOWNLOAD HAD BEEN INTERRUPT , PLEASE TRY AGAIN.\r\n" + textBox1.Text; }));
                        br.Close();
                        return false;
                    }
                    tp_file_pack.CurPack = (UInt16)(tp_file_pack.CurPack + 1);
                    if (tp_file_pack.CurPack == tp_file_pack.TotalPacks)
                        this.Invoke((EventHandler)(delegate { textBox1.Text = FileName + " OK.\r\n" + textBox1.Text; }));
                    else
                        this.Invoke((EventHandler)(delegate { textBox1.Text =  FileName + " FAILED.\r\n" + textBox1.Text; }));
                    br.Close();
                }
            }
            return true;
        }

        public void clerCommBuf()
        {
            for (int i = 0; i < 4096; i++)
                commBuf[i] = 0;
            CommBuffWriteIndex = 0;
            CommBuffReadIndex = 0;
        }
        public void downLoadBIN()  //等待 开始测试指令回复  后开始下载MCU 的BIN文件
        {
            try
            {
                if ((!File.Exists("MCU_BIN_FIRST.BIN")) || (!File.Exists("PIC_20180322.BIN")) || (!File.Exists("FPGA_01.BIN")) || (!File.Exists("MCU_BIN.BIN")))
                {
                    MessageBox.Show("请检查BIN文件是否全");
                    return;
                }
                this.Invoke((EventHandler)(delegate { progressBar1.Minimum = 0; }));
                this.Invoke((EventHandler)(delegate { progressBar1.Value = 0; }));
                long lSize = new FileInfo("MCU_BIN_FIRST.BIN").Length;
                if (lSize % 1024 == 0)
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024); }));
                else
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum = ((int)lSize / 1024) + 1; }));

                lSize = new FileInfo("PIC_20180322.BIN").Length;
                if (lSize % 1024 == 0)
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024); }));
                else
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024) + 1; }));

                lSize = new FileInfo("FPGA_01.BIN").Length;
                if (lSize % 1024 == 0)
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024); }));
                else
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024) + 1; }));

                lSize = new FileInfo("MCU_BIN.BIN").Length;
                if (lSize % 1024 == 0)
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024); }));
                else
                    this.Invoke((EventHandler)(delegate { progressBar1.Maximum += ((int)lSize / 1024) + 1; }));

                this.Invoke((EventHandler)(delegate { textBox1.Text = "开始烧录\r\n" + textBox1.Text; }));
                if (writeSTM32MCU("MCU_BIN_FIRST.BIN") == false)
                {
                    MessageBox.Show("烧录MCU_BIN_FIRST失败");
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true;textBox2.Enabled = true; }));
                    return;
                }
                Thread.Sleep(5000);
                clerCommBuf();
                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nBEGAIN TO WRITE FILE "; }));
                if (DownLoadFileToMCU("PIC_20180322.BIN") == false)
                {
                    MessageBox.Show("烧录PIC_20180322失败");
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true; textBox2.Enabled = true; }));
                    return;
                }
                Thread.Sleep(5000);
                clerCommBuf();
                if (DownLoadFileToMCU("FPGA_01.BIN") == false)//write FPGA bin
                {
                    MessageBox.Show("烧录FPGA_01失败");
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true; textBox2.Enabled = true; }));
                    return;
                }
                Thread.Sleep(5000);
                clerCommBuf();
                if (writeSTM32MCU("MCU_BIN.BIN") == false)
                {
                    MessageBox.Show("烧录MCU_BIN失败");
                    this.Invoke((EventHandler)(delegate { btn_Write_MCUBIN.Enabled = true; textBox2.Enabled = true; }));
                    return;
                }
                clerCommBuf();

                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text = "等待FPGA写入完成\r\n\r\n烧录成功\r\n" + textBox1.Text;
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Blue;
                }));
                
                for (int i = 60; i > 0; i--)
                {
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Text = i.ToString();
                    }));
                    Thread.Sleep(1000);
                }



                this.Invoke((EventHandler)(delegate
                {
                    //btn_Write_MCUBIN.Enabled = false;
                    textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                    textBox1.ForeColor = Color.Black;
                    textBox1.Text = "请开始测试";
                    btn_Test_self.Enabled = true;
                    //btn_switch_video.Enabled = true;
                    //btn_beep.Enabled = true;
                    btn_search_wifi.Enabled = true;
                    btn_test_wifi.Enabled = true;
                }));
            }
            catch (Exception ex)
            {
                this.Close();
            }
        }

        public class ADASLeader_Device
        {
            public int rssi { get; set; }
            public string name;
            public string is_tested;
            public string is_pass;
        };
        List<ADASLeader_Device> ADASLeader_device = new List<ADASLeader_Device>();
        public int ADASLeader_NUM = 0;
        public int look_for_wifi_returnF = 0;

        public bool START_LOOK_FOR_WIFI_CMD()
        {
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = 00;
            msg_head.msg_type = 00;
            msg_head.resp = 0x01;
            msg_head.msg_len =
            (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head));
            msg_head.crc32 = 0;// GetCrc32(framebuff, msg_len);
            int len = System.Runtime.InteropServices.Marshal.SizeOf(msg_head);
            byte[] tmpbuff = StructToBytes(msg_head);
            //发送测试指令
            int i = 0;
            while ((commWifiTestBusy == 1) | (!commWifiTest.IsOpen))
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
                commWifiTestBusy = 1;
                if (commWifiTest.IsOpen)
                {
                    commWifiTest.Write(tmpbuff, 0, len);
                }
                else
                {
                    commWifiTest.Open();
                    commWifiTest.Write(tmpbuff, 0, len);
                }
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;
                }));

                MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
                commWifiTest.Close();
                commWifiTestBusy = 0;
                return false;
            }
            commWifiTestBusy = 0;
            return true;
        }

        public void RSSI_ORDER()
        {
            ADASLeader_device = ADASLeader_device.Distinct().ToList();
            ADASLeader_device = ADASLeader_device.OrderByDescending(ADASLeader_device => ADASLeader_device.rssi).ToList();
        }

        public void LOOK_FOR_WIFI()
        {
            int length; int cnt;
            byte[] tmp = new byte[32];

            //Thread.Sleep(5000);
            //this.Invoke((EventHandler)(delegate {listView1.Clear();}));
            //等待结果
            for (cnt = 0; cnt < 2; cnt++)
            {
                look_for_wifi_returnF = 0;
                //找WIFI
                START_LOOK_FOR_WIFI_CMD();//
                for (int i = 0; i < 4096 * 10; i++)
                    commWifiTestBuf[i] = 0;
                this.Invoke((EventHandler)(delegate { btn_search_wifi.Enabled = false; }));
                Thread.Sleep(3000);
                TimerArry[(int)timeF.lookforwifi] = 8;
                while (TimerArry[(int)timeF.lookforwifi] > 0)
                {
                    if (look_for_wifi_returnF == 1)
                    {
                        int c = 0; ; int rssi = 0; string name = null; string is_tested = null; ; string is_test_pass = null;
                        this.Invoke((EventHandler)(delegate {
                            int count = listView1.Items.Count;
                            for (int i = 0; i < count; i++)
                            {
                                ADASLeader_device.RemoveAt(0);
                                listView1.Items.Remove(listView1.Items[0]);
                            }
                        }));

                        ADASLeader_NUM = 0;
                        length = commWifiTestBuf[CommWifiTestBuffReadIndex + 4] + commWifiTestBuf[CommWifiTestBuffReadIndex + 5] * 256;
                        if (CommWifiTestBuffWriteIndex > CommWifiTestBuffReadIndex)
                        {
                            if ((CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex) < length)
                                continue;
                        }
                        else
                        {
                            if ((CommWifiTestBuffWriteIndex + 10 * 4096 - CommWifiTestBuffReadIndex) < length)
                                continue;
                        }
                        //for (int i = 0; i < length; i++)
                        //this.Invoke((EventHandler)(delegate { textBox1.Text += string.Format("{0:X000}", commWifiTestBuf[CommWifiTestBuffReadIndex + i]) + " "; }));

                        for (int i = 0; i < 30; i++)
                        {
                            for (int j = 0; j < 32; j++)
                            {
                                tmp[j] = commWifiTestBuf[CommWifiTestBuffReadIndex + 16 + 44 * i + j];
                            }
                            name = System.Text.Encoding.ASCII.GetString(tmp);
                            name = name.Substring(0, name.IndexOf('\0'));
                            if (name.Trim() == "")
                                break;
                            else if (!name.Contains("ADASLeader-"))
                                continue;
                            else
                            {
                                int blankRow = 0;
                                int j = 0;
                                for (j = 1; j < 65536; j++)
                                {
                                    ran = (Microsoft.Office.Interop.Excel.Range)ws.Cells[j, 1];
                                    if (ran.Value == null)
                                    {
                                        blankRow = j;
                                        break;
                                    }
                                }
                                for (j = 1; j < blankRow; j++)
                                {
                                    if (((Microsoft.Office.Interop.Excel.Range)ws.Cells[j, (int)column.wifiName]).Text == name.Trim())
                                        break;
                                }
                                if (j == blankRow)
                                {
                                    is_tested = "否";
                                    is_test_pass = "否";
                                }
                                else
                                {
                                    this.Invoke((EventHandler)(delegate {
                                        Microsoft.Office.Interop.Excel.Range ran1;
                                        Microsoft.Office.Interop.Excel.Range ran2;
                                        is_tested = "是";
                                        ran1 = (Microsoft.Office.Interop.Excel.Range)ws.Cells[j, (int)column.WIFI_RSSI];
                                        ran2 = (Microsoft.Office.Interop.Excel.Range)ws.Cells[j, (int)column.WIFI_FUC];
                                        if ((ran1.Text != "NO") && (ran2.Text != "NO"))
                                        {
                                            is_test_pass = "是";
                                        }
                                        else
                                        {
                                            is_test_pass = "否";
                                        }
                                    }));
                                }

                                this.Invoke((EventHandler)(delegate
                                {
                                    c = (((~((int)commWifiTestBuf[CommWifiTestBuffReadIndex + 16 + 32 + 6 + 1 + 44 * i])) & 0x000000ff) + 1);
                                    rssi = 0 - c;
                                    ADASLeader_device.Add(new ADASLeader_Device() { name = name, rssi = (0 - c), is_tested = is_tested, is_pass = is_test_pass });
                                }));
                                ADASLeader_NUM++;
                            }
                        }

                        this.Invoke((EventHandler)(delegate {
                            RSSI_ORDER();
                            ListViewItem item = new ListViewItem();
                            for (int i = 0; i < ADASLeader_device.Count; i++)
                            {
                                item = listView1.Items.Add(ADASLeader_device[i].name.Trim());
                                item.SubItems.Add(ADASLeader_device[i].rssi.ToString());
                                item.SubItems.Add(ADASLeader_device[i].is_tested.ToString());
                                item.SubItems.Add(ADASLeader_device[i].is_pass.ToString());
                                if (ADASLeader_device[i].is_tested == "是")
                                {
                                    if (ADASLeader_device[i].is_pass == "是")
                                        item.BackColor = Color.Green;
                                    else
                                        item.BackColor = Color.Red;
                                }
                                else
                                    item.BackColor = Color.Yellow;
                            }
                            if (ADASLeader_device.Count > 0)
                            {
                                CommWifiTestBuffWriteIndex = 0;
                                CommWifiTestBuffReadIndex = 0;
                                commWifiTest.DiscardInBuffer();
                            }
                            else
                            {
                                CommWifiTestBuffReadIndex = (CommWifiTestBuffReadIndex + length) % (4096 * 10);
                                commWifiTest.DiscardInBuffer();
                            }
                        }));


                        if (length > 1000)
                            break;
                    }
                    else
                        Thread.Sleep(1000);
                }
                if ((TimerArry[(int)timeF.lookforwifi] > 0) && (look_for_wifi_returnF == 1))
                {
                    break;
                }
            }
            if ((look_for_wifi_returnF <= 0) && (cnt == 1))
                MessageBox.Show("搜索超时", "请示");
            this.Invoke((EventHandler)(delegate { btn_search_wifi.Enabled = true; btn_test_wifi.Enabled = true;  }));
        }
        public int WIFITest_RUN = 0;

        public bool Start_Wifi_Test_CMD()
        {
            msg_head_t msg_head = new msg_head_t();
            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = 00;
            msg_head.msg_type = 00;
            msg_head.msg_len = 16 + 32;
            msg_head.resp = 00;
            msg_head.crc32 = 0;// GetCrc32(framebuff, msg_len);
            int len = System.Runtime.InteropServices.Marshal.SizeOf(msg_head);
            byte[] tmpbuff = StructToBytes(msg_head);
            string str; byte[] SSID;
            byte[] sendBuf = new byte[16 + 32];
            int i = 0;
            this.Invoke((EventHandler)(delegate
            {
                ListViewItem item_viewr = listView1.SelectedItems[0];
                str = item_viewr.SubItems[0].Text;
                SSID = System.Text.Encoding.ASCII.GetBytes(str);
                for (i = 0; i < tmpbuff.Length; i++)
                    sendBuf[i] = tmpbuff[i];
                for (i = 0; i < SSID.Length; i++)
                    sendBuf[i + tmpbuff.Length] = SSID[i];
                //for (i = 0; i < 16 + 32; i++)
                //textBox1.Text += string.Format("{0:X000}", sendBuf[i]) + " "; 
            }));
            try
            {
                commWifiTestBusy = 1;
                if (commWifiTest.IsOpen)
                {
                    commWifiTest.Write(sendBuf, 0, sendBuf.Length);
                }
                else
                {
                    commWifiTest.Open();
                    commWifiTest.Write(sendBuf, 0, sendBuf.Length);
                }
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;
                }));

                MessageBox.Show("ADASLeader测试工具  串口通信异常 ", e.Message);
                commWifiTest.Close();
                commWifiTestBusy = 0;
                return false;
            }
            commWifiTestBusy = 0;
            return true;
        }
        public struct TestResault
        {
            public byte[] data;
            public int returnF;
        };
        TestResault testresault;
        public string wifi_name = null;
        public int heartBeat = 0;
        public void commWifiTestDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            //commWifiTest.DiscardInBuffer();
            //搬移combuf，先整理一下combuf
            if (((CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex) == 0))
            {
                CommWifiTestBuffWriteIndex = 0;
                CommWifiTestBuffReadIndex = 0;

                for (int i = 0; i < CommWifiTestBuffWriteIndex; i++)
                {
                    commWifiTestBuf[i] = 0;
                }

            }
            /*
            else if (((CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex) > 0) & (CommWifiTestBuffReadIndex > 0))
            {
                for (int i = 0; i < CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex; i++)
                {
                    commWifiTestBuf[i] = commWifiTestBuf[i + CommWifiTestBuffReadIndex];
                    commWifiTestBuf[i + CommWifiTestBuffReadIndex] = 0;
                }

                CommWifiTestBuffReadIndex = 0;
                CommWifiTestBuffWriteIndex = CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex;
            }
            */
            //将接收到的数据写入commbuf

            Thread.Sleep(1500);
            if (!commWifiTest.IsOpen)
                return;
            int len = commWifiTest.BytesToRead;
            if ((len + CommWifiTestBuffWriteIndex) > 10 * 1024)
            {
                CommWifiTestBuffWriteIndex = 0;
                CommWifiTestBuffReadIndex = 0;
                for (int i = 0; i < 10 * 1024; i++)
                    commWifiTestBuf[i] = 0;
            }
            len = commWifiTest.Read(commWifiTestBuf, CommWifiTestBuffWriteIndex, len);
            // CommWifiTestBuffWriteIndex += len;
            CommWifiTestBuffWriteIndex = (CommWifiTestBuffWriteIndex + len) % (4096 * 10);
            //获取ACK状态
            //DealWithACK();
            string str = System.Text.Encoding.ASCII.GetString(commWifiTestBuf);

            //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n\r\n" + str; }));
            if ((str.Contains("CC3220 MQTT client Application")) || (str.Contains("heart beat")))
                commWifiTest_returnF = 1;
            str = str.Substring(CommWifiTestBuffWriteIndex - len);
            if (str.Contains("heart beat"))
            {
                heartBeat = 1;
                //str = str.Replace("heartBeat", "");
            }
            while ((CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex) >= 0x15)//Device could not connect to ADASLeader
            {
                if ((commWifiTestBuf[CommWifiTestBuffReadIndex + 0] == 0xaa) && (commWifiTestBuf[CommWifiTestBuffReadIndex + 1] == 0xaa) && (commWifiTestBuf[CommWifiTestBuffReadIndex + 2] == 0xaa) && (commWifiTestBuf[CommWifiTestBuffReadIndex + 3] == 0xaa))
                {
                    if ((commWifiTestBuf[CommWifiTestBuffReadIndex + 4] == 0x15) && (commWifiTestBuf[CommWifiTestBuffReadIndex + 7] == 0x00))
                    {
                        //this.Invoke((EventHandler)(delegate{textBox1.Text += "\r\n\r\n收到测试数据";}));
                        for (int i = 0; i < 0x15; i++)
                            testresault.data[i] = commWifiTestBuf[CommWifiTestBuffReadIndex + i];
                        testresault.returnF = 1;
                        break;
                    }
                    else if ((commWifiTestBuf[CommWifiTestBuffReadIndex + 4] == 0x38) && (commWifiTestBuf[CommWifiTestBuffReadIndex + 5] == 0x05))
                    {
                        if ((CommWifiTestBuffWriteIndex - CommWifiTestBuffReadIndex) >= 0x0538)
                            look_for_wifi_returnF = 1;
                    }
                    else
                        CommWifiTestBuffReadIndex = (CommWifiTestBuffReadIndex + 1) % (4096 * 10);
                }
                else
                    CommWifiTestBuffReadIndex = (CommWifiTestBuffReadIndex + 1) % (4096 * 10);
            }

            //CommWifiTestBuffWriteIndex = 0;
            //CommWifiTestBuffReadIndex = 0;
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
                MessageBox.Show(ex.Message);
                //textBox1.Text += "\r\n" + ex.Message;
            }
            return false;
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
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_12V.Text = tlv_reply.Voltage1.ToString(); }));        //12V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_5V.Text = tlv_reply.Voltage2.ToString(); }));         //5V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_3_3V.Text = tlv_reply.Voltage0.ToString(); }));       //3.3V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_2_5V.Text = tlv_reply.Voltage3.ToString(); }));       //2.5V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_1_8V.Text = tlv_reply.Voltage4.ToString(); }));       //1.8V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_1_2V.Text = tlv_reply.Voltage5.ToString(); }));       //1.2V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER1.Text = tlv_reply.Voltage6.ToString(); }));   //3.3V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER2.Text = tlv_reply.Voltage7.ToString(); }));   //3.3V
                        this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER3.Text = tlv_reply.Voltage8.ToString(); }));   //3.3V

                        this.Invoke((EventHandler)(delegate { textBox_static_Curent.Text = tlv_reply.Current0.ToString(); }));  //静态电流
                        this.Invoke((EventHandler)(delegate { textBox_work_Curent.Text = tlv_reply.Current1.ToString(); }));    //工作电流

                        Compare_Reply(tlv_reply, tlv_reply_max, tlv_reply_min);

                        //12V
                        if (Test_Resault.Voltage1_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_12V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_12V.BackColor = Color.Red; }));
                        //5V
                        if (Test_Resault.Voltage2_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_5V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_5V.BackColor = Color.Red; }));
                        //3.3V
                        if (Test_Resault.Voltage3_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_3_3V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_3_3V.BackColor = Color.Red; }));
                        //2.5V
                        if (Test_Resault.Voltage4_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_2_5V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_2_5V.BackColor = Color.Red; }));
                        //1.8V
                        if (Test_Resault.Voltage5_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_1_8V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_1_8V.BackColor = Color.Red; }));
                        //1.2V
                        if (Test_Resault.Voltage6_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_1_2V.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_1_2V.BackColor = Color.Red; }));
                        //3.3V
                        if (Test_Resault.Voltage7_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER1.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER1.BackColor = Color.Red; }));
                        //3.3V
                        if (Test_Resault.Voltage8_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER2.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER2.BackColor = Color.Red; }));
                        //3.3V
                        if (Test_Resault.Voltage9_OK_NOT == 1)
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER3.BackColor = Color.Green; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_Voltage_INPOWER3.BackColor = Color.Red; }));

                        if (Test_Resault.CAN_OK_NOT == 1)
                        {
                            checkBox_CAN_ok.Checked = true;
                            checkBox_CAN_failed.Checked = false;
                        }
                        else
                        {
                            checkBox_CAN_ok.Checked = false;
                            checkBox_CAN_failed.Checked = true;
                        }

                        if (PASS_FLAG == 1)
                        {
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Green;
                                textBox1.Text = "PASS";
                                btn_switch_video.Enabled = true;
                                btn_beep.Enabled = true;
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
                                btn_switch_video.Enabled = true;
                                btn_beep.Enabled = true;
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
        public void wifiselfTest()
        {
            while (true)
            {
                if (WIFITest_RUN == 1)
                {
                    this.Invoke((EventHandler)(delegate
                    { btn_test_wifi.Enabled = false; }));
                    int i = 0;
                    ListViewItem item = new ListViewItem(); int flag = 0;
                    for (i = 0; i < ADASLeader_device.Count; i++)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            item = listView1.Items[i];
                            if (item.Selected == true)
                                flag = 1;
                        }));
                        if (flag == 1)
                            break;
                    }
                    if (flag == 0)
                    {
                        WIFITest_RUN = 0;
                        MessageBox.Show("请选择要测试的WIFI", "提示");
                        this.Invoke((EventHandler)(delegate
                        { btn_test_wifi.Text = "测试WIFI"; btn_search_wifi.Enabled = true;btn_test_wifi.Enabled = true; }));
                        continue;
                    }
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                        textBox1.ForeColor = Color.Black;
                    }));
                    for (i = 0; i < 1; i++)
                    {
                        Thread.Sleep(100);
                        if (Start_Wifi_Test_CMD() == false)
                        {
                            WIFITest_RUN = 0;
                        }
                        WIFITest_RUN = 0;
                        testresault.returnF = 0;
                        TimerArry[(int)timeF.wifiTest] = 6;
                        while (TimerArry[(int)timeF.wifiTest] > 0)
                        {
                            if (testresault.returnF == 1)
                            {

                                if (((testresault.data[16] == 0xff) && (testresault.data[17] == 0xff) && testresault.data[20] == 0xff))
                                {
                                    break;
                                }

                                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n" + testresault.data[16].ToString() + " , " + testresault.data[17].ToString() + " , " + testresault.data[20].ToString() + "\r\n"; }));
                                int c = (((~((int)testresault.data[20])) & 0x000000ff) + 1);

                                if ((testresault.data[16] == 0x01) && (testresault.data[17] == 0x01))
                                {
                                    this.Invoke((EventHandler)(delegate
                                    {
                                        ListViewItem item_viewr = listView1.SelectedItems[0];
                                        wifi_name = item_viewr.SubItems[0].Text;
                                        textBox1.Text = "\r\n" + DateTime.Now.ToLocalTime().ToString() + " , 已连接WIFI" + " ," + wifi_name;
                                    }));
                                }
                                if ((testresault.data[16] == 0x01) && (testresault.data[17] == 0x01))
                                {
                                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.Text = "Yes"; }));
                                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.BackColor = Color.Green; }));
                                }
                                else
                                {
                                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.Text = "No"; }));
                                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.BackColor = Color.Red; }));
                                }

                                if (testresault.data[20] == 0xff)
                                {
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.Text = "FF"; }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.BackColor = Color.Red; }));
                                }
                                else
                                {
                                    //this.Invoke((EventHandler)(delegate { textBox_rssi.Text = "Yes"; }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.Text = (0 - c).ToString(); }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.BackColor = Color.Green; }));
                                }




                                this.Invoke((EventHandler)(delegate
                                {
                                    WIFITest_RUN = 0;
                                    //button1.Enabled = true;
                                    //button2.Enabled = true;
                                    //button3.Enabled = true;
                                }));
                                //this.Invoke((EventHandler)(delegate { btn_self_test.Text = "开始WIFI测试"; }));
                                break;
                            }
                            else
                            {

                            }
                        }
                        if ((i == 5) | ((testresault.data[16] == 0xff) && (testresault.data[17] == 0xff) && testresault.data[20] == 0xff))
                        {
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                                textBox1.ForeColor = Color.Black;
                                textBox1.Text = "未检测名称为ADASLeader的WIFI设备!";
                            }));
                            break;
                        }
                        //sendCMD(MSG_TYPE_CMD_TEST_WIFI_TEST_REQ, 0x00);
                    }
                    this.Invoke((EventHandler)(delegate
                    { btn_test_wifi.Enabled = true; }));
                    WIFITest_RUN = 0;
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
        }

        public void waitDownloadACK()   //MCU BIN文件下载  回应
        {

        }
        

        public void waitTestInfo()
        {

        }

        public void beep()
        {

        }

        public void switchVideo()
        {

        }
    }
}

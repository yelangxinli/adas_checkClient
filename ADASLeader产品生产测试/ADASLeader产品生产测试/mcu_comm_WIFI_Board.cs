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
using System.Diagnostics;


namespace ADASLeader产品生产测试
{
    public partial class WIFI_Board_Test : Form
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

        public const byte MSG_TYPE_CMD_CAN_SENSOR_TEST_REQ = 0x10;
        public const byte MSG_TYPE_CMD_CAN_SENSOR_TEST_RESP = 0x10;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_CONNECT_REQ = 0x11;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_CONNECT_RESP = 0x11;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_SRAM_REQ = 0x12;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_SRAM_RESP = 0x12;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_REQ = 0x13;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_RESP = 0x13;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_FLASH_REQ = 0x14;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_FLASH_RESP = 0x14;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FLASH_PATCH_REQ = 0x15;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FLASH_PATCH_RESP = 0x15;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ = 0x16;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_RESP = 0x16;

        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_CHUNK_REQ = 0x17;
        public const byte MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_CHUNK_RESP = 0x17;

        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_REQ = 0x18;
        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_COMM_RESP = 0x18;

        public const byte MSG_TYPE_CMD_TEST_WIFI_TEST_REQ = 0x19;
        public const byte MSG_TYPE_CMD_TEST_WIFI_TEST_RESP = 0x19;

        public const byte MSG_TYPE_CMD_TEST_WIFI_RESTART_REQ = 0x1a;
        public const byte MSG_TYPE_CMD_TEST_WIFI_RESTART_RESP = 0x1a;

        public const byte MSG_TYPE_CMD_TEST_WIFI_POWERON_REQ = 0x1b;
        public const byte MSG_TYPE_CMD_TEST_WIFI_POWERON_RESP = 0x1b;

        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_WIFI_REQ = 0x22;
        public const byte MSG_TYPE_CMD_TEST_LOOK_FOR_WIFI_RESP = 0x22;


        public int commBusy = 0;
        public int commWifiTestBusy = 0;
        public int CommBuffWriteIndex = 0;
        public int CommBuffReadIndex = 0;

        public int CommWifiTestBuffWriteIndex = 0;
        public int CommWifiTestBuffReadIndex = 0;

        public struct TestResault
        {
            public byte[] data;
            public int returnF;
        };
        TestResault testresault;

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
         
        public void RSSI_ORDER()
        {
            ADASLeader_device = ADASLeader_device.Distinct().ToList();
            ADASLeader_device = ADASLeader_device.OrderByDescending(ADASLeader_device => ADASLeader_device.rssi).ToList();
        }
        public struct tlv_firm_file_t
        {
            public UInt16 type;
            public UInt16 len;
            public UInt32 RemainLength;
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

        public bool sendCMD(byte REQ , byte Opration)
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
            if (!comm.IsOpen)
                openCOM(COMMPORT);
            //comm.DiscardInBuffer();
            comm.DiscardOutBuffer();
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
                    //comm.Open();
                    //comm.Write(tmpbuff, 0, len);
                    MessageBox.Show("串口已关闭，请重启软件");
                    return false;
                }
            }
            catch (Exception e)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n 打开串口、或者通过串口发送数据时产生异常！ " + e.Message;
                }));

                MessageBox.Show("ADASLeader测试工具"  + comm.PortName.ToString() + "串口通信异常 ", e.Message);
                comm.Close();
                commBusy = 0;
                return false;
            }
            commBusy = 0;
            return true;
        }

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
            string str;  byte[] SSID;
            byte[] sendBuf = new byte[16 + 32];
            int i = 0;
            this.Invoke((EventHandler)(delegate
            {
                ListViewItem item_viewr = listView1.SelectedItems[0];
                str = item_viewr.SubItems[0].Text;
                SSID =  System.Text.Encoding.ASCII.GetBytes(str);
                for (i = 0; i < tmpbuff.Length; i++)
                    sendBuf[i] = tmpbuff[i];
                for (i = 0 ; i < SSID.Length ;i++)
                    sendBuf[i + tmpbuff.Length] = SSID[i];
                //for (i = 0; i < 16 + 32; i++)
                    //textBox1.Text += string.Format("{0:X000}", sendBuf[i]) + " "; 
            }));
            //发送测试指令
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
                this.Invoke((EventHandler)(delegate { btn_look_for_wifi.Enabled = false; }));
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
                                        ran1 = (Microsoft.Office.Interop.Excel.Range)ws.Cells[j, (int)column.rssi];
                                        ran2 = (Microsoft.Office.Interop.Excel.Range)ws.Cells[j, (int)column.func];
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
            this.Invoke((EventHandler)(delegate { btn_look_for_wifi.Enabled = true; btn_Test_self.Enabled = true;btn_write_bin.Enabled = true; }));
        }
        public void DeviceERest()
        {
            sendCMD(MSG_TYPE_CMD_TEST_MCU_CODE_RESTART,0x00);
            while (ACKinfo.ACK_CMD_TEST_MCU_RESTART == 0) ;
            ACKinfo.ACK_CMD_TEST_MCU_RESTART = 0;

            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\n被测设备启动成功！";
            }));
        }


        public void CommFuncWIFIDownload(UInt32 RemainLen, byte[] buff, UInt16 len, bool NeedAck , byte Command)
        {
            UInt32 i, pos;
            msg_head_t msg_head = new msg_head_t();
            tlv_firm_file_t tlv_firm_file = new tlv_firm_file_t();



            //byte[] framebuff = new byte[4096 + 8];

            msg_head.start_bytes = 0xAAAAAAAA;
            msg_head.service_type = SERVICE_TYPE_CMD;
            msg_head.msg_type = Command;
            msg_head.msg_len =
            (UInt16)(System.Runtime.InteropServices.Marshal.SizeOf(msg_head) + System.Runtime.InteropServices.Marshal.SizeOf(tlv_firm_file) + len);

            byte[] framebuff = new byte[msg_head.msg_len];

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
            tlv_firm_file.RemainLength = RemainLen;

            tmpbuff = StructToBytes(tlv_firm_file);
            for (i = 0; i < System.Runtime.InteropServices.Marshal.SizeOf(tlv_firm_file); i++)
                framebuff[pos + i] = tmpbuff[i];
            pos = pos + i;

            for (i = 0; i < len; i++)
                framebuff[pos + i] = buff[i];
            pos = pos + i;
            comm.Write(framebuff, 0, msg_head.msg_len);

            
            //等待回复
            if (Command == MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_REQ)
                while (ACKinfo.ACK_CMD_WIFI_CODE_WRITE_SRAM_PATCH == 0) ;
                    ACKinfo.ACK_CMD_WIFI_CODE_WRITE_SRAM_PATCH = 0;

            if (Command == MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FLASH_PATCH_REQ)
                while (ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FLASH_PATCH == 0) ;
                    ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FLASH_PATCH = 0;

            if (Command == MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ)
                while (ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FS == 0) ;
                    ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FS = 0;

        }

        public void EnterBootMod()
        {
            sendCMD(MSG_TYPE_CMD_BOOT_IN_IAP_MODE_REQ,0x00);
        }
        public bool completedFrame(int WriteIndex , int ReadIndex)
        {
            if ((WriteIndex - ReadIndex) >= 16)
            {
                if ((commBuf[ReadIndex] == 0xAA) & (commBuf[ReadIndex + 1] == 0xAA) & (commBuf[ReadIndex + 2] == 0xAA) & (commBuf[ReadIndex + 3] == 0xAA))
                {
                    if ((WriteIndex - ReadIndex) >= commBuf[ReadIndex + 4] + commBuf[ReadIndex + 5] * 256)
                        return true;
                    else
                        return false;
                }
                else
                {
                    if(comm.PortName.ToString() == "COM22")
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Text += comm.PortName.ToString() + "\r\n从测试板 发送 给PC的数据 可能存在丢失！";
                    }));
                    for (int start = 0; start < WriteIndex; start++)
                    {
                        if ((commBuf[start] == 0xAA) & (commBuf[start + 1] == 0xAA) & (commBuf[start + 2] == 0xAA) & (commBuf[start + 3] == 0xAA))
                        {
                            ReadIndex = start;
                        }
                    }
                    return false;
                }
            }
            else
                return false;

        }
        public static bool IfReceivingData;
        public static bool PortIsCloseing;
        public void commDataReceivedHandler(object sender, SerialDataReceivedEventArgs e)
        {
            IfReceivingData = true;
            if (PortIsCloseing)
            {
                return;
            }
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

            string str = System.Text.Encoding.ASCII.GetString(commBuf);
            if (str.Contains("system startup"))
                MessageBox.Show("测试板已经重启！");

            //获取ACK状态
            DealWithACK();

            IfReceivingData = false;
        }


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

        public void DealWithACK()
        {
            byte serviceType; byte message; int valid = 1;
            //分析commbuf
            if (completedFrame(CommBuffWriteIndex , CommBuffReadIndex))
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
                        case MSG_TYPE_CMD_TEST_WIFI_TEST_REQ:
                            {
                                msg_head_t msg_head = new msg_head_t();
                                msg_head = (msg_head_t)BytesToStruct(commBuf, typeof(msg_head_t));
                                byte[] tmpbuff = new byte[1088];
                                for (int i = 0; i < msg_head.msg_len - 16; i++)
                                {
                                    tmpbuff[i] = commBuf[i + 16];
                                }
                                tlv_reply = (tlv_reply_t)BytesToStruct(tmpbuff, typeof(tlv_reply_t));
                                ACKinfo.ACK_CMD_WIFI_Test_Resualt = 1;
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
                        case MSG_TYPE_CMD_CAN_SENSOR_TEST_REQ:
                            ACKinfo.ACK_CMD_CANSENSOR = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_CONNECT_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_CONNECT = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_SRAM_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_SRAM = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_SRAM_PATCH = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_FLASH_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_FLASH = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FLASH_PATCH_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FLASH_PATCH = 1;
                            break;
                        case MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ:
                            ACKinfo.ACK_CMD_WIFI_CODE_WRITE_FS = 1;
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


        public void downLoadWIFIBIN()
        {
            //if (commWifiTest.IsOpen)
              //commWifiTest.Close();

            //SerialPort.GetPortNames
            //string[] ports = SerialPort.GetPortNames();
            //Array.Sort(ports);
            //foreach(string port in ports)
            //拉高Rst
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ, 0x01);   //power on wifi
            Thread.Sleep(5000);
            //运行ImageProgramming

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo("ImageProgramming.exe");
                if (commWifiWriteBIN != null)
                {
                    startInfo.WindowStyle = ProcessWindowStyle.Normal;
                    startInfo.Arguments = "-p " + commWifiWriteBIN.Substring(3) + " -i WIFI_product.ucf";
                    //startInfo.Arguments = "-p 6 -i WIFI_test.ucf";
                    Process.Start(startInfo);
                }
                else
                {
                    this.Invoke((EventHandler)(delegate { textBox1.Text += "请检查串口是否连接，并重新启动测试软件。"; }));
                }

                
                /*
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "ImageProgramming.exe";
                p.StartInfo.Arguments = "-p " + commWifiWriteBIN.Substring(3) + " -i WIFI_product.ucf";//启动参数   
                                                                                                      
                //p.StartInfo.Arguments = "-p " + commWifiTest.PortName.ToString().Substring(3) + " -i WIFI_test.ucf";//启动参数  
                this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n002"; }));
                p.Start();*/
            }
            catch (Exception e0)
            {
                MessageBox.Show("启动应用程序时出错！原因：" + e0.Message);
            }
            this.Invoke((EventHandler)(delegate{ btn_write_bin.Enabled = false; }));
            Thread.Sleep(5000);
            //sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ, 0x00);   //reset wifi
            UInt32 ImageProgramming_RUN_FLAG = 0;
            Process[] localAll = Process.GetProcesses();//ImageProgramming
            for (int i = 0; i <=5; i++)
            {
                foreach (Process A in localAll)
                    if (A.ProcessName.ToString() == "ImageProgramming")
                    {
                        ImageProgramming_RUN_FLAG = 1;
                        break;
                    }
                if (ImageProgramming_RUN_FLAG == 1)
                {
                    //power off wifi
                    //sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ, 0x00);
                    //Thread.Sleep(500);
                    //power on wifi
                    //sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ, 0x01);
                    Thread.Sleep(2000);
                    sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ, 0x00);   //reset wifi
                    break;
                }
                if (i == 5)
                    return;
                else
                    Thread.Sleep(1000);
            }

            ImageProgramming_RUN_FLAG = 1;
            Thread.Sleep(500);
            while (ImageProgramming_RUN_FLAG == 1)
            {
                localAll = Process.GetProcesses();//ImageProgramming
                foreach (Process A in localAll)
                {
                    if (A.ProcessName.ToString() == "ImageProgramming")
                    {
                        ImageProgramming_RUN_FLAG = 1;
                        break;
                    }
                    else
                    {
                        ImageProgramming_RUN_FLAG = 0;
                        continue;
                    }
                }
                Array.Clear(localAll,0, localAll.Length);
                Thread.Sleep(500);
            }
            
            //Restart WIFI
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_RESTART_REQ, 0x00);


            Thread.Sleep(6000);

            this.Invoke((EventHandler)(delegate {
                btn_look_for_wifi.Enabled = true;
                btn_write_bin.Enabled = true;
                textBox1.Text = "请继续\r\n1:如果您手动关闭了烧录程序，请重新扫码烧录\r\n2:如果烧录正常完成，请点击开始测试！";
            }));


            /*
            if (!commWifiTest.IsOpen)
            {
                commWifiTest.Close();
            }

            
            if (openWifiTestCOM("COM23"))
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n测试串口打开成功！";
                }));
            }
            else
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\n测试串口打开失败！";
                }));
            }
            
            BinaryReader br; FileStream fs;
            //连接
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_CONNECT_REQ);
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nConnecting";
            }));
            //等待回复
            while (ACKinfo.ACK_CMD_WIFI_CODE_CONNECT == 0) ;
            ACKinfo.ACK_CMD_WIFI_CODE_CONNECT = 0;

            if (commBuf[0x07] == 1)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\nConnect failed";
                }));
                return;
            }
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nConnect success";
            }));


            //擦除 SRAM
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_SRAM_REQ);
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nErasing SRAM";
            }));
            //等待回复
            while (ACKinfo.ACK_CMD_WIFI_CODE_ERASE_SRAM == 0) ;
            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_SRAM = 0;

            if (commBuf[0x07] == 1)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\nErase SRAM failed";
                }));
                return;
            }

            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nErase SRAM success";
            }));
            

            //下载 sram patch
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nDownload SRAM PATCH";
            }));
            fs = new FileStream("BTL_ram.ptc", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            br = new BinaryReader(fs);
            byte[] binbuf = new byte[4080];
            while (true)
            {
                binbuf = br.ReadBytes(4080);
                if (binbuf.Length == 0)         //finished
                {
                    CommFuncWIFIDownload(0, binbuf, 0, true, MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_REQ);
                    if (commBuf[0x07] == 1)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Text += "\r\nDownload SRAM PATCH failed";
                        }));
                        fs.Close();
                        return;
                    }
                    break;
                }
                else
                {
                    CommFuncWIFIDownload(0, binbuf, (UInt16)binbuf.Length, true, MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_SRAM_PATCH_REQ);
                    if (commBuf[0x07] == 1)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Text += "\r\nDownload SRAM PATCH failed";
                        }));
                        fs.Close();
                        return;
                    }
                }
            }
            //擦除 flash
            sendCMD(MSG_TYPE_CMD_TEST_WIFI_CODE_ERASE_FLASH_REQ);
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nErasing FLASH";
            }));
            //等待回复
            while (ACKinfo.ACK_CMD_WIFI_CODE_ERASE_FLASH == 0) ;
            ACKinfo.ACK_CMD_WIFI_CODE_ERASE_FLASH = 0;
            if (commBuf[0x07] == 1)
            {
                this.Invoke((EventHandler)(delegate
                {
                    textBox1.Text += "\r\nErasing FLASH failed";
                }));
                fs.Close();
                return;
            }

            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nErasing FLASH success";
            }));

            //下载 flash patch
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nDownload FLASH PATCH";
            }));
            fs = new FileStream("BTL_sflash.ptc", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            br = new BinaryReader(fs);
            binbuf = new byte[4080];
            while (true)
            {
                binbuf = br.ReadBytes(4080);
                if (binbuf.Length == 0)         //finished
                    break;
                else
                {
                    CommFuncWIFIDownload(0, binbuf, (UInt16)binbuf.Length, true, MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FLASH_PATCH_REQ);
                    if (commBuf[0x07] == 1)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Text += "\r\nDownload FLASH PATCH failed";
                        }));
                        fs.Close();
                        return;
                    }
                }
            }


            //写 image
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\nDownloading IMAGE\r\n";
            }));
            fs = new FileStream("cc3220S_uniflash_programming.logicdata", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            br = new BinaryReader(fs);
            binbuf = new byte[4096];
            while (true)
            {
                binbuf = br.ReadBytes(4096);
                if (binbuf.Length == 0)         //finished
                    break;
                else
                {
                    CommFuncWIFIDownload(0, binbuf, (UInt16)binbuf.Length, true, MSG_TYPE_CMD_TEST_WIFI_CODE_WRITE_FS_REQ);
                    if (commBuf[0x07] == 1)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Text += "\r\nDownload IMAGE failed";
                        }));
                        fs.Close();
                        return;
                    }
                }
            }
            return;

            */
        }




        public void selfTest()
        {
            while(true)
            {
                if (WIFITest_RUN == 1)
                {
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
                        MessageBox.Show("请选择要测试的WIFI" ,"提示");
                        this.Invoke((EventHandler)(delegate
                        { btn_Test_self.Text = "开始测试"; }));
                        continue;
                    }
                    this.Invoke((EventHandler)(delegate
                    {
                        textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                        textBox1.ForeColor = Color.Black;
                    }));
                    this.Invoke((EventHandler)(delegate{textBox1.Text = "开始WIFI测试";}));
                    this.Invoke((EventHandler)(delegate { textBox_rssi.Text = ""; }));
                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.Text = ""; }));
                    this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.BackColor = Color.Gray; }));
                    this.Invoke((EventHandler)(delegate { textBox_rssi.BackColor = Color.Gray; }));
                    for (i = 0; i < 1; i++)
                    {
                        Thread.Sleep(100);
                        if (Start_Wifi_Test_CMD() == false)
                        {
                            WIFITest_RUN = 0;
                            this.Invoke((EventHandler)(delegate { btn_Test_self.Text = "开始测试"; }));
                        }
                        WIFITest_RUN = 0;
                        testresault.returnF = 0;
                        TimerArry[(int)timeF.selfTest] = 8;
                        while (TimerArry[(int)timeF.selfTest] > 0)
                        {
                            if (testresault.returnF == 1)
                            {

                                if (((testresault.data[16] == 0xff) && (testresault.data[17] == 0xff) && testresault.data[20] == 0xff))
                                {
                                    break;
                                }

                                //this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\n" + testresault.data[16].ToString() + " , " + testresault.data[17].ToString() + " , " + testresault.data[20].ToString() + "\r\n"; }));
                                int c = (((~((int)testresault.data[20])) & 0x000000ff) + 1);

                                if (testresault.data[16] == 0x01)
                                    this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nTCP OK!"; }));
                                if (testresault.data[17] == 0x01)
                                    this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nUDP OK!"; }));
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

                                if (testresault.data[20] == 0)
                                {
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.Text = (0 - c).ToString(); }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.BackColor = Color.Red; }));
                                }
                                else
                                {
                                    //this.Invoke((EventHandler)(delegate { textBox_rssi.Text = "Yes"; }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.Text = (0 - c).ToString(); }));
                                    this.Invoke((EventHandler)(delegate { textBox_rssi.BackColor = Color.Green; }));
                                }


                                this.Invoke((EventHandler)(delegate { textBox1.Text += "\r\nRSSI = -" + (((~((int)testresault.data[20])) & 0x000000ff) + 1).ToString(); }));


                                if ((testresault.data[20] != 0xff) && (testresault.data[16] == 0x01) && (testresault.data[17] == 0x01))
                                {
                                    this.Invoke((EventHandler)(delegate
                                    {
                                        textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                        textBox1.ForeColor = Color.Green;
                                        textBox1.Text = "PASS";
                                    }));
                                }
                                else
                                {
                                    this.Invoke((EventHandler)(delegate
                                    {
                                        textBox1.Font = new Font(textBox1.Font.FontFamily, 100, textBox1.Font.Style);
                                        textBox1.ForeColor = Color.Red;
                                        textBox1.Text = "NO";
                                    }));
                                }
                                this.Invoke((EventHandler)(delegate { btn_submit.Enabled = true; }));
                                WIFITest_RUN = 0;
                                this.Invoke((EventHandler)(delegate { btn_Test_self.Text = "开始测试";btn_look_for_wifi.Enabled = true;btn_write_bin.Enabled = true; }));
                                break;
                            }
                            else
                            {

                            }
                        }
                        if (((testresault.data[16] == 0xff) && (testresault.data[17] == 0xff) && testresault.data[20] == 0xff))
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
                    if (TimerArry[(int)timeF.selfTest] <= 0)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                            textBox1.ForeColor = Color.Black;
                            textBox1.Text = "测试超时";
                        }));
                    }
                    WIFITest_RUN = 0;
                    this.Invoke((EventHandler)(delegate { btn_Test_self.Text = "开始测试"; btn_look_for_wifi.Enabled = true; btn_write_bin.Enabled = true; }));
                    /*
                    if ((!((i == 5) | ((testresault.data[16] == 0xff) && (testresault.data[17] == 0xff) && testresault.data[20] == 0xff))) && (heartBeat == 0))
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Font = new Font(textBox1.Font.FontFamily, 12, textBox1.Font.Style);
                            textBox1.ForeColor = Color.Black;
                            textBox1.Text = "未检测到心跳信号超时！\r\n重启WIFI测试模块中。。。";
                        }));
                        //sendCMD(MSG_TYPE_CMD_TEST_WIFI_POWERON_REQ, 0x00);
                        //btn_Test_self.Enabled = false;
                        for (i = 5; i >= 0; i--)
                        {
                            Thread.Sleep(1000);
                            this.Invoke((EventHandler)(delegate
                            {
                                textBox1.Text = "\r" + i.ToString();
                            }));
                        }
                    }*/

                }
                else
                    Thread.Sleep(1000);
                
            }
            
            btn_Test_self.Enabled = true;
            /*
            TimerArry[(int)timeF.selfTest] = 2;
            while (TimerArry[(int)timeF.selfTest] > 0) ;  //等待2S
            ACKinfo.ACK_CMD_WIFI_Test_Resualt = 0;
            //获取自测结果(下发自测指令、等待ACK)
            for (int i = 5; i > 0; i--)
            {
                //发送 请求测试结果指令
                sendCMD(MSG_TYPE_CMD_TEST_WIFI_TEST_REQ, 0x00);
                TimerArry[(int)timeF.selfTest] = 10;
                while (TimerArry[(int)timeF.selfTest] > 0)  //等待10S钟
                {
                    if (ACKinfo.ACK_CMD_WIFI_Test_Resualt == 1)
                    {
                        //显示测试结果，电压电流
                        this.Invoke((EventHandler)(delegate { textBox_static_Curent.Text = tlv_reply.Current1.ToString(); }));  //静态电流
                        this.Invoke((EventHandler)(delegate { textBox_work_voltage.Text = tlv_reply.Current0.ToString(); }));    //工作电流
                        this.Invoke((EventHandler)(delegate { textBox_work_current.Text = tlv_reply.Voltage1.ToString(); }));

                        //显示测试结果，WIFI功能性能
                        if(tlv_reply.TestResult0 == 1)
                            this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.Text = "PASS"; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_WIFI_FUC.Text = "NO"; }));

                        if (tlv_reply.TestResult0 != 0)
                            this.Invoke((EventHandler)(delegate { textBox_WIFI_xingNeng.Text = "PASS ," + tlv_reply.TestResult0.ToString() + "DB"; }));
                        else
                            this.Invoke((EventHandler)(delegate { textBox_WIFI_xingNeng.Text = "NO ," + tlv_reply.TestResult0.ToString() + "DB"; }));
                    }
                    Compare_Reply(tlv_reply, tlv_reply_max, tlv_reply_min);
                    if (PASS_FLAG == 1)
                    {
                        this.Invoke((EventHandler)(delegate
                        {
                            textBox1.Text += "\r\n\r\n通过自测！";
                        }));
                        return;
                    }
                }
            }
            this.Invoke((EventHandler)(delegate
            {
                textBox1.Text += "\r\n\r\n未通过自测！";
            }));
            */
        }
        
    }
}

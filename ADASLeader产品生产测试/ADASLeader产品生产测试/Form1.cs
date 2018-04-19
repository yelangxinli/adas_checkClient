using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO.Ports;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Office.Interop;
using System.Diagnostics;
namespace ADASLeader产品生产测试
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            timer1.Interval = 1000;
            timer1.Start();
            ADAS_PROCESS = Process.GetProcesses();
        }
        public Process[] ADAS_PROCESS;
        private void btn_ScanBar_Click(object sender, EventArgs e)
        {
            textBox_OP_CODE.Clear();
            textBox_OP_CODE.Focus();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
        }

        private void btn_Start_Test_Click(object sender, EventArgs e)
        {
            switch (textBox_OP_CODE.Text.ToString())
            {
                case "888001":    //主板测试
                    textBox1.Text += "\r\n主板测试";
                    MainBoard_PCBATest MainBoardPCBATest = new MainBoard_PCBATest();
                    MainBoardPCBATest.ShowDialog();
                    break;
                case "888002":    //主机模块测试
                    textBox1.Text += "\r\n主板模块测试";
                    EndProduct mainboard_moudle = new EndProduct("mainboard moudle");
                    mainboard_moudle.ShowDialog();
                    break;
                case "888003":    //摄像头板测试
                    textBox1.Text += "\r\n摄像头板测试";
                    CameraBoard_PCBATest CameraBoardPCBATest = new CameraBoard_PCBATest(textBox_OP_CODE.Text.ToString());
                    CameraBoardPCBATest.ShowDialog();
                    break;
                case "888004":    //摄像头模块测试
                    textBox1.Text += "\r\n摄像头模块测试";
                    CameraMoudle_EndProduct_PCBATest CameraMoudleEndProductTest = new CameraMoudle_EndProduct_PCBATest(textBox_OP_CODE.Text.ToString());
                    CameraMoudleEndProductTest.ShowDialog();
                    break;
                case "888005":    //DVR板测试
                    textBox1.Text += "\r\nDVR板测试";
                    DVRBoard_PCBATest DVRBoardTest = new DVRBoard_PCBATest(textBox_OP_CODE.Text.ToString());
                    DVRBoardTest.ShowDialog();
                    break;
                case "888006":    //CAN-SENSOR测试
                    textBox1.Text += "\r\nCAN-SENSOR测试";
                    CAN_Sensor_Test CANSensorTest = new CAN_Sensor_Test();
                    CANSensorTest.ShowDialog();
                    break;
                case "888007":    //WIFI模块测试
                    textBox1.Text += "\r\nWIFI模块测试";
                    WIFI_Board_Test wifi_board_Test = new WIFI_Board_Test();
                    wifi_board_Test.ShowDialog();
                    break;
                case "888008":    //WIFI模块测试
                    textBox1.Text += "\r\n成品测试";
                    EndProduct endproduct = new EndProduct("end product");
                    endproduct.ShowDialog();
                    break;
                default:
                    MessageBox.Show("请输入正确的口令","提示");
                    break;
            }
        }

        private void btn_MainBoard_PCBA_Test_Click(object sender, EventArgs e)
        {
            MainBoard_PCBATest MainBoardPCBATest = new MainBoard_PCBATest();
            MainBoardPCBATest.ShowDialog();
        }

        private void btn_MainBoardMoudel_PCBA_Test_Click(object sender, EventArgs e)
        {
            MainBoardMoudle_PCBATest0 MainBoardMoudeelPCBATest = new MainBoardMoudle_PCBATest0(textBox_OP_CODE.Text.ToString());
            MainBoardMoudeelPCBATest.ShowDialog();
        }

        private void btn_CameraBoard_PCBA_Test_Click(object sender, EventArgs e)
        {
            CameraBoard_PCBATest CameraBoardPCBATest = new CameraBoard_PCBATest(textBox_OP_CODE.Text.ToString());
            CameraBoardPCBATest.ShowDialog();
        }

        private void btn_DVRBoard_PCBA_Test_Click(object sender, EventArgs e)
        {
            DVRBoard_PCBATest DVRBoardTest = new DVRBoard_PCBATest(textBox_OP_CODE.Text.ToString());
            DVRBoardTest.ShowDialog();
        }

        private void btn_CameraEndProduct_Test_Click(object sender, EventArgs e)
        {
            CameraMoudle_EndProduct_PCBATest CameraMoudleEndProductTest = new CameraMoudle_EndProduct_PCBATest(textBox_OP_CODE.Text.ToString());
            CameraMoudleEndProductTest.ShowDialog();
        }

        private void btn_CAN_Sensor_Test_Click(object sender, EventArgs e)
        {
            CAN_Sensor_Test CANSensorTest = new CAN_Sensor_Test();
            CANSensorTest.ShowDialog();
        }

        private void btn_ADASLeader_EndProduct_Test_Click(object sender, EventArgs e)
        {
            WIFI_Board_Test wifi_board_Test = new WIFI_Board_Test();
            wifi_board_Test.ShowDialog();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {

                foreach (Process A in ADAS_PROCESS)
                {
                    if (A.ProcessName.ToString() == "ADASLeader产品生产测试")
                        A.Kill();
                    //textBox1.Text += "\r\n" + A.ProcessName.ToString();
                }
                
                //this.Close();

            }
            catch (StackOverflowException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void textBox_OP_CODE_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Char.IsNumber(e.KeyChar)) || (e.KeyChar == (char)8))
            {
                //e.Handled = true;
            }
            else
            {
                e.Handled = true;
            }
            if ((textBox_OP_CODE.Text.ToString().Length >= 6) && (e.KeyChar != (char)8))
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = (char)Keys.None;
        }

        private void textBox_OP_CODE_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

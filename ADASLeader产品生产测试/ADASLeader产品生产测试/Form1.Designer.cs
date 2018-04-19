namespace ADASLeader产品生产测试
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textBox_OP_CODE = new System.Windows.Forms.TextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.btn_Start_Test = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label15 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label16 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox_OP_CODE
            // 
            this.textBox_OP_CODE.Font = new System.Drawing.Font("宋体", 68F);
            this.textBox_OP_CODE.Location = new System.Drawing.Point(12, 12);
            this.textBox_OP_CODE.Multiline = true;
            this.textBox_OP_CODE.Name = "textBox_OP_CODE";
            this.textBox_OP_CODE.Size = new System.Drawing.Size(560, 102);
            this.textBox_OP_CODE.TabIndex = 7;
            this.textBox_OP_CODE.Text = "输入口令";
            this.textBox_OP_CODE.TextChanged += new System.EventHandler(this.textBox_OP_CODE_TextChanged);
            this.textBox_OP_CODE.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_OP_CODE_KeyPress);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // btn_Start_Test
            // 
            this.btn_Start_Test.BackColor = System.Drawing.Color.LightGreen;
            this.btn_Start_Test.Font = new System.Drawing.Font("宋体", 42F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Start_Test.Location = new System.Drawing.Point(591, 12);
            this.btn_Start_Test.Name = "btn_Start_Test";
            this.btn_Start_Test.Size = new System.Drawing.Size(343, 102);
            this.btn_Start_Test.TabIndex = 10;
            this.btn_Start_Test.Text = "开始测试";
            this.btn_Start_Test.UseVisualStyleBackColor = false;
            this.btn_Start_Test.Click += new System.EventHandler(this.btn_Start_Test_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.DarkSalmon;
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(12, 134);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(922, 293);
            this.panel1.TabIndex = 11;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Silver;
            this.panel2.Controls.Add(this.label15);
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(922, 52);
            this.panel2.TabIndex = 13;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("宋体", 30F);
            this.label15.ForeColor = System.Drawing.Color.Red;
            this.label15.Location = new System.Drawing.Point(261, 6);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(357, 40);
            this.label15.TabIndex = 12;
            this.label15.Text = "测   试   口   令";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 30F);
            this.label4.ForeColor = System.Drawing.SystemColors.Info;
            this.label4.Location = new System.Drawing.Point(559, 73);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(257, 40);
            this.label4.TabIndex = 0;
            this.label4.Text = "主机模块测试";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("宋体", 30F);
            this.label12.ForeColor = System.Drawing.SystemColors.Info;
            this.label12.Location = new System.Drawing.Point(559, 185);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(297, 40);
            this.label12.TabIndex = 0;
            this.label12.Text = "CAN-SENSOR测试";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 30F);
            this.label8.ForeColor = System.Drawing.SystemColors.Info;
            this.label8.Location = new System.Drawing.Point(559, 128);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(297, 40);
            this.label8.TabIndex = 0;
            this.label8.Text = "摄像头模块测试";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("宋体", 30F);
            this.label18.ForeColor = System.Drawing.SystemColors.Info;
            this.label18.Location = new System.Drawing.Point(559, 240);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(177, 40);
            this.label18.TabIndex = 0;
            this.label18.Text = "成品测试";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("宋体", 30F);
            this.label14.ForeColor = System.Drawing.SystemColors.Info;
            this.label14.Location = new System.Drawing.Point(145, 240);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(257, 40);
            this.label14.TabIndex = 0;
            this.label14.Text = "WIFI模块测试";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("宋体", 30F);
            this.label10.ForeColor = System.Drawing.SystemColors.Info;
            this.label10.Location = new System.Drawing.Point(145, 186);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(197, 40);
            this.label10.TabIndex = 0;
            this.label10.Text = "DVR板测试";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 30F);
            this.label6.ForeColor = System.Drawing.SystemColors.Info;
            this.label6.Location = new System.Drawing.Point(145, 129);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(257, 40);
            this.label6.TabIndex = 0;
            this.label6.Text = "摄像头板测试";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 30F);
            this.label2.ForeColor = System.Drawing.SystemColors.Info;
            this.label2.Location = new System.Drawing.Point(145, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(177, 40);
            this.label2.TabIndex = 0;
            this.label2.Text = "主板测试";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 30F);
            this.label3.Location = new System.Drawing.Point(423, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(137, 40);
            this.label3.TabIndex = 0;
            this.label3.Text = "888002";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("宋体", 30F);
            this.label11.Location = new System.Drawing.Point(423, 186);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(137, 40);
            this.label11.TabIndex = 0;
            this.label11.Text = "888006";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 30F);
            this.label7.Location = new System.Drawing.Point(423, 129);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(137, 40);
            this.label7.TabIndex = 0;
            this.label7.Text = "888004";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("宋体", 30F);
            this.label17.Location = new System.Drawing.Point(423, 241);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(137, 40);
            this.label17.TabIndex = 0;
            this.label17.Text = "888008";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("宋体", 30F);
            this.label13.Location = new System.Drawing.Point(9, 241);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(137, 40);
            this.label13.TabIndex = 0;
            this.label13.Text = "888007";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("宋体", 30F);
            this.label9.Location = new System.Drawing.Point(9, 187);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(137, 40);
            this.label9.TabIndex = 0;
            this.label9.Text = "888005";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 30F);
            this.label5.Location = new System.Drawing.Point(9, 130);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(137, 40);
            this.label5.TabIndex = 0;
            this.label5.Text = "888003";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 30F);
            this.label1.Location = new System.Drawing.Point(9, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(137, 40);
            this.label1.TabIndex = 0;
            this.label1.Text = "888001";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("宋体", 12F);
            this.textBox1.Location = new System.Drawing.Point(0, 34);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(922, 217);
            this.textBox1.TabIndex = 12;
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.YellowGreen;
            this.panel3.Controls.Add(this.label16);
            this.panel3.Controls.Add(this.textBox1);
            this.panel3.Location = new System.Drawing.Point(12, 433);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(922, 254);
            this.panel3.TabIndex = 13;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("宋体", 18F);
            this.label16.Location = new System.Drawing.Point(3, 6);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(58, 24);
            this.label16.TabIndex = 13;
            this.label16.Text = "提示";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(946, 687);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btn_Start_Test);
            this.Controls.Add(this.textBox_OP_CODE);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ADASLeader生产测试软件-主界面";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox textBox_OP_CODE;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button btn_Start_Test;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
    }
}


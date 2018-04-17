namespace ADASLeader产品生产测试
{
    partial class WIFIPCBATest
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btn_write_bin = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_openCOM = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_select_ComPort = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btn_function_test = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.btn_performance_test = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.btn_submit = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel4.Controls.Add(this.label7);
            this.panel4.Controls.Add(this.textBox1);
            this.panel4.Location = new System.Drawing.Point(332, 1);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(454, 410);
            this.panel4.TabIndex = 15;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(3, 7);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(52, 21);
            this.label7.TabIndex = 7;
            this.label7.Text = "提示";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(3, 32);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(451, 375);
            this.textBox1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel3.Controls.Add(this.textBox2);
            this.panel3.Controls.Add(this.checkBox4);
            this.panel3.Controls.Add(this.checkBox3);
            this.panel3.Controls.Add(this.label5);
            this.panel3.Controls.Add(this.label4);
            this.panel3.Controls.Add(this.label6);
            this.panel3.Location = new System.Drawing.Point(332, 417);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(454, 187);
            this.panel3.TabIndex = 14;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(170, 65);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(60, 16);
            this.checkBox4.TabIndex = 11;
            this.checkBox4.Text = "不正常";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(117, 65);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(48, 16);
            this.checkBox3.TabIndex = 11;
            this.checkBox3.Text = "正常";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(4, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(104, 16);
            this.label5.TabIndex = 10;
            this.label5.Text = "WIFI信号强度";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(36, 63);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 16);
            this.label4.TabIndex = 10;
            this.label4.Text = "WIFI功能";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(5, 5);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(98, 21);
            this.label6.TabIndex = 9;
            this.label6.Text = "测试结果";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Info;
            this.panel1.Controls.Add(this.btn_openCOM);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.comboBox_select_ComPort);
            this.panel1.Controls.Add(this.btn_submit);
            this.panel1.Controls.Add(this.btn_performance_test);
            this.panel1.Controls.Add(this.btn_function_test);
            this.panel1.Controls.Add(this.btn_write_bin);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(3, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(323, 601);
            this.panel1.TabIndex = 13;
            // 
            // btn_write_bin
            // 
            this.btn_write_bin.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_write_bin.Location = new System.Drawing.Point(14, 205);
            this.btn_write_bin.Name = "btn_write_bin";
            this.btn_write_bin.Size = new System.Drawing.Size(281, 50);
            this.btn_write_bin.TabIndex = 5;
            this.btn_write_bin.Text = "烧录程序";
            this.btn_write_bin.UseVisualStyleBackColor = true;
            this.btn_write_bin.Click += new System.EventHandler(this.btn_write_bin_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 42F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Coral;
            this.label1.Location = new System.Drawing.Point(72, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 56);
            this.label1.TabIndex = 3;
            this.label1.Text = "测试";
            // 
            // btn_openCOM
            // 
            this.btn_openCOM.Location = new System.Drawing.Point(160, 115);
            this.btn_openCOM.Name = "btn_openCOM";
            this.btn_openCOM.Size = new System.Drawing.Size(153, 23);
            this.btn_openCOM.TabIndex = 9;
            this.btn_openCOM.Text = "打开串口";
            this.btn_openCOM.UseVisualStyleBackColor = true;
            this.btn_openCOM.Click += new System.EventHandler(this.btn_openCOM_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(12, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(295, 35);
            this.label2.TabIndex = 8;
            this.label2.Text = "第一步：选择串口";
            // 
            // comboBox_select_ComPort
            // 
            this.comboBox_select_ComPort.FormattingEnabled = true;
            this.comboBox_select_ComPort.Location = new System.Drawing.Point(7, 115);
            this.comboBox_select_ComPort.Name = "comboBox_select_ComPort";
            this.comboBox_select_ComPort.Size = new System.Drawing.Size(135, 20);
            this.comboBox_select_ComPort.TabIndex = 7;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(12, 167);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(120, 35);
            this.label8.TabIndex = 8;
            this.label8.Text = "第二步";
            // 
            // btn_function_test
            // 
            this.btn_function_test.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_function_test.Location = new System.Drawing.Point(12, 322);
            this.btn_function_test.Name = "btn_function_test";
            this.btn_function_test.Size = new System.Drawing.Size(281, 50);
            this.btn_function_test.TabIndex = 5;
            this.btn_function_test.Text = "WIFI功能测试";
            this.btn_function_test.UseVisualStyleBackColor = true;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(10, 284);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(120, 35);
            this.label9.TabIndex = 8;
            this.label9.Text = "第三步";
            // 
            // btn_performance_test
            // 
            this.btn_performance_test.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_performance_test.Location = new System.Drawing.Point(12, 429);
            this.btn_performance_test.Name = "btn_performance_test";
            this.btn_performance_test.Size = new System.Drawing.Size(281, 50);
            this.btn_performance_test.TabIndex = 5;
            this.btn_performance_test.Text = "WIFI性能测试";
            this.btn_performance_test.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(10, 391);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(120, 35);
            this.label10.TabIndex = 8;
            this.label10.Text = "第四步";
            // 
            // btn_submit
            // 
            this.btn_submit.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_submit.Location = new System.Drawing.Point(12, 538);
            this.btn_submit.Name = "btn_submit";
            this.btn_submit.Size = new System.Drawing.Size(281, 50);
            this.btn_submit.TabIndex = 5;
            this.btn_submit.Text = "提交测试结果";
            this.btn_submit.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.Location = new System.Drawing.Point(10, 500);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(120, 35);
            this.label11.TabIndex = 8;
            this.label11.Text = "第五步";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(113, 117);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(117, 21);
            this.textBox2.TabIndex = 12;
            // 
            // WIFIPCBATest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(786, 603);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "WIFIPCBATest";
            this.Text = "WIFIPCBATest";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.WIFIPCBATest_FormClosing);
            this.Load += new System.EventHandler(this.WIFIPCBATest_Load);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btn_write_bin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btn_openCOM;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox_select_ComPort;
        private System.Windows.Forms.Button btn_submit;
        private System.Windows.Forms.Button btn_performance_test;
        private System.Windows.Forms.Button btn_function_test;
        private System.Windows.Forms.Timer timer1;
    }
}
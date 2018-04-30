namespace Web_Crawler
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.Get_day_Data = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.EXIT = new System.Windows.Forms.Button();
            this.Get_Calendar = new System.Windows.Forms.Button();
            this.dateTimePicker_day = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_month_1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_month_2 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Upload_Calendar = new System.Windows.Forms.Button();
            this.Delete_Table = new System.Windows.Forms.Button();
            this.Create_Table = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.Clear_Table = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.Auto_upload_Timer = new System.Windows.Forms.Timer(this.components);
            this.Get_Announcement = new System.Windows.Forms.Button();
            this.Upload_Text = new System.Windows.Forms.TextBox();
            this.Create_Text = new System.Windows.Forms.TextBox();
            this.Delete_Text = new System.Windows.Forms.TextBox();
            this.Clear_Text = new System.Windows.Forms.TextBox();
            this.Upload_Announcement = new System.Windows.Forms.Button();
            this.bt_TEST = new System.Windows.Forms.Button();
            this.Auto_Upload_ckbox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // Get_day_Data
            // 
            this.Get_day_Data.Location = new System.Drawing.Point(12, 12);
            this.Get_day_Data.Name = "Get_day_Data";
            this.Get_day_Data.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.Get_day_Data.Size = new System.Drawing.Size(119, 82);
            this.Get_day_Data.TabIndex = 0;
            this.Get_day_Data.Text = "      取得今日活動       (秀在右邊)";
            this.Get_day_Data.UseVisualStyleBackColor = true;
            this.Get_day_Data.Click += new System.EventHandler(this.Get_day_data_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(383, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(647, 804);
            this.textBox1.TabIndex = 1;
            // 
            // EXIT
            // 
            this.EXIT.Location = new System.Drawing.Point(12, 734);
            this.EXIT.Name = "EXIT";
            this.EXIT.Size = new System.Drawing.Size(118, 82);
            this.EXIT.TabIndex = 3;
            this.EXIT.Text = "EXIT";
            this.EXIT.UseVisualStyleBackColor = true;
            this.EXIT.Click += new System.EventHandler(this.EXIT_Click);
            // 
            // Get_Calendar
            // 
            this.Get_Calendar.Location = new System.Drawing.Point(13, 100);
            this.Get_Calendar.Name = "Get_Calendar";
            this.Get_Calendar.Size = new System.Drawing.Size(118, 82);
            this.Get_Calendar.TabIndex = 4;
            this.Get_Calendar.Text = "      取得整月活動       (存成Excel)";
            this.Get_Calendar.UseVisualStyleBackColor = true;
            this.Get_Calendar.Click += new System.EventHandler(this.Get_Calendar_Click);
            // 
            // dateTimePicker_day
            // 
            this.dateTimePicker_day.Location = new System.Drawing.Point(177, 40);
            this.dateTimePicker_day.Name = "dateTimePicker_day";
            this.dateTimePicker_day.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker_day.TabIndex = 5;
            // 
            // dateTimePicker_month_1
            // 
            this.dateTimePicker_month_1.Location = new System.Drawing.Point(177, 114);
            this.dateTimePicker_month_1.Name = "dateTimePicker_month_1";
            this.dateTimePicker_month_1.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker_month_1.TabIndex = 6;
            // 
            // dateTimePicker_month_2
            // 
            this.dateTimePicker_month_2.Location = new System.Drawing.Point(177, 142);
            this.dateTimePicker_month_2.Name = "dateTimePicker_month_2";
            this.dateTimePicker_month_2.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker_month_2.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(210, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "選年月就好，日隨便";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(146, 119);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "開始";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(146, 147);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "結束";
            // 
            // Upload_Calendar
            // 
            this.Upload_Calendar.Location = new System.Drawing.Point(12, 188);
            this.Upload_Calendar.Name = "Upload_Calendar";
            this.Upload_Calendar.Size = new System.Drawing.Size(118, 82);
            this.Upload_Calendar.TabIndex = 10;
            this.Upload_Calendar.Text = "上傳行事曆";
            this.Upload_Calendar.UseVisualStyleBackColor = true;
            this.Upload_Calendar.Click += new System.EventHandler(this.Upload_Calendar_Click);
            // 
            // Delete_Table
            // 
            this.Delete_Table.Location = new System.Drawing.Point(12, 364);
            this.Delete_Table.Name = "Delete_Table";
            this.Delete_Table.Size = new System.Drawing.Size(118, 82);
            this.Delete_Table.TabIndex = 11;
            this.Delete_Table.Text = "刪除表格";
            this.Delete_Table.UseVisualStyleBackColor = true;
            this.Delete_Table.Click += new System.EventHandler(this.Delete_Table_Click);
            // 
            // Create_Table
            // 
            this.Create_Table.Location = new System.Drawing.Point(13, 276);
            this.Create_Table.Name = "Create_Table";
            this.Create_Table.Size = new System.Drawing.Size(118, 82);
            this.Create_Table.TabIndex = 12;
            this.Create_Table.Text = "新增表格";
            this.Create_Table.UseVisualStyleBackColor = true;
            this.Create_Table.Click += new System.EventHandler(this.Create_Table_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(141, 310);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 12);
            this.label4.TabIndex = 14;
            this.label4.Text = "create";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(140, 399);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "delete";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(140, 223);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "upload";
            // 
            // Clear_Table
            // 
            this.Clear_Table.Location = new System.Drawing.Point(13, 452);
            this.Clear_Table.Name = "Clear_Table";
            this.Clear_Table.Size = new System.Drawing.Size(118, 82);
            this.Clear_Table.TabIndex = 15;
            this.Clear_Table.Text = "清除表格";
            this.Clear_Table.UseVisualStyleBackColor = true;
            this.Clear_Table.Click += new System.EventHandler(this.Clear_Table_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(141, 487);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(27, 12);
            this.label7.TabIndex = 14;
            this.label7.Text = "clear";
            // 
            // Auto_upload_Timer
            // 
            this.Auto_upload_Timer.Enabled = true;
            this.Auto_upload_Timer.Interval = 3600000;
            this.Auto_upload_Timer.Tick += new System.EventHandler(this.Auto_upload_Timer_Tick);
            // 
            // Get_Announcement
            // 
            this.Get_Announcement.Location = new System.Drawing.Point(13, 540);
            this.Get_Announcement.Name = "Get_Announcement";
            this.Get_Announcement.Size = new System.Drawing.Size(118, 82);
            this.Get_Announcement.TabIndex = 15;
            this.Get_Announcement.Text = "取得公告";
            this.Get_Announcement.UseVisualStyleBackColor = true;
            this.Get_Announcement.Click += new System.EventHandler(this.Get_Announcement_Click);
            // 
            // Upload_Text
            // 
            this.Upload_Text.Location = new System.Drawing.Point(184, 219);
            this.Upload_Text.Name = "Upload_Text";
            this.Upload_Text.Size = new System.Drawing.Size(100, 22);
            this.Upload_Text.TabIndex = 20;
            // 
            // Create_Text
            // 
            this.Create_Text.Location = new System.Drawing.Point(184, 307);
            this.Create_Text.Name = "Create_Text";
            this.Create_Text.Size = new System.Drawing.Size(100, 22);
            this.Create_Text.TabIndex = 21;
            // 
            // Delete_Text
            // 
            this.Delete_Text.Location = new System.Drawing.Point(184, 396);
            this.Delete_Text.Name = "Delete_Text";
            this.Delete_Text.Size = new System.Drawing.Size(100, 22);
            this.Delete_Text.TabIndex = 22;
            // 
            // Clear_Text
            // 
            this.Clear_Text.Location = new System.Drawing.Point(184, 483);
            this.Clear_Text.Name = "Clear_Text";
            this.Clear_Text.Size = new System.Drawing.Size(100, 22);
            this.Clear_Text.TabIndex = 23;
            // 
            // Upload_Announcement
            // 
            this.Upload_Announcement.Location = new System.Drawing.Point(12, 628);
            this.Upload_Announcement.Name = "Upload_Announcement";
            this.Upload_Announcement.Size = new System.Drawing.Size(118, 82);
            this.Upload_Announcement.TabIndex = 24;
            this.Upload_Announcement.Text = "上傳公告";
            this.Upload_Announcement.UseVisualStyleBackColor = true;
            this.Upload_Announcement.Click += new System.EventHandler(this.Upload_Announcement_Click);
            // 
            // bt_TEST
            // 
            this.bt_TEST.Location = new System.Drawing.Point(166, 628);
            this.bt_TEST.Name = "bt_TEST";
            this.bt_TEST.Size = new System.Drawing.Size(118, 82);
            this.bt_TEST.TabIndex = 25;
            this.bt_TEST.Text = "TEST";
            this.bt_TEST.UseVisualStyleBackColor = true;
            this.bt_TEST.Click += new System.EventHandler(this.bt_TEST_Click);
            // 
            // Auto_Upload_ckbox
            // 
            this.Auto_Upload_ckbox.AutoSize = true;
            this.Auto_Upload_ckbox.Location = new System.Drawing.Point(148, 768);
            this.Auto_Upload_ckbox.Name = "Auto_Upload_ckbox";
            this.Auto_Upload_ckbox.Size = new System.Drawing.Size(72, 16);
            this.Auto_Upload_ckbox.TabIndex = 26;
            this.Auto_Upload_ckbox.Text = "自動上傳";
            this.Auto_Upload_ckbox.UseVisualStyleBackColor = true;
            this.Auto_Upload_ckbox.CheckedChanged += new System.EventHandler(this.Auto_Upload_ckbox_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1042, 825);
            this.Controls.Add(this.Auto_Upload_ckbox);
            this.Controls.Add(this.bt_TEST);
            this.Controls.Add(this.Upload_Announcement);
            this.Controls.Add(this.Clear_Text);
            this.Controls.Add(this.Delete_Text);
            this.Controls.Add(this.Create_Text);
            this.Controls.Add(this.Upload_Text);
            this.Controls.Add(this.Get_Announcement);
            this.Controls.Add(this.Clear_Table);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.Create_Table);
            this.Controls.Add(this.Delete_Table);
            this.Controls.Add(this.Upload_Calendar);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker_month_2);
            this.Controls.Add(this.dateTimePicker_month_1);
            this.Controls.Add(this.dateTimePicker_day);
            this.Controls.Add(this.Get_Calendar);
            this.Controls.Add(this.EXIT);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.Get_day_Data);
            this.Name = "Form1";
            this.Text = "Web Crawler";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Get_day_Data;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button EXIT;
        private System.Windows.Forms.Button Get_Calendar;
        private System.Windows.Forms.DateTimePicker dateTimePicker_day;
        private System.Windows.Forms.DateTimePicker dateTimePicker_month_1;
        private System.Windows.Forms.DateTimePicker dateTimePicker_month_2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button Upload_Calendar;
        private System.Windows.Forms.Button Delete_Table;
        private System.Windows.Forms.Button Create_Table;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button Clear_Table;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Timer Auto_upload_Timer;
        private System.Windows.Forms.Button Get_Announcement;
        private System.Windows.Forms.TextBox Upload_Text;
        private System.Windows.Forms.TextBox Create_Text;
        private System.Windows.Forms.TextBox Delete_Text;
        private System.Windows.Forms.TextBox Clear_Text;
        private System.Windows.Forms.Button Upload_Announcement;
        private System.Windows.Forms.Button bt_TEST;
        private System.Windows.Forms.CheckBox Auto_Upload_ckbox;
    }
}


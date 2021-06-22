namespace ExcelAddProject
{
    partial class COKReport
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
            this.dtpStart = new System.Windows.Forms.DateTimePicker();
            this.dtpEnd = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cbWBCount = new System.Windows.Forms.CheckBox();
            this.bCountReject = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rbRange = new System.Windows.Forms.RadioButton();
            this.rbMonth = new System.Windows.Forms.RadioButton();
            this.dtpMonth = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtpStart
            // 
            this.dtpStart.Location = new System.Drawing.Point(16, 125);
            this.dtpStart.Name = "dtpStart";
            this.dtpStart.Size = new System.Drawing.Size(137, 20);
            this.dtpStart.TabIndex = 0;
            // 
            // dtpEnd
            // 
            this.dtpEnd.Location = new System.Drawing.Point(159, 125);
            this.dtpEnd.Name = "dtpEnd";
            this.dtpEnd.Size = new System.Drawing.Size(138, 20);
            this.dtpEnd.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(175, 105);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 17);
            this.label2.TabIndex = 16;
            this.label2.Text = "Конец периода";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(20, 105);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 17);
            this.label1.TabIndex = 15;
            this.label1.Text = "Начало периода";
            // 
            // cbWBCount
            // 
            this.cbWBCount.AutoSize = true;
            this.cbWBCount.Location = new System.Drawing.Point(12, 176);
            this.cbWBCount.Name = "cbWBCount";
            this.cbWBCount.Size = new System.Drawing.Size(134, 17);
            this.cbWBCount.TabIndex = 14;
            this.cbWBCount.Text = "Пересчет WB сборки";
            this.cbWBCount.UseVisualStyleBackColor = true;
            // 
            // bCountReject
            // 
            this.bCountReject.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.bCountReject.Location = new System.Drawing.Point(12, 199);
            this.bCountReject.Name = "bCountReject";
            this.bCountReject.Size = new System.Drawing.Size(317, 56);
            this.bCountReject.TabIndex = 17;
            this.bCountReject.Text = "Посчитать отчет в ЦОК";
            this.bCountReject.UseVisualStyleBackColor = true;
            this.bCountReject.Click += new System.EventHandler(this.bCountReject_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dtpMonth);
            this.groupBox1.Controls.Add(this.rbMonth);
            this.groupBox1.Controls.Add(this.rbRange);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dtpStart);
            this.groupBox1.Controls.Add(this.dtpEnd);
            this.groupBox1.Location = new System.Drawing.Point(15, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(314, 162);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            // 
            // rbRange
            // 
            this.rbRange.AutoSize = true;
            this.rbRange.Location = new System.Drawing.Point(24, 85);
            this.rbRange.Name = "rbRange";
            this.rbRange.Size = new System.Drawing.Size(96, 17);
            this.rbRange.TabIndex = 0;
            this.rbRange.TabStop = true;
            this.rbRange.Text = "Диапазон дат";
            this.rbRange.UseVisualStyleBackColor = true;
            this.rbRange.CheckedChanged += new System.EventHandler(this.rbRange_CheckedChanged);
            // 
            // rbMonth
            // 
            this.rbMonth.AutoSize = true;
            this.rbMonth.Location = new System.Drawing.Point(24, 19);
            this.rbMonth.Name = "rbMonth";
            this.rbMonth.Size = new System.Drawing.Size(58, 17);
            this.rbMonth.TabIndex = 17;
            this.rbMonth.TabStop = true;
            this.rbMonth.Text = "Месяц";
            this.rbMonth.UseVisualStyleBackColor = true;
            this.rbMonth.CheckedChanged += new System.EventHandler(this.rbMonth_CheckedChanged);
            // 
            // dtpMonth
            // 
            this.dtpMonth.CustomFormat = "MMMMyyyy";
            this.dtpMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpMonth.Location = new System.Drawing.Point(24, 59);
            this.dtpMonth.Name = "dtpMonth";
            this.dtpMonth.Size = new System.Drawing.Size(116, 20);
            this.dtpMonth.TabIndex = 18;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(21, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(119, 17);
            this.label3.TabIndex = 19;
            this.label3.Text = "Отчетный месяц";
            // 
            // COKReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(341, 265);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.bCountReject);
            this.Controls.Add(this.cbWBCount);
            this.Name = "COKReport";
            this.Text = "COKReport";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpStart;
        private System.Windows.Forms.DateTimePicker dtpEnd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cbWBCount;
        private System.Windows.Forms.Button bCountReject;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpMonth;
        private System.Windows.Forms.RadioButton rbMonth;
        private System.Windows.Forms.RadioButton rbRange;
    }
}
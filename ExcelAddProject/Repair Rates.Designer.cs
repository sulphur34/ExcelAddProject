namespace ExcelAddProject
{
    partial class RepairForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RepairForm));
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.rbWorkshop = new System.Windows.Forms.RadioButton();
            this.rbAll = new System.Windows.Forms.RadioButton();
            this.rbErection = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cbOfficial = new System.Windows.Forms.CheckBox();
            this.cbDiameter = new System.Windows.Forms.CheckBox();
            this.cbRepairValid = new System.Windows.Forms.CheckBox();
            this.cbVolume = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(158, 29);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(130, 20);
            this.dateTimePicker1.TabIndex = 0;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(15, 29);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(131, 20);
            this.dateTimePicker2.TabIndex = 1;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(152, 55);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(136, 81);
            this.button1.TabIndex = 4;
            this.button1.Text = "Сформировать процент брака сварщиков";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(12, 144);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(134, 17);
            this.checkBox1.TabIndex = 6;
            this.checkBox1.Text = "Пересчет WB сборки";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // rbWorkshop
            // 
            this.rbWorkshop.AutoSize = true;
            this.rbWorkshop.Location = new System.Drawing.Point(24, 58);
            this.rbWorkshop.Name = "rbWorkshop";
            this.rbWorkshop.Size = new System.Drawing.Size(44, 17);
            this.rbWorkshop.TabIndex = 7;
            this.rbWorkshop.TabStop = true;
            this.rbWorkshop.Text = "Цех";
            this.rbWorkshop.UseVisualStyleBackColor = true;
            this.rbWorkshop.CheckedChanged += new System.EventHandler(this.rbWorkshop_CheckedChanged);
            // 
            // rbAll
            // 
            this.rbAll.AutoSize = true;
            this.rbAll.Location = new System.Drawing.Point(24, 10);
            this.rbAll.Name = "rbAll";
            this.rbAll.Size = new System.Drawing.Size(88, 17);
            this.rbAll.TabIndex = 8;
            this.rbAll.TabStop = true;
            this.rbAll.Text = "Весь проект";
            this.rbAll.UseVisualStyleBackColor = true;
            this.rbAll.CheckedChanged += new System.EventHandler(this.rbAll_CheckedChanged);
            // 
            // rbErection
            // 
            this.rbErection.AutoSize = true;
            this.rbErection.Location = new System.Drawing.Point(24, 33);
            this.rbErection.Name = "rbErection";
            this.rbErection.Size = new System.Drawing.Size(65, 17);
            this.rbErection.TabIndex = 9;
            this.rbErection.TabStop = true;
            this.rbErection.Text = "Монтаж";
            this.rbErection.UseVisualStyleBackColor = true;
            this.rbErection.CheckedChanged += new System.EventHandler(this.rbErection_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbAll);
            this.groupBox1.Controls.Add(this.rbErection);
            this.groupBox1.Controls.Add(this.rbWorkshop);
            this.groupBox1.Location = new System.Drawing.Point(15, 55);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(130, 81);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(17, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 17);
            this.label1.TabIndex = 11;
            this.label1.Text = "Начало периода";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(170, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 17);
            this.label2.TabIndex = 12;
            this.label2.Text = "Конец периода";
            // 
            // cbOfficial
            // 
            this.cbOfficial.AutoSize = true;
            this.cbOfficial.Location = new System.Drawing.Point(12, 167);
            this.cbOfficial.Name = "cbOfficial";
            this.cbOfficial.Size = new System.Drawing.Size(58, 17);
            this.cbOfficial.TabIndex = 13;
            this.cbOfficial.Text = "Official";
            this.cbOfficial.UseVisualStyleBackColor = true;
            // 
            // cbDiameter
            // 
            this.cbDiameter.AutoSize = true;
            this.cbDiameter.Location = new System.Drawing.Point(152, 142);
            this.cbDiameter.Name = "cbDiameter";
            this.cbDiameter.Size = new System.Drawing.Size(141, 17);
            this.cbDiameter.TabIndex = 14;
            this.cbDiameter.Text = "Не учитывать диаметр";
            this.cbDiameter.UseVisualStyleBackColor = true;
            // 
            // cbRepairValid
            // 
            this.cbRepairValid.AutoSize = true;
            this.cbRepairValid.Location = new System.Drawing.Point(152, 167);
            this.cbRepairValid.Name = "cbRepairValid";
            this.cbRepairValid.Size = new System.Drawing.Size(143, 17);
            this.cbRepairValid.TabIndex = 15;
            this.cbRepairValid.Text = "Не учитывать ремонты";
            this.cbRepairValid.UseVisualStyleBackColor = true;
            // 
            // cbVolume
            // 
            this.cbVolume.AutoSize = true;
            this.cbVolume.Location = new System.Drawing.Point(76, 167);
            this.cbVolume.Name = "cbVolume";
            this.cbVolume.Size = new System.Drawing.Size(61, 17);
            this.cbVolume.TabIndex = 16;
            this.cbVolume.Text = "Volume";
            this.cbVolume.UseVisualStyleBackColor = true;
            this.cbVolume.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // RepairForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 196);
            this.Controls.Add(this.cbVolume);
            this.Controls.Add(this.cbRepairValid);
            this.Controls.Add(this.cbDiameter);
            this.Controls.Add(this.cbOfficial);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "RepairForm";
            this.Text = "Ultimate repair rate counter 9000";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.RadioButton rbWorkshop;
        private System.Windows.Forms.RadioButton rbAll;
        private System.Windows.Forms.RadioButton rbErection;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cbOfficial;
        private System.Windows.Forms.CheckBox cbDiameter;
        private System.Windows.Forms.CheckBox cbRepairValid;
        private System.Windows.Forms.CheckBox cbVolume;
    }
}
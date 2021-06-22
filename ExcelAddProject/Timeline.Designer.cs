namespace ExcelAddProject
{
    partial class Timeline
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
            this.bCountReject = new System.Windows.Forms.Button();
            this.cbWBCount = new System.Windows.Forms.CheckBox();
            this.dtpEnd = new System.Windows.Forms.DateTimePicker();
            this.dtpStart = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.cbOfficial = new System.Windows.Forms.CheckBox();
            this.gbMaterial = new System.Windows.Forms.GroupBox();
            this.rbLTCS = new System.Windows.Forms.RadioButton();
            this.rbSS = new System.Windows.Forms.RadioButton();
            this.rbALLOY = new System.Windows.Forms.RadioButton();
            this.rbF22 = new System.Windows.Forms.RadioButton();
            this.gbMaterial.SuspendLayout();
            this.SuspendLayout();
            // 
            // bCountReject
            // 
            this.bCountReject.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.bCountReject.Location = new System.Drawing.Point(12, 134);
            this.bCountReject.Name = "bCountReject";
            this.bCountReject.Size = new System.Drawing.Size(281, 56);
            this.bCountReject.TabIndex = 20;
            this.bCountReject.Text = "Посчитать";
            this.bCountReject.UseVisualStyleBackColor = true;
            this.bCountReject.Click += new System.EventHandler(this.bCountReject_Click);
            // 
            // cbWBCount
            // 
            this.cbWBCount.AutoSize = true;
            this.cbWBCount.Location = new System.Drawing.Point(12, 55);
            this.cbWBCount.Name = "cbWBCount";
            this.cbWBCount.Size = new System.Drawing.Size(134, 17);
            this.cbWBCount.TabIndex = 19;
            this.cbWBCount.Text = "Пересчет WB сборки";
            this.cbWBCount.UseVisualStyleBackColor = true;
            // 
            // dtpEnd
            // 
            this.dtpEnd.Location = new System.Drawing.Point(155, 29);
            this.dtpEnd.Name = "dtpEnd";
            this.dtpEnd.Size = new System.Drawing.Size(138, 20);
            this.dtpEnd.TabIndex = 1;
            // 
            // dtpStart
            // 
            this.dtpStart.Location = new System.Drawing.Point(12, 29);
            this.dtpStart.Name = "dtpStart";
            this.dtpStart.Size = new System.Drawing.Size(137, 20);
            this.dtpStart.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(171, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 17);
            this.label2.TabIndex = 16;
            this.label2.Text = "Конец периода";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(16, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(118, 17);
            this.label1.TabIndex = 15;
            this.label1.Text = "Начало периода";
            // 
            // cbOfficial
            // 
            this.cbOfficial.AutoSize = true;
            this.cbOfficial.Location = new System.Drawing.Point(12, 78);
            this.cbOfficial.Name = "cbOfficial";
            this.cbOfficial.Size = new System.Drawing.Size(58, 17);
            this.cbOfficial.TabIndex = 21;
            this.cbOfficial.Text = "Official";
            this.cbOfficial.UseVisualStyleBackColor = true;
            // 
            // gbMaterial
            // 
            this.gbMaterial.Controls.Add(this.rbF22);
            this.gbMaterial.Controls.Add(this.rbALLOY);
            this.gbMaterial.Controls.Add(this.rbSS);
            this.gbMaterial.Controls.Add(this.rbLTCS);
            this.gbMaterial.Location = new System.Drawing.Point(155, 55);
            this.gbMaterial.Name = "gbMaterial";
            this.gbMaterial.Size = new System.Drawing.Size(125, 73);
            this.gbMaterial.TabIndex = 22;
            this.gbMaterial.TabStop = false;
            // 
            // rbLTCS
            // 
            this.rbLTCS.AutoSize = true;
            this.rbLTCS.Location = new System.Drawing.Point(6, 19);
            this.rbLTCS.Name = "rbLTCS";
            this.rbLTCS.Size = new System.Drawing.Size(52, 17);
            this.rbLTCS.TabIndex = 0;
            this.rbLTCS.TabStop = true;
            this.rbLTCS.Text = "LTCS";
            this.rbLTCS.UseVisualStyleBackColor = true;
            this.rbLTCS.CheckedChanged += new System.EventHandler(this.rbLTCS_CheckedChanged);
            // 
            // rbSS
            // 
            this.rbSS.AutoSize = true;
            this.rbSS.Location = new System.Drawing.Point(64, 19);
            this.rbSS.Name = "rbSS";
            this.rbSS.Size = new System.Drawing.Size(39, 17);
            this.rbSS.TabIndex = 1;
            this.rbSS.TabStop = true;
            this.rbSS.Text = "SS";
            this.rbSS.UseVisualStyleBackColor = true;
            this.rbSS.CheckedChanged += new System.EventHandler(this.rbSS_CheckedChanged);
            // 
            // rbALLOY
            // 
            this.rbALLOY.AutoSize = true;
            this.rbALLOY.Location = new System.Drawing.Point(6, 42);
            this.rbALLOY.Name = "rbALLOY";
            this.rbALLOY.Size = new System.Drawing.Size(59, 17);
            this.rbALLOY.TabIndex = 2;
            this.rbALLOY.TabStop = true;
            this.rbALLOY.Text = "ALLOY";
            this.rbALLOY.UseVisualStyleBackColor = true;
            this.rbALLOY.CheckedChanged += new System.EventHandler(this.rbALLOY_CheckedChanged);
            // 
            // rbF22
            // 
            this.rbF22.AutoSize = true;
            this.rbF22.Location = new System.Drawing.Point(64, 42);
            this.rbF22.Name = "rbF22";
            this.rbF22.Size = new System.Drawing.Size(43, 17);
            this.rbF22.TabIndex = 3;
            this.rbF22.TabStop = true;
            this.rbF22.Text = "F22";
            this.rbF22.UseVisualStyleBackColor = true;
            this.rbF22.CheckedChanged += new System.EventHandler(this.rbF22_CheckedChanged);
            // 
            // Timeline
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(306, 204);
            this.Controls.Add(this.gbMaterial);
            this.Controls.Add(this.cbOfficial);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpStart);
            this.Controls.Add(this.bCountReject);
            this.Controls.Add(this.dtpEnd);
            this.Controls.Add(this.cbWBCount);
            this.Name = "Timeline";
            this.Text = "Timeline";
            this.gbMaterial.ResumeLayout(false);
            this.gbMaterial.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bCountReject;
        private System.Windows.Forms.CheckBox cbWBCount;
        private System.Windows.Forms.DateTimePicker dtpEnd;
        private System.Windows.Forms.DateTimePicker dtpStart;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox cbOfficial;
        private System.Windows.Forms.GroupBox gbMaterial;
        private System.Windows.Forms.RadioButton rbALLOY;
        private System.Windows.Forms.RadioButton rbSS;
        private System.Windows.Forms.RadioButton rbLTCS;
        private System.Windows.Forms.RadioButton rbF22;
    }
}
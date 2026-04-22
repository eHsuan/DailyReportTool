namespace DailyReportTool
{
    partial class Form1
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuPareto = new System.Windows.Forms.ToolStripMenuItem();
            this.menuBalance = new System.Windows.Forms.ToolStripMenuItem();
            this.pnlPareto = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.txtMaintenancePath = new System.Windows.Forms.TextBox();
            this.btnSelectMaintenance = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtConnectionPath = new System.Windows.Forms.TextBox();
            this.btnSelectConnection = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.pnlBalance = new System.Windows.Forms.Panel();
            this.lblDailyHours = new System.Windows.Forms.Label();
            this.numDailyHours = new System.Windows.Forms.NumericUpDown();
            this.lblDataDays = new System.Windows.Forms.Label();
            this.numDataDays = new System.Windows.Forms.NumericUpDown();
            this.btnImportIE = new System.Windows.Forms.Button();
            this.btnGenerateBalance = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtIEPath = new System.Windows.Forms.TextBox();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.pnlPareto.SuspendLayout();
            this.pnlBalance.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDailyHours)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDataDays)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuPareto,
            this.menuBalance});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(600, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuPareto
            // 
            this.menuPareto.Name = "menuPareto";
            this.menuPareto.Size = new System.Drawing.Size(82, 20);
            this.menuPareto.Text = "機故柏拉圖";
            this.menuPareto.Click += new System.EventHandler(this.menuPareto_Click);
            // 
            // menuBalance
            // 
            this.menuBalance.Name = "menuBalance";
            this.menuBalance.Size = new System.Drawing.Size(82, 20);
            this.menuBalance.Text = "產能平衡圖";
            this.menuBalance.Click += new System.EventHandler(this.menuBalance_Click);
            // 
            // pnlPareto
            // 
            this.pnlPareto.Controls.Add(this.label1);
            this.pnlPareto.Controls.Add(this.txtMaintenancePath);
            this.pnlPareto.Controls.Add(this.btnSelectMaintenance);
            this.pnlPareto.Controls.Add(this.label2);
            this.pnlPareto.Controls.Add(this.txtConnectionPath);
            this.pnlPareto.Controls.Add(this.btnSelectConnection);
            this.pnlPareto.Controls.Add(this.btnGenerate);
            this.pnlPareto.Controls.Add(this.btnClear);
            this.pnlPareto.Location = new System.Drawing.Point(0, 27);
            this.pnlPareto.Name = "pnlPareto";
            this.pnlPareto.Size = new System.Drawing.Size(600, 115);
            this.pnlPareto.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "維修保養紀錄";
            // 
            // txtMaintenancePath
            // 
            this.txtMaintenancePath.Location = new System.Drawing.Point(95, 12);
            this.txtMaintenancePath.Name = "txtMaintenancePath";
            this.txtMaintenancePath.ReadOnly = true;
            this.txtMaintenancePath.Size = new System.Drawing.Size(400, 22);
            this.txtMaintenancePath.TabIndex = 1;
            // 
            // btnSelectMaintenance
            // 
            this.btnSelectMaintenance.Location = new System.Drawing.Point(501, 10);
            this.btnSelectMaintenance.Name = "btnSelectMaintenance";
            this.btnSelectMaintenance.Size = new System.Drawing.Size(75, 23);
            this.btnSelectMaintenance.TabIndex = 2;
            this.btnSelectMaintenance.Text = "選擇檔案";
            this.btnSelectMaintenance.UseVisualStyleBackColor = true;
            this.btnSelectMaintenance.Click += new System.EventHandler(this.btnSelectMaintenance_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "設備連線資料";
            // 
            // txtConnectionPath
            // 
            this.txtConnectionPath.Location = new System.Drawing.Point(95, 47);
            this.txtConnectionPath.Name = "txtConnectionPath";
            this.txtConnectionPath.ReadOnly = true;
            this.txtConnectionPath.Size = new System.Drawing.Size(400, 22);
            this.txtConnectionPath.TabIndex = 4;
            // 
            // btnSelectConnection
            // 
            this.btnSelectConnection.Location = new System.Drawing.Point(501, 45);
            this.btnSelectConnection.Name = "btnSelectConnection";
            this.btnSelectConnection.Size = new System.Drawing.Size(75, 23);
            this.btnSelectConnection.TabIndex = 5;
            this.btnSelectConnection.Text = "選擇檔案";
            this.btnSelectConnection.UseVisualStyleBackColor = true;
            this.btnSelectConnection.Click += new System.EventHandler(this.btnSelectConnection_Click);
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(95, 85);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(100, 30);
            this.btnGenerate.TabIndex = 6;
            this.btnGenerate.Text = "產生日報表";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(210, 85);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(100, 30);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "清除匯入資料";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // pnlBalance
            // 
            this.pnlBalance.Controls.Add(this.lblDailyHours);
            this.pnlBalance.Controls.Add(this.numDailyHours);
            this.pnlBalance.Controls.Add(this.lblDataDays);
            this.pnlBalance.Controls.Add(this.numDataDays);
            this.pnlBalance.Controls.Add(this.btnImportIE);
            this.pnlBalance.Controls.Add(this.btnGenerateBalance);
            this.pnlBalance.Controls.Add(this.label3);
            this.pnlBalance.Controls.Add(this.txtIEPath);
            this.pnlBalance.Location = new System.Drawing.Point(0, 27);
            this.pnlBalance.Name = "pnlBalance";
            this.pnlBalance.Size = new System.Drawing.Size(600, 115);
            this.pnlBalance.TabIndex = 2;
            this.pnlBalance.Visible = false;
            // 
            // lblDailyHours
            // 
            this.lblDailyHours.AutoSize = true;
            this.lblDailyHours.Location = new System.Drawing.Point(180, 50);
            this.lblDailyHours.Name = "lblDailyHours";
            this.lblDailyHours.Size = new System.Drawing.Size(77, 12);
            this.lblDailyHours.TabIndex = 7;
            this.lblDailyHours.Text = "每日生產時間";
            // 
            // numDailyHours
            // 
            this.numDailyHours.Location = new System.Drawing.Point(265, 47);
            this.numDailyHours.Maximum = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.numDailyHours.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numDailyHours.Name = "numDailyHours";
            this.numDailyHours.Size = new System.Drawing.Size(60, 22);
            this.numDailyHours.TabIndex = 6;
            this.numDailyHours.Value = new decimal(new int[] {
            20,
            0,
            0,
            0});
            // 
            // lblDataDays
            // 
            this.lblDataDays.AutoSize = true;
            this.lblDataDays.Location = new System.Drawing.Point(12, 50);
            this.lblDataDays.Name = "lblDataDays";
            this.lblDataDays.Size = new System.Drawing.Size(53, 12);
            this.lblDailyHours.TabIndex = 5;
            this.lblDataDays.Text = "資料天數";
            // 
            // numDataDays
            // 
            this.numDataDays.Location = new System.Drawing.Point(95, 47);
            this.numDataDays.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numDataDays.Name = "numDataDays";
            this.numDataDays.Size = new System.Drawing.Size(60, 22);
            this.numDataDays.TabIndex = 4;
            this.numDataDays.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // btnImportIE
            // 
            this.btnImportIE.Location = new System.Drawing.Point(501, 10);
            this.btnImportIE.Name = "btnImportIE";
            this.btnImportIE.Size = new System.Drawing.Size(75, 23);
            this.btnImportIE.TabIndex = 3;
            this.btnImportIE.Text = "選擇檔案";
            this.btnImportIE.UseVisualStyleBackColor = true;
            this.btnImportIE.Click += new System.EventHandler(this.btnImportIE_Click);
            // 
            // btnGenerateBalance
            // 
            this.btnGenerateBalance.Location = new System.Drawing.Point(95, 85);
            this.btnGenerateBalance.Name = "btnGenerateBalance";
            this.btnGenerateBalance.Size = new System.Drawing.Size(120, 30);
            this.btnGenerateBalance.TabIndex = 2;
            this.btnGenerateBalance.Text = "產生產能平衡圖";
            this.btnGenerateBalance.UseVisualStyleBackColor = true;
            this.btnGenerateBalance.Click += new System.EventHandler(this.btnGenerateBalance_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "IE 資料檔案";
            // 
            // txtIEPath
            // 
            this.txtIEPath.Location = new System.Drawing.Point(95, 12);
            this.txtIEPath.Name = "txtIEPath";
            this.txtIEPath.ReadOnly = true;
            this.txtIEPath.Size = new System.Drawing.Size(400, 22);
            this.txtIEPath.TabIndex = 0;
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(12, 147);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(576, 201);
            this.txtLog.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 360);
            this.Controls.Add(this.pnlPareto);
            this.Controls.Add(this.pnlBalance);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "DailyReportTool v1.0.1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.pnlPareto.ResumeLayout(false);
            this.pnlPareto.PerformLayout();
            this.pnlBalance.ResumeLayout(false);
            this.pnlBalance.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDailyHours)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numDataDays)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuPareto;
        private System.Windows.Forms.ToolStripMenuItem menuBalance;
        private System.Windows.Forms.Panel pnlPareto;
        private System.Windows.Forms.Panel pnlBalance;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMaintenancePath;
        private System.Windows.Forms.Button btnSelectMaintenance;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtConnectionPath;
        private System.Windows.Forms.Button btnSelectConnection;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtIEPath;
        private System.Windows.Forms.Button btnImportIE;
        private System.Windows.Forms.Button btnGenerateBalance;
        private System.Windows.Forms.Label lblDataDays;
        private System.Windows.Forms.NumericUpDown numDataDays;
        private System.Windows.Forms.Label lblDailyHours;
        private System.Windows.Forms.NumericUpDown numDailyHours;
        private System.Windows.Forms.TextBox txtLog;
    }
}

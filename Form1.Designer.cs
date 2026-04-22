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
            this.label1 = new System.Windows.Forms.Label();
            this.txtMaintenancePath = new System.Windows.Forms.TextBox();
            this.btnSelectMaintenance = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtConnectionPath = new System.Windows.Forms.TextBox();
            this.btnSelectConnection = new System.Windows.Forms.Button();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
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
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(15, 130);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(561, 200);
            this.txtLog.TabIndex = 8;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(600, 350);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.btnSelectConnection);
            this.Controls.Add(this.txtConnectionPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSelectMaintenance);
            this.Controls.Add(this.txtMaintenancePath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "DailyReportTool";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMaintenancePath;
        private System.Windows.Forms.Button btnSelectMaintenance;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtConnectionPath;
        private System.Windows.Forms.Button btnSelectConnection;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.TextBox txtLog;
    }
}

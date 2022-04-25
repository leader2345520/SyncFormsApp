
namespace SyncFormsApp
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
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
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
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnBrowseDell = new System.Windows.Forms.Button();
            this.BtnBrowseRMS = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.TxtDellPath = new System.Windows.Forms.TextBox();
            this.TxtRmsPath = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.BtnBrowseMultiDell = new System.Windows.Forms.Button();
            this.TxtDellMultiPath = new System.Windows.Forms.TextBox();
            this.BtnExe = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(454, 328);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.BtnBrowseDell);
            this.tabPage1.Controls.Add(this.BtnBrowseRMS);
            this.tabPage1.Controls.Add(this.textBox3);
            this.tabPage1.Controls.Add(this.TxtDellPath);
            this.tabPage1.Controls.Add(this.TxtRmsPath);
            this.tabPage1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(446, 297);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "RMS to Dell";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(13, 159);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(414, 54);
            this.label3.TabIndex = 7;
            this.label3.Text = "Enter project filter\r\n\r\n(\"None\" for all project. Use \",\" to separate multiple pro" +
    "jects)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(108, 15);
            this.label2.TabIndex = 6;
            this.label2.Text = "Choose Dell file";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 15);
            this.label1.TabIndex = 5;
            this.label1.Text = "Choose RMS file";
            // 
            // BtnBrowseDell
            // 
            this.BtnBrowseDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.BtnBrowseDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnBrowseDell.ForeColor = System.Drawing.Color.Black;
            this.BtnBrowseDell.Location = new System.Drawing.Point(336, 112);
            this.BtnBrowseDell.Name = "BtnBrowseDell";
            this.BtnBrowseDell.Size = new System.Drawing.Size(91, 23);
            this.BtnBrowseDell.TabIndex = 4;
            this.BtnBrowseDell.Text = "Browse";
            this.BtnBrowseDell.UseVisualStyleBackColor = false;
            this.BtnBrowseDell.Click += new System.EventHandler(this.BtnBrowseDell_Click);
            // 
            // BtnBrowseRMS
            // 
            this.BtnBrowseRMS.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.BtnBrowseRMS.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnBrowseRMS.ForeColor = System.Drawing.Color.Black;
            this.BtnBrowseRMS.Location = new System.Drawing.Point(336, 41);
            this.BtnBrowseRMS.Name = "BtnBrowseRMS";
            this.BtnBrowseRMS.Size = new System.Drawing.Size(91, 23);
            this.BtnBrowseRMS.TabIndex = 3;
            this.BtnBrowseRMS.Text = "Browse";
            this.BtnBrowseRMS.UseVisualStyleBackColor = false;
            this.BtnBrowseRMS.Click += new System.EventHandler(this.BtnBrowseRMS_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(12, 217);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(318, 23);
            this.textBox3.TabIndex = 2;
            this.textBox3.Text = "None";
            // 
            // TxtDellPath
            // 
            this.TxtDellPath.Location = new System.Drawing.Point(12, 113);
            this.TxtDellPath.Name = "TxtDellPath";
            this.TxtDellPath.Size = new System.Drawing.Size(318, 23);
            this.TxtDellPath.TabIndex = 1;
            // 
            // TxtRmsPath
            // 
            this.TxtRmsPath.BackColor = System.Drawing.SystemColors.Window;
            this.TxtRmsPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.TxtRmsPath.Location = new System.Drawing.Point(12, 42);
            this.TxtRmsPath.Name = "TxtRmsPath";
            this.TxtRmsPath.Size = new System.Drawing.Size(318, 23);
            this.TxtRmsPath.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.BtnBrowseMultiDell);
            this.tabPage2.Controls.Add(this.TxtDellMultiPath);
            this.tabPage2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage2.Location = new System.Drawing.Point(4, 27);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(446, 297);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Dell merge";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(108, 15);
            this.label4.TabIndex = 8;
            this.label4.Text = "Choose Dell file";
            // 
            // BtnBrowseMultiDell
            // 
            this.BtnBrowseMultiDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.BtnBrowseMultiDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnBrowseMultiDell.ForeColor = System.Drawing.Color.Black;
            this.BtnBrowseMultiDell.Location = new System.Drawing.Point(336, 41);
            this.BtnBrowseMultiDell.Name = "BtnBrowseMultiDell";
            this.BtnBrowseMultiDell.Size = new System.Drawing.Size(91, 23);
            this.BtnBrowseMultiDell.TabIndex = 7;
            this.BtnBrowseMultiDell.Text = "Browse";
            this.BtnBrowseMultiDell.UseVisualStyleBackColor = false;
            this.BtnBrowseMultiDell.Click += new System.EventHandler(this.BtnBrowseMultiDell_Click);
            // 
            // TxtDellMultiPath
            // 
            this.TxtDellMultiPath.BackColor = System.Drawing.SystemColors.Window;
            this.TxtDellMultiPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.TxtDellMultiPath.Location = new System.Drawing.Point(12, 42);
            this.TxtDellMultiPath.Name = "TxtDellMultiPath";
            this.TxtDellMultiPath.Size = new System.Drawing.Size(318, 23);
            this.TxtDellMultiPath.TabIndex = 6;
            // 
            // BtnExe
            // 
            this.BtnExe.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnExe.Location = new System.Drawing.Point(12, 287);
            this.BtnExe.Name = "BtnExe";
            this.BtnExe.Size = new System.Drawing.Size(91, 30);
            this.BtnExe.TabIndex = 5;
            this.BtnExe.Text = "Execute";
            this.BtnExe.UseVisualStyleBackColor = true;
            this.BtnExe.Click += new System.EventHandler(this.BtnExe_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCancel.Location = new System.Drawing.Point(110, 287);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(91, 30);
            this.BtnCancel.TabIndex = 6;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(454, 328);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnExe);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Sync_Tool";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnBrowseDell;
        private System.Windows.Forms.Button BtnBrowseRMS;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox TxtDellPath;
        private System.Windows.Forms.TextBox TxtRmsPath;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button BtnExe;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button BtnBrowseMultiDell;
        private System.Windows.Forms.TextBox TxtDellMultiPath;
    }
}


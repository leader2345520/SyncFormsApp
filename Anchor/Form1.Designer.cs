
namespace Anchor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnBrowseDell = new System.Windows.Forms.Button();
            this.btnBrowseRMS = new System.Windows.Forms.Button();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.txtDellPath = new System.Windows.Forms.TextBox();
            this.txtRmsPath = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBrowseMultiDell = new System.Windows.Forms.Button();
            this.txtDellMultiPath = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.btnBrowseSap = new System.Windows.Forms.Button();
            this.txtSapPath = new System.Windows.Forms.TextBox();
            this.btnExe = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.prgsBar = new System.Windows.Forms.ProgressBar();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage2);
            this.tabControl.Controls.Add(this.tabPage3);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(449, 327);
            this.tabControl.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnBrowseDell);
            this.tabPage1.Controls.Add(this.btnBrowseRMS);
            this.tabPage1.Controls.Add(this.txtFilter);
            this.tabPage1.Controls.Add(this.txtDellPath);
            this.tabPage1.Controls.Add(this.txtRmsPath);
            this.tabPage1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage1.Location = new System.Drawing.Point(4, 27);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(441, 296);
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
            this.label1.Size = new System.Drawing.Size(173, 15);
            this.label1.TabIndex = 5;
            this.label1.Text = "Choose RMS file(.csv,.xls)";
            // 
            // btnBrowseDell
            // 
            this.btnBrowseDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseDell.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseDell.Location = new System.Drawing.Point(336, 112);
            this.btnBrowseDell.Name = "btnBrowseDell";
            this.btnBrowseDell.Size = new System.Drawing.Size(91, 23);
            this.btnBrowseDell.TabIndex = 4;
            this.btnBrowseDell.Text = "Browse";
            this.btnBrowseDell.UseVisualStyleBackColor = false;
            this.btnBrowseDell.Click += new System.EventHandler(this.BtnBrowseDell_Click);
            // 
            // btnBrowseRMS
            // 
            this.btnBrowseRMS.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseRMS.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseRMS.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseRMS.Location = new System.Drawing.Point(336, 41);
            this.btnBrowseRMS.Name = "btnBrowseRMS";
            this.btnBrowseRMS.Size = new System.Drawing.Size(91, 23);
            this.btnBrowseRMS.TabIndex = 3;
            this.btnBrowseRMS.Text = "Browse";
            this.btnBrowseRMS.UseVisualStyleBackColor = false;
            this.btnBrowseRMS.Click += new System.EventHandler(this.BtnBrowseRMS_Click);
            // 
            // txtFilter
            // 
            this.txtFilter.Location = new System.Drawing.Point(12, 217);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(318, 23);
            this.txtFilter.TabIndex = 2;
            this.txtFilter.Text = "None";
            // 
            // txtDellPath
            // 
            this.txtDellPath.Location = new System.Drawing.Point(12, 113);
            this.txtDellPath.Name = "txtDellPath";
            this.txtDellPath.Size = new System.Drawing.Size(318, 23);
            this.txtDellPath.TabIndex = 1;
            // 
            // txtRmsPath
            // 
            this.txtRmsPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtRmsPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtRmsPath.Location = new System.Drawing.Point(12, 42);
            this.txtRmsPath.Name = "txtRmsPath";
            this.txtRmsPath.Size = new System.Drawing.Size(318, 23);
            this.txtRmsPath.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.btnBrowseMultiDell);
            this.tabPage2.Controls.Add(this.txtDellMultiPath);
            this.tabPage2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage2.Location = new System.Drawing.Point(4, 27);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(441, 296);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Dell merge";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 24);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(147, 15);
            this.label4.TabIndex = 8;
            this.label4.Text = "Choose Dell file(.xlsx)";
            // 
            // btnBrowseMultiDell
            // 
            this.btnBrowseMultiDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseMultiDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseMultiDell.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseMultiDell.Location = new System.Drawing.Point(336, 41);
            this.btnBrowseMultiDell.Name = "btnBrowseMultiDell";
            this.btnBrowseMultiDell.Size = new System.Drawing.Size(91, 23);
            this.btnBrowseMultiDell.TabIndex = 7;
            this.btnBrowseMultiDell.Text = "Browse";
            this.btnBrowseMultiDell.UseVisualStyleBackColor = false;
            this.btnBrowseMultiDell.Click += new System.EventHandler(this.BtnBrowseMultiDell_Click);
            // 
            // txtDellMultiPath
            // 
            this.txtDellMultiPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtDellMultiPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtDellMultiPath.Location = new System.Drawing.Point(12, 42);
            this.txtDellMultiPath.Name = "txtDellMultiPath";
            this.txtDellMultiPath.Size = new System.Drawing.Size(318, 23);
            this.txtDellMultiPath.TabIndex = 6;
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.btnBrowseSap);
            this.tabPage3.Controls.Add(this.txtSapPath);
            this.tabPage3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage3.Location = new System.Drawing.Point(4, 27);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(441, 296);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "SAP format";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(161, 15);
            this.label6.TabIndex = 11;
            this.label6.Text = "Choose SAP table(.xlsx)";
            // 
            // btnBrowseSap
            // 
            this.btnBrowseSap.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseSap.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseSap.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseSap.Location = new System.Drawing.Point(336, 41);
            this.btnBrowseSap.Name = "btnBrowseSap";
            this.btnBrowseSap.Size = new System.Drawing.Size(91, 23);
            this.btnBrowseSap.TabIndex = 9;
            this.btnBrowseSap.Text = "Browse";
            this.btnBrowseSap.UseVisualStyleBackColor = false;
            this.btnBrowseSap.Click += new System.EventHandler(this.btnBrowseSap_Click);
            // 
            // txtSapPath
            // 
            this.txtSapPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtSapPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtSapPath.Location = new System.Drawing.Point(12, 42);
            this.txtSapPath.Name = "txtSapPath";
            this.txtSapPath.Size = new System.Drawing.Size(318, 23);
            this.txtSapPath.TabIndex = 7;
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExe.Location = new System.Drawing.Point(18, 286);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(91, 30);
            this.btnExe.TabIndex = 5;
            this.btnExe.Text = "Execute";
            this.btnExe.UseVisualStyleBackColor = true;
            this.btnExe.Click += new System.EventHandler(this.BtnExe_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(115, 286);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(91, 30);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // prgsBar
            // 
            this.prgsBar.Location = new System.Drawing.Point(213, 287);
            this.prgsBar.Name = "prgsBar";
            this.prgsBar.Size = new System.Drawing.Size(220, 26);
            this.prgsBar.TabIndex = 7;
            this.prgsBar.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(449, 327);
            this.Controls.Add(this.prgsBar);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.tabControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Anchor";
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBrowseDell;
        private System.Windows.Forms.Button btnBrowseRMS;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.TextBox txtDellPath;
        private System.Windows.Forms.TextBox txtRmsPath;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnExe;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBrowseMultiDell;
        private System.Windows.Forms.TextBox txtDellMultiPath;
        private System.Windows.Forms.ProgressBar prgsBar;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnBrowseSap;
        private System.Windows.Forms.TextBox txtSapPath;
    }
}


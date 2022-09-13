
using System.Drawing;

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
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.label12 = new System.Windows.Forms.Label();
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
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.label10 = new System.Windows.Forms.Label();
            this.btnSumTable = new System.Windows.Forms.Button();
            this.txtSumTablePath = new System.Windows.Forms.TextBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnUpdat = new System.Windows.Forms.Button();
            this.dgvSow = new System.Windows.Forms.DataGridView();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txtManu = new System.Windows.Forms.TextBox();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnInsert = new System.Windows.Forms.Button();
            this.txtDpn = new System.Windows.Forms.TextBox();
            this.txtCategory = new System.Windows.Forms.TextBox();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.label11 = new System.Windows.Forms.Label();
            this.btnBrowsePo = new System.Windows.Forms.Button();
            this.txtPoPath = new System.Windows.Forms.TextBox();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.label17 = new System.Windows.Forms.Label();
            this.txtYHeight = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.txtYWidth = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtXHeight = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtXWidth = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label13 = new System.Windows.Forms.Label();
            this.btnBrowseTwoD = new System.Windows.Forms.Button();
            this.txtTowDPath = new System.Windows.Forms.TextBox();
            this.tabPage7 = new System.Windows.Forms.TabPage();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.btnBrowse2DTo1DNew = new System.Windows.Forms.Button();
            this.btnBrowse2DTo1DOld = new System.Windows.Forms.Button();
            this.txt2DTo1DNewPath = new System.Windows.Forms.TextBox();
            this.txt2DTo1DOldPath = new System.Windows.Forms.TextBox();
            this.btnExe = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.prgsBar = new System.Windows.Forms.ProgressBar();
            this.tabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSow)).BeginInit();
            this.tabPage5.SuspendLayout();
            this.tabPage6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.tabPage7.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl.Controls.Add(this.tabPage1);
            this.tabControl.Controls.Add(this.tabPage2);
            this.tabControl.Controls.Add(this.tabPage3);
            this.tabControl.Controls.Add(this.tabPage4);
            this.tabControl.Controls.Add(this.tabPage5);
            this.tabControl.Controls.Add(this.tabPage6);
            this.tabControl.Controls.Add(this.tabPage7);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Margin = new System.Windows.Forms.Padding(4);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1262, 748);
            this.tabControl.TabIndex = 1;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage1.Controls.Add(this.dateTimePicker);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnBrowseDell);
            this.tabPage1.Controls.Add(this.btnBrowseRMS);
            this.tabPage1.Controls.Add(this.txtFilter);
            this.tabPage1.Controls.Add(this.txtDellPath);
            this.tabPage1.Controls.Add(this.txtRmsPath);
            this.tabPage1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage1.Location = new System.Drawing.Point(4, 35);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage1.Size = new System.Drawing.Size(1254, 709);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "RMS to Dell";
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker.Location = new System.Drawing.Point(291, 363);
            this.dateTimePicker.Margin = new System.Windows.Forms.Padding(4);
            this.dateTimePicker.MaxDate = new System.DateTime(2200, 12, 31, 0, 0, 0, 0);
            this.dateTimePicker.MinDate = new System.DateTime(2018, 1, 1, 0, 0, 0, 0);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.ShowCheckBox = true;
            this.dateTimePicker.Size = new System.Drawing.Size(474, 30);
            this.dateTimePicker.TabIndex = 11;
            this.dateTimePicker.Value = new System.DateTime(2022, 7, 1, 0, 0, 0, 0);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(286, 333);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(360, 26);
            this.label12.TabIndex = 10;
            this.label12.Text = "Choose a verification date for TJ";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(286, 444);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(621, 81);
            this.label3.TabIndex = 7;
            this.label3.Text = "Enter project filter\r\n\r\n(\"None\" for all project. Use \",\" to separate multiple pro" +
    "jects)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(286, 216);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(177, 26);
            this.label2.TabIndex = 6;
            this.label2.Text = "Choose Dell file";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(286, 111);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(288, 26);
            this.label1.TabIndex = 5;
            this.label1.Text = "Choose RMS file(.csv,.xls)";
            // 
            // btnBrowseDell
            // 
            this.btnBrowseDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseDell.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseDell.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseDell.Location = new System.Drawing.Point(776, 243);
            this.btnBrowseDell.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowseDell.Name = "btnBrowseDell";
            this.btnBrowseDell.Size = new System.Drawing.Size(136, 38);
            this.btnBrowseDell.TabIndex = 4;
            this.btnBrowseDell.Text = "Browse";
            this.btnBrowseDell.UseVisualStyleBackColor = false;
            this.btnBrowseDell.Click += new System.EventHandler(this.BtnBrowseDell_Click);
            // 
            // btnBrowseRMS
            // 
            this.btnBrowseRMS.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseRMS.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseRMS.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseRMS.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseRMS.Location = new System.Drawing.Point(776, 136);
            this.btnBrowseRMS.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowseRMS.Name = "btnBrowseRMS";
            this.btnBrowseRMS.Size = new System.Drawing.Size(136, 38);
            this.btnBrowseRMS.TabIndex = 3;
            this.btnBrowseRMS.Text = "Browse";
            this.btnBrowseRMS.UseVisualStyleBackColor = false;
            this.btnBrowseRMS.Click += new System.EventHandler(this.BtnBrowseRMS_Click);
            // 
            // txtFilter
            // 
            this.txtFilter.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilter.Location = new System.Drawing.Point(290, 530);
            this.txtFilter.Margin = new System.Windows.Forms.Padding(4);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(475, 33);
            this.txtFilter.TabIndex = 2;
            this.txtFilter.Text = "None";
            // 
            // txtDellPath
            // 
            this.txtDellPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDellPath.Location = new System.Drawing.Point(290, 244);
            this.txtDellPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtDellPath.Name = "txtDellPath";
            this.txtDellPath.Size = new System.Drawing.Size(475, 33);
            this.txtDellPath.TabIndex = 1;
            // 
            // txtRmsPath
            // 
            this.txtRmsPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtRmsPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRmsPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtRmsPath.Location = new System.Drawing.Point(290, 138);
            this.txtRmsPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtRmsPath.Name = "txtRmsPath";
            this.txtRmsPath.Size = new System.Drawing.Size(475, 33);
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
            this.tabPage2.Location = new System.Drawing.Point(4, 35);
            this.tabPage2.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage2.Size = new System.Drawing.Size(1254, 709);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Dell merge";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(286, 111);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(240, 26);
            this.label4.TabIndex = 8;
            this.label4.Text = "Choose Dell file(.xlsx)";
            // 
            // btnBrowseMultiDell
            // 
            this.btnBrowseMultiDell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseMultiDell.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseMultiDell.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseMultiDell.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseMultiDell.Location = new System.Drawing.Point(776, 136);
            this.btnBrowseMultiDell.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowseMultiDell.Name = "btnBrowseMultiDell";
            this.btnBrowseMultiDell.Size = new System.Drawing.Size(136, 38);
            this.btnBrowseMultiDell.TabIndex = 7;
            this.btnBrowseMultiDell.Text = "Browse";
            this.btnBrowseMultiDell.UseVisualStyleBackColor = false;
            this.btnBrowseMultiDell.Click += new System.EventHandler(this.BtnBrowseMultiDell_Click);
            // 
            // txtDellMultiPath
            // 
            this.txtDellMultiPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtDellMultiPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDellMultiPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtDellMultiPath.Location = new System.Drawing.Point(290, 138);
            this.txtDellMultiPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtDellMultiPath.Name = "txtDellMultiPath";
            this.txtDellMultiPath.Size = new System.Drawing.Size(475, 33);
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
            this.tabPage3.Location = new System.Drawing.Point(4, 35);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage3.Size = new System.Drawing.Size(1254, 709);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "SAP format";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(286, 111);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(267, 26);
            this.label6.TabIndex = 11;
            this.label6.Text = "Choose SAP table(.xlsx)";
            // 
            // btnBrowseSap
            // 
            this.btnBrowseSap.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseSap.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseSap.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseSap.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseSap.Location = new System.Drawing.Point(776, 136);
            this.btnBrowseSap.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowseSap.Name = "btnBrowseSap";
            this.btnBrowseSap.Size = new System.Drawing.Size(136, 38);
            this.btnBrowseSap.TabIndex = 9;
            this.btnBrowseSap.Text = "Browse";
            this.btnBrowseSap.UseVisualStyleBackColor = false;
            this.btnBrowseSap.Click += new System.EventHandler(this.BtnBrowseSap_Click);
            // 
            // txtSapPath
            // 
            this.txtSapPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtSapPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSapPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtSapPath.Location = new System.Drawing.Point(290, 138);
            this.txtSapPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtSapPath.Name = "txtSapPath";
            this.txtSapPath.Size = new System.Drawing.Size(475, 33);
            this.txtSapPath.TabIndex = 7;
            // 
            // tabPage4
            // 
            this.tabPage4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage4.Controls.Add(this.label10);
            this.tabPage4.Controls.Add(this.btnSumTable);
            this.tabPage4.Controls.Add(this.txtSumTablePath);
            this.tabPage4.Controls.Add(this.btnDelete);
            this.tabPage4.Controls.Add(this.btnUpdat);
            this.tabPage4.Controls.Add(this.dgvSow);
            this.tabPage4.Controls.Add(this.label8);
            this.tabPage4.Controls.Add(this.label9);
            this.tabPage4.Controls.Add(this.txtManu);
            this.tabPage4.Controls.Add(this.txtDescription);
            this.tabPage4.Controls.Add(this.label5);
            this.tabPage4.Controls.Add(this.label7);
            this.tabPage4.Controls.Add(this.btnInsert);
            this.tabPage4.Controls.Add(this.txtDpn);
            this.tabPage4.Controls.Add(this.txtCategory);
            this.tabPage4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage4.Location = new System.Drawing.Point(4, 35);
            this.tabPage4.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage4.Size = new System.Drawing.Size(1254, 709);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "SOW to PDL";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(15, 561);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(319, 26);
            this.label10.TabIndex = 24;
            this.label10.Text = "Choose summary table(.xlsx)";
            // 
            // btnSumTable
            // 
            this.btnSumTable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnSumTable.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSumTable.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSumTable.ForeColor = System.Drawing.Color.Black;
            this.btnSumTable.Location = new System.Drawing.Point(1084, 555);
            this.btnSumTable.Margin = new System.Windows.Forms.Padding(4);
            this.btnSumTable.Name = "btnSumTable";
            this.btnSumTable.Size = new System.Drawing.Size(136, 38);
            this.btnSumTable.TabIndex = 23;
            this.btnSumTable.Text = "Browse";
            this.btnSumTable.UseVisualStyleBackColor = false;
            this.btnSumTable.Click += new System.EventHandler(this.BtnBrowseSumTable_Click);
            // 
            // txtSumTablePath
            // 
            this.txtSumTablePath.BackColor = System.Drawing.SystemColors.Window;
            this.txtSumTablePath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSumTablePath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtSumTablePath.Location = new System.Drawing.Point(350, 555);
            this.txtSumTablePath.Margin = new System.Windows.Forms.Padding(4);
            this.txtSumTablePath.Name = "txtSumTablePath";
            this.txtSumTablePath.Size = new System.Drawing.Size(724, 33);
            this.txtSumTablePath.TabIndex = 22;
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDelete.ForeColor = System.Drawing.Color.Black;
            this.btnDelete.Location = new System.Drawing.Point(1084, 160);
            this.btnDelete.Margin = new System.Windows.Forms.Padding(4);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(136, 45);
            this.btnDelete.TabIndex = 21;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.BtnDelete_Click);
            // 
            // btnUpdat
            // 
            this.btnUpdat.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnUpdat.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUpdat.ForeColor = System.Drawing.Color.Black;
            this.btnUpdat.Location = new System.Drawing.Point(939, 160);
            this.btnUpdat.Margin = new System.Windows.Forms.Padding(4);
            this.btnUpdat.Name = "btnUpdat";
            this.btnUpdat.Size = new System.Drawing.Size(136, 45);
            this.btnUpdat.TabIndex = 20;
            this.btnUpdat.Text = "Update";
            this.btnUpdat.UseVisualStyleBackColor = false;
            this.btnUpdat.Click += new System.EventHandler(this.BtnUpdat_Click);
            // 
            // dgvSow
            // 
            this.dgvSow.AllowUserToAddRows = false;
            this.dgvSow.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvSow.BackgroundColor = System.Drawing.SystemColors.ControlDarkDark;
            this.dgvSow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSow.GridColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.dgvSow.Location = new System.Drawing.Point(20, 218);
            this.dgvSow.Margin = new System.Windows.Forms.Padding(4);
            this.dgvSow.Name = "dgvSow";
            this.dgvSow.ReadOnly = true;
            this.dgvSow.RowHeadersWidth = 62;
            this.dgvSow.RowTemplate.Height = 24;
            this.dgvSow.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvSow.Size = new System.Drawing.Size(1202, 322);
            this.dgvSow.TabIndex = 19;
            this.dgvSow.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvSow_CellClick);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(656, 102);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(181, 26);
            this.label8.TabIndex = 18;
            this.label8.Text = "Manufacturers :";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(15, 100);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(149, 26);
            this.label9.TabIndex = 17;
            this.label9.Text = "Description :";
            // 
            // txtManu
            // 
            this.txtManu.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtManu.Location = new System.Drawing.Point(850, 96);
            this.txtManu.Margin = new System.Windows.Forms.Padding(4);
            this.txtManu.Name = "txtManu";
            this.txtManu.Size = new System.Drawing.Size(368, 33);
            this.txtManu.TabIndex = 14;
            // 
            // txtDescription
            // 
            this.txtDescription.BackColor = System.Drawing.SystemColors.Window;
            this.txtDescription.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDescription.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtDescription.Location = new System.Drawing.Point(165, 96);
            this.txtDescription.Margin = new System.Windows.Forms.Padding(4);
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(475, 33);
            this.txtDescription.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(664, 42);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(73, 26);
            this.label5.TabIndex = 12;
            this.label5.Text = "DPN :";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(40, 42);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(124, 26);
            this.label7.TabIndex = 11;
            this.label7.Text = "Category :";
            // 
            // btnInsert
            // 
            this.btnInsert.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnInsert.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnInsert.ForeColor = System.Drawing.Color.Black;
            this.btnInsert.Location = new System.Drawing.Point(747, 160);
            this.btnInsert.Margin = new System.Windows.Forms.Padding(4);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(136, 45);
            this.btnInsert.TabIndex = 10;
            this.btnInsert.Text = "Insert";
            this.btnInsert.UseVisualStyleBackColor = false;
            this.btnInsert.Click += new System.EventHandler(this.BtnInsert_Click);
            // 
            // txtDpn
            // 
            this.txtDpn.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDpn.Location = new System.Drawing.Point(747, 38);
            this.txtDpn.Margin = new System.Windows.Forms.Padding(4);
            this.txtDpn.Name = "txtDpn";
            this.txtDpn.Size = new System.Drawing.Size(301, 33);
            this.txtDpn.TabIndex = 8;
            // 
            // txtCategory
            // 
            this.txtCategory.BackColor = System.Drawing.SystemColors.Window;
            this.txtCategory.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCategory.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtCategory.Location = new System.Drawing.Point(165, 38);
            this.txtCategory.Margin = new System.Windows.Forms.Padding(4);
            this.txtCategory.Name = "txtCategory";
            this.txtCategory.Size = new System.Drawing.Size(367, 33);
            this.txtCategory.TabIndex = 7;
            // 
            // tabPage5
            // 
            this.tabPage5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage5.Controls.Add(this.label11);
            this.tabPage5.Controls.Add(this.btnBrowsePo);
            this.tabPage5.Controls.Add(this.txtPoPath);
            this.tabPage5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage5.Location = new System.Drawing.Point(4, 35);
            this.tabPage5.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(4);
            this.tabPage5.Size = new System.Drawing.Size(1254, 709);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "PO weekly report";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(286, 111);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(348, 26);
            this.label11.TabIndex = 14;
            this.label11.Text = "Choose PO weekly report(.xlsx)";
            // 
            // btnBrowsePo
            // 
            this.btnBrowsePo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowsePo.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowsePo.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowsePo.ForeColor = System.Drawing.Color.Black;
            this.btnBrowsePo.Location = new System.Drawing.Point(776, 136);
            this.btnBrowsePo.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowsePo.Name = "btnBrowsePo";
            this.btnBrowsePo.Size = new System.Drawing.Size(136, 38);
            this.btnBrowsePo.TabIndex = 13;
            this.btnBrowsePo.Text = "Browse";
            this.btnBrowsePo.UseVisualStyleBackColor = false;
            this.btnBrowsePo.Click += new System.EventHandler(this.BtnBrowsePo_Click);
            // 
            // txtPoPath
            // 
            this.txtPoPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtPoPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPoPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtPoPath.Location = new System.Drawing.Point(290, 138);
            this.txtPoPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtPoPath.Name = "txtPoPath";
            this.txtPoPath.Size = new System.Drawing.Size(475, 33);
            this.txtPoPath.TabIndex = 12;
            // 
            // tabPage6
            // 
            this.tabPage6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.tabPage6.Controls.Add(this.label17);
            this.tabPage6.Controls.Add(this.txtYHeight);
            this.tabPage6.Controls.Add(this.label16);
            this.tabPage6.Controls.Add(this.txtYWidth);
            this.tabPage6.Controls.Add(this.label15);
            this.tabPage6.Controls.Add(this.txtXHeight);
            this.tabPage6.Controls.Add(this.label14);
            this.tabPage6.Controls.Add(this.txtXWidth);
            this.tabPage6.Controls.Add(this.label18);
            this.tabPage6.Controls.Add(this.pictureBox1);
            this.tabPage6.Controls.Add(this.label13);
            this.tabPage6.Controls.Add(this.btnBrowseTwoD);
            this.tabPage6.Controls.Add(this.txtTowDPath);
            this.tabPage6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.tabPage6.Location = new System.Drawing.Point(4, 35);
            this.tabPage6.Margin = new System.Windows.Forms.Padding(4);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Size = new System.Drawing.Size(1254, 709);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.Text = "2D to 1D";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(582, 202);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(116, 26);
            this.label17.TabIndex = 35;
            this.label17.Text = "Y Height :";
            // 
            // txtYHeight
            // 
            this.txtYHeight.BackColor = System.Drawing.SystemColors.Window;
            this.txtYHeight.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtYHeight.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtYHeight.Location = new System.Drawing.Point(706, 198);
            this.txtYHeight.Margin = new System.Windows.Forms.Padding(4);
            this.txtYHeight.Name = "txtYHeight";
            this.txtYHeight.Size = new System.Drawing.Size(144, 33);
            this.txtYHeight.TabIndex = 34;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(226, 202);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(107, 26);
            this.label16.TabIndex = 33;
            this.label16.Text = "Y Width :";
            // 
            // txtYWidth
            // 
            this.txtYWidth.BackColor = System.Drawing.SystemColors.Window;
            this.txtYWidth.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtYWidth.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtYWidth.Location = new System.Drawing.Point(342, 198);
            this.txtYWidth.Margin = new System.Windows.Forms.Padding(4);
            this.txtYWidth.Name = "txtYWidth";
            this.txtYWidth.Size = new System.Drawing.Size(144, 33);
            this.txtYWidth.TabIndex = 32;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(582, 126);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(115, 26);
            this.label15.TabIndex = 31;
            this.label15.Text = "X Height :";
            // 
            // txtXHeight
            // 
            this.txtXHeight.BackColor = System.Drawing.SystemColors.Window;
            this.txtXHeight.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtXHeight.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtXHeight.Location = new System.Drawing.Point(706, 122);
            this.txtXHeight.Margin = new System.Windows.Forms.Padding(4);
            this.txtXHeight.Name = "txtXHeight";
            this.txtXHeight.Size = new System.Drawing.Size(144, 33);
            this.txtXHeight.TabIndex = 30;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(226, 126);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(106, 26);
            this.label14.TabIndex = 29;
            this.label14.Text = "X Width :";
            // 
            // txtXWidth
            // 
            this.txtXWidth.BackColor = System.Drawing.SystemColors.Window;
            this.txtXWidth.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtXWidth.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtXWidth.Location = new System.Drawing.Point(342, 122);
            this.txtXWidth.Margin = new System.Windows.Forms.Padding(4);
            this.txtXWidth.Name = "txtXWidth";
            this.txtXWidth.Size = new System.Drawing.Size(144, 33);
            this.txtXWidth.TabIndex = 28;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.BackColor = System.Drawing.Color.Lime;
            this.label18.Font = new System.Drawing.Font("微軟正黑體", 10F);
            this.label18.ForeColor = System.Drawing.Color.Red;
            this.label18.Location = new System.Drawing.Point(496, 27);
            this.label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(438, 25);
            this.label18.TabIndex = 27;
            this.label18.Text = "*請將要轉譯的二維資料放在 Excel 的第一個頁籤";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Anchor.Properties.Resources.TwoDParameter;
            this.pictureBox1.Location = new System.Drawing.Point(230, 250);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(783, 368);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 26;
            this.pictureBox1.TabStop = false;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(226, 28);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(264, 26);
            this.label13.TabIndex = 17;
            this.label13.Text = "Choose 2D report(.xlsx)";
            this.label13.Click += new System.EventHandler(this.BtnBrowseTwoD_Click);
            // 
            // btnBrowseTwoD
            // 
            this.btnBrowseTwoD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowseTwoD.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowseTwoD.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowseTwoD.ForeColor = System.Drawing.Color.Black;
            this.btnBrowseTwoD.Location = new System.Drawing.Point(876, 66);
            this.btnBrowseTwoD.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowseTwoD.Name = "btnBrowseTwoD";
            this.btnBrowseTwoD.Size = new System.Drawing.Size(136, 38);
            this.btnBrowseTwoD.TabIndex = 16;
            this.btnBrowseTwoD.Text = "Browse";
            this.btnBrowseTwoD.UseVisualStyleBackColor = false;
            this.btnBrowseTwoD.Click += new System.EventHandler(this.BtnBrowseTwoD_Click);
            // 
            // txtTowDPath
            // 
            this.txtTowDPath.BackColor = System.Drawing.SystemColors.Window;
            this.txtTowDPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTowDPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtTowDPath.Location = new System.Drawing.Point(230, 66);
            this.txtTowDPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtTowDPath.Name = "txtTowDPath";
            this.txtTowDPath.Size = new System.Drawing.Size(620, 33);
            this.txtTowDPath.TabIndex = 15;
            // 
            // tabPage7
            // 
            this.tabPage7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.tabPage7.Controls.Add(this.label19);
            this.tabPage7.Controls.Add(this.label20);
            this.tabPage7.Controls.Add(this.btnBrowse2DTo1DNew);
            this.tabPage7.Controls.Add(this.btnBrowse2DTo1DOld);
            this.tabPage7.Controls.Add(this.txt2DTo1DNewPath);
            this.tabPage7.Controls.Add(this.txt2DTo1DOldPath);
            this.tabPage7.Location = new System.Drawing.Point(4, 35);
            this.tabPage7.Name = "tabPage7";
            this.tabPage7.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage7.Size = new System.Drawing.Size(1254, 709);
            this.tabPage7.TabIndex = 6;
            this.tabPage7.Text = "1D Mapping";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label19.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.label19.Location = new System.Drawing.Point(314, 215);
            this.label19.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(344, 26);
            this.label19.TabIndex = 12;
            this.label19.Text = "Choose New 2D to 1D file(.xlsx)";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label20.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.label20.Location = new System.Drawing.Point(314, 110);
            this.label20.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(333, 26);
            this.label20.TabIndex = 11;
            this.label20.Text = "Choose Old 2D to 1D file(.xlsx)";
            // 
            // btnBrowse2DTo1DNew
            // 
            this.btnBrowse2DTo1DNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowse2DTo1DNew.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowse2DTo1DNew.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowse2DTo1DNew.ForeColor = System.Drawing.Color.Black;
            this.btnBrowse2DTo1DNew.Location = new System.Drawing.Point(804, 243);
            this.btnBrowse2DTo1DNew.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowse2DTo1DNew.Name = "btnBrowse2DTo1DNew";
            this.btnBrowse2DTo1DNew.Size = new System.Drawing.Size(136, 38);
            this.btnBrowse2DTo1DNew.TabIndex = 10;
            this.btnBrowse2DTo1DNew.Text = "Browse";
            this.btnBrowse2DTo1DNew.UseVisualStyleBackColor = false;
            this.btnBrowse2DTo1DNew.Click += new System.EventHandler(this.BtnBrowse2DTo1DNew_Click);
            // 
            // btnBrowse2DTo1DOld
            // 
            this.btnBrowse2DTo1DOld.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(253)))), ((int)(((byte)(203)))), ((int)(((byte)(82)))));
            this.btnBrowse2DTo1DOld.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowse2DTo1DOld.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowse2DTo1DOld.ForeColor = System.Drawing.Color.Black;
            this.btnBrowse2DTo1DOld.Location = new System.Drawing.Point(804, 136);
            this.btnBrowse2DTo1DOld.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowse2DTo1DOld.Name = "btnBrowse2DTo1DOld";
            this.btnBrowse2DTo1DOld.Size = new System.Drawing.Size(136, 38);
            this.btnBrowse2DTo1DOld.TabIndex = 9;
            this.btnBrowse2DTo1DOld.Text = "Browse";
            this.btnBrowse2DTo1DOld.UseVisualStyleBackColor = false;
            this.btnBrowse2DTo1DOld.Click += new System.EventHandler(this.BtnBrowse2DTo1DOld_Click);
            // 
            // txt2DTo1DNewPath
            // 
            this.txt2DTo1DNewPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt2DTo1DNewPath.Location = new System.Drawing.Point(318, 244);
            this.txt2DTo1DNewPath.Margin = new System.Windows.Forms.Padding(4);
            this.txt2DTo1DNewPath.Name = "txt2DTo1DNewPath";
            this.txt2DTo1DNewPath.Size = new System.Drawing.Size(475, 33);
            this.txt2DTo1DNewPath.TabIndex = 8;
            // 
            // txt2DTo1DOldPath
            // 
            this.txt2DTo1DOldPath.BackColor = System.Drawing.SystemColors.Window;
            this.txt2DTo1DOldPath.Font = new System.Drawing.Font("Arial Rounded MT Bold", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt2DTo1DOldPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txt2DTo1DOldPath.Location = new System.Drawing.Point(318, 138);
            this.txt2DTo1DOldPath.Margin = new System.Windows.Forms.Padding(4);
            this.txt2DTo1DOldPath.Name = "txt2DTo1DOldPath";
            this.txt2DTo1DOldPath.Size = new System.Drawing.Size(475, 33);
            this.txt2DTo1DOldPath.TabIndex = 7;
            // 
            // btnExe
            // 
            this.btnExe.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExe.Location = new System.Drawing.Point(962, 692);
            this.btnExe.Margin = new System.Windows.Forms.Padding(4);
            this.btnExe.Name = "btnExe";
            this.btnExe.Size = new System.Drawing.Size(136, 45);
            this.btnExe.TabIndex = 5;
            this.btnExe.Text = "Execute";
            this.btnExe.UseVisualStyleBackColor = true;
            this.btnExe.Click += new System.EventHandler(this.BtnExe_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(1107, 692);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(136, 45);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // prgsBar
            // 
            this.prgsBar.Location = new System.Drawing.Point(622, 696);
            this.prgsBar.Margin = new System.Windows.Forms.Padding(4);
            this.prgsBar.Name = "prgsBar";
            this.prgsBar.Size = new System.Drawing.Size(330, 39);
            this.prgsBar.TabIndex = 7;
            this.prgsBar.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(1262, 748);
            this.Controls.Add(this.prgsBar);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExe);
            this.Controls.Add(this.tabControl);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Anchor";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSow)).EndInit();
            this.tabPage5.ResumeLayout(false);
            this.tabPage5.PerformLayout();
            this.tabPage6.ResumeLayout(false);
            this.tabPage6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.tabPage7.ResumeLayout(false);
            this.tabPage7.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnBrowseRMS;
        private System.Windows.Forms.TextBox txtFilter;
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
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtManu;
        private System.Windows.Forms.TextBox txtDescription;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.TextBox txtDpn;
        private System.Windows.Forms.TextBox txtCategory;
        private System.Windows.Forms.DataGridView dgvSow;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnUpdat;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnSumTable;
        private System.Windows.Forms.TextBox txtSumTablePath;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnBrowsePo;
        private System.Windows.Forms.TextBox txtPoPath;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button btnBrowseDell;
        private System.Windows.Forms.TextBox txtDellPath;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnBrowseTwoD;
        private System.Windows.Forms.TextBox txtTowDPath;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox txtYHeight;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox txtYWidth;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtXHeight;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtXWidth;
        private System.Windows.Forms.TabPage tabPage7;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Button btnBrowse2DTo1DNew;
        private System.Windows.Forms.Button btnBrowse2DTo1DOld;
        private System.Windows.Forms.TextBox txt2DTo1DNewPath;
        private System.Windows.Forms.TextBox txt2DTo1DOldPath;
    }
}


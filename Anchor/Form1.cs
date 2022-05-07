using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Anchor
{
    public partial class Form1 : Form
    {
        OpenFileDialog opfd;
        SaveFileDialog sfd = new SaveFileDialog();
        public Form1()
        {
            InitializeComponent();
            opfd = new OpenFileDialog();
            opfd.RestoreDirectory = true;
            opfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private void BtnExe_Click(object sender, EventArgs e)
        {
            try
            {
                string saveFilePath = GetSaveFilePath();
                if (!string.IsNullOrEmpty(saveFilePath))
                {
                    switch (tabControl.SelectedIndex)
                    {
                        //RMS to Dell
                        case 0:

                            DellModel dm = new DellModel();
                            DataTable rmsDt = dm.RMSCsvToDataTable(txtRmsPath.Text.Trim(), ",");
                            List<DellModel> rmsList = dm.RMSDataTableToList(rmsDt);
                            List<DellModel> rmsColorList = dm.RMSListInputColor(rmsList);

                            List<DellModel> dellColorList;
                            if (!string.IsNullOrEmpty(txtDellPath.Text))
                            {
                                DataTable dellDt = dm.DellExcelToDataTable(txtDellPath.Text.Trim(), "Add Part", 5);
                                List<DellModel> dellList = dm.DellDataTableToList(dellDt);
                                dellColorList = dm.DellListInputColor(dellList);

                                //合併rms .csv & dell .xlsx
                                rmsColorList.AddRange(dellColorList);
                            }

                            

                            //staus: broken, sold filter
                            dm.StautsFilter(rmsColorList);

                            //project filter
                            Dictionary<string, object> dic = dm.PorjectFilter(rmsColorList, txtFilter.Text);

                            if (!"OK".Equals(dic["msg"]))
                                MessageBox.Show("Filter Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                //匯出excel
                                string msg = dm.ColorListToExcel((List<DellModel>)dic["returnObject"], saveFilePath);

                                if (msg.Equals("OK"))
                                    MessageBox.Show("Merge OK!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else
                                    MessageBox.Show("Merge Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            break;


                        //Dell Merge
                        case 1:

                            string filePaths = txtDellMultiPath.Text.Trim();
                            DellCellModel dmc = new DellCellModel();
                            List<List<DellCellModel>> rowList = dmc.CopySheetsToRowList(filePaths);


                            if (rowList.Count > 0)
                            {
                                dmc.RowListToExcel(rowList, saveFilePath);
                                MessageBox.Show("Merge OK!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                                MessageBox.Show("Merge Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void BtnBrowseRMS_Click(object sender, EventArgs e)
        {

            opfd.Filter = "CSV Files (*.csv)|*.csv";
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtRmsPath.Text = sFileName;
            }

        }

        private void BtnBrowseDell_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtDellPath.Text = sFileName;
            }
        }

        private void BtnBrowseMultiDell_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = true;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string[] sFileName = opfd.FileNames;
                txtDellMultiPath.Text = string.Join(";", sFileName);
            }

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            //this.Close();
            //StartProgress();
            //MessageBox.Show(GetSaveFilePath());
            DellModel dm = new DellModel();
            dm.DellExcelToDataTable(@"D:\Desktop\test.xls", 0, 1);

        }

        private void StartProgress()
        {
            int fileCount = 15;

            prgsBar.Visible = true;// 顯示進度條控件.
            prgsBar.Minimum = 0;// 設置進度條最小值.
            prgsBar.Maximum = fileCount;// 設置進度條最大值.
            prgsBar.Value = 1;// 設置進度條初始值
            prgsBar.Step = 1;// 設置每次增加的步長
            Graphics g = this.prgsBar.CreateGraphics();//創建Graphics對象

            for (int x = 1; x <= fileCount; x++)
            {
                //執行PerformStep()函數
                prgsBar.PerformStep();
                string str = Math.Round((100 * x / 15.0), 2).ToString("#0.00 ") + "%";
                Font font = new Font("Times New Roman", (float)10, FontStyle.Regular);
                PointF pt = new PointF(this.prgsBar.Width / 2 - 17, this.prgsBar.Height / 2 - 7);
                g.DrawString(str, font, Brushes.Blue, pt);
                //每次循環讓程序休眠300毫秒
                System.Threading.Thread.Sleep(1000);
            }
            //prgsBar.Visible = false;
            //MessageBox.Show("success!");
        }
        private string GetSaveFilePath()
        {
            sfd.Filter = "Excel Files|*.xlsx"; ;//設定檔案型別
            sfd.FileName = "Inventory report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss"));//設定預設檔名
            sfd.DefaultExt = "xlsx";//設定預設格式（可以不設）
            sfd.AddExtension = true;//設定自動在檔名中新增副檔名
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            else return "";
        }
    }
}

using Anchor.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Font = System.Drawing.Font;

namespace Anchor
{
    public partial class Form1 : Form
    {
        DataTable dt = new DataTable();
        int gdvSowIndex;
        OpenFileDialog opfd;
        SaveFileDialog sfd = new SaveFileDialog();
        public Form1()
        {
            InitializeComponent();
            opfd = new OpenFileDialog();
            opfd.RestoreDirectory = true;
            opfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //SOW 先隱藏
            this.tabControl.TabPages.Remove(this.tabPage4);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dt.Columns.Add("Category", Type.GetType("System.String"));
            dt.Columns.Add("DPN", Type.GetType("System.String"));
            dt.Columns.Add("Description", Type.GetType("System.String"));
            dt.Columns.Add("Manufacturers", Type.GetType("System.String"));
            dgvSow.DataSource = dt;

        }

        private void BtnExe_Click(object sender, EventArgs e)
        {
            string message;
            try
            {
                //前兩個Dell 相關的 tab 使用 Inventory 開頭的預設檔名
                string saveFilePath;
                if (tabControl.SelectedIndex < 2)
                    saveFilePath = GetSaveFilePath();
                else if (tabControl.SelectedIndex == 6)
                    saveFilePath = GetSaveFilePathForDellRandomCheck();
                else saveFilePath = GetSaveFilePathWithOutInventory();

                if (!string.IsNullOrEmpty(saveFilePath))
                {
                    switch (tabControl.SelectedIndex)
                    {
                        //RMS to Dell
                        #region
                        case 0:
                            DellModel dm = new DellModel();
                            List<string> newRmsPathList = new List<string>();
                            List<DellModel> rmsList = new List<DellModel>();
                            string[] tokens = txtRmsPath.Text.Split(';');
                            string tjValidDate = dateTimePicker.Value.ToString("yyyy/MM/dd 00:00:00");

                            foreach (string fp in tokens)
                            {
                               if (fp.Contains(".xlsx"))
                                    newRmsPathList.Add(fp);
                            }


                            if (newRmsPathList != null)
                            {
                                foreach (string nrpl in newRmsPathList)
                                {
                                    DataTable newRmsDt = dm.ExcelToDataTable(nrpl.Trim(), 0, 1);
                                    rmsList.AddRange(dm.NewRmsDataTableToList(newRmsDt));
                                }
                            }


                            List<DellModel> rmsColorList = dm.RMSListInputColor(rmsList);

                            List<DellModel> dellColorList;
                            //如果沒選Dell excel 直接轉RMS
                            if (!string.IsNullOrEmpty(txtDellPath.Text))
                            {
                                DataTable dellDt = dm.ExcelToDataTable(txtDellPath.Text.Trim(), "Add Part", 5);
                                List<DellModel> dellList = dm.DellDataTableToList(dellDt);
                                dellColorList = dm.DellListInputColor(dellList);

                                //合併rms & dell .xlsx
                                rmsColorList = dm.CreateAndUpdateDellFile(rmsColorList, dellColorList, tjValidDate);
                                //rmsColorList.AddRange(dellColorList);
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

                        #endregion

                        //Dell Merge
                        #region
                        case 1:
                            string filePaths = txtDellMultiPath.Text.Trim();
                            DellMergeModel dmc = new DellMergeModel();
                            List<List<DellMergeModel>> rowList = dmc.CopySheetsToRowList(filePaths);


                            if (rowList.Count > 0)
                            {
                                dmc.RowListToExcel(rowList, saveFilePath);
                                MessageBox.Show("Merge OK!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                                MessageBox.Show("Merge Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        #endregion

                        //SAP format
                        #region
                        case 2:
                            SapModel sm = new SapModel();
                            DataTable dt = sm.ExcelToDataTable(txtSapPath.Text.Trim(), 0, 1);
                            List<SapModel> sapList = sm.SapDataTableToList(dt);

                            //匯出excel
                            message = sm.ListToExcel(sapList, saveFilePath);

                            if (message.Equals("OK"))
                                MessageBox.Show("Done", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show("Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            break;
                        #endregion

                        //PO weekly report
                        #region
                        case 3:
                            string filePath = txtPoPath.Text.Trim();
                            PoModel pm = new PoModel();
                            message = pm.AnalysisWeeklyReport(filePath, saveFilePath);

                            if (message.Equals("OK"))
                                MessageBox.Show("Done", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show(message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            break;
                        #endregion

                        //2D report
                        #region
                        case 4:

                            TwoDimensionModel tdm = new TwoDimensionModel(
                                 Int32.Parse(txtXWidth.Text),
                                 Int32.Parse(txtXHeight.Text),
                                 Int32.Parse(txtYWidth.Text),
                                 Int32.Parse(txtYHeight.Text)
                            );
                            message = tdm.ConvertTwoDimensionData(txtTowDPath.Text.Trim(), saveFilePath);

                            if (message.Equals("OK"))
                                MessageBox.Show("Done", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show(message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            break;
                        #endregion

                        //1D mapping
                        #region
                        case 5:
                            //Do somthing
                            OneDimensionMappingModel odm = new OneDimensionMappingModel();
                            string oldFilePath = txt2DTo1DOldPath.Text.Trim();
                            string newFilePath = txt2DTo1DNewPath.Text.Trim();
                            message = odm.MappingOneDimensionData(oldFilePath, newFilePath, saveFilePath);

                            if (message.Equals("OK"))
                                MessageBox.Show("DONE", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            else
                                MessageBox.Show(message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        #endregion

                        //RMS to Dell ( New )
                        #region
                        case 6:

                            RmsModel rm = new RmsModel();
                            newRmsPathList = new List<string>();
                            List<RmsModel> newRmsList = new List<RmsModel>();
                            tokens = txtRmsPathNew.Text.Split(';');

                            foreach (string fp in tokens)
                            {
                                if (fp.Contains(".xlsx"))
                                    newRmsPathList.Add(fp);
                            }

                            if (newRmsPathList != null)
                            {
                                foreach (string nrpl in newRmsPathList)
                                {
                                    label21.Text = "STEP 1. 正在載入資料 : " + nrpl;
                                    label21.Refresh();
                                    DataTable newRmsDt = rm.ExcelToDataTable(nrpl.Trim(), 0, 1);
                                    newRmsList.AddRange(rm.RmsDataTableToModelList(newRmsDt));
                                }
                            }


                            label21.Text = "STEP 2. 過濾所需資料 ( 從 " + newRmsList.Count + " 列資料過濾 )";
                            label21.Refresh();
                            dic = rm.PorjectFilter(newRmsList);

                            if (!"OK".Equals(dic["msg"]))
                                MessageBox.Show("Filter Fail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                List<RmsModel> finalList = (List<RmsModel>)dic["returnObject"];
                                label21.Text = "STEP 3. 匯出指定格式 ( 匯出 " + finalList.Count + " 列資料)";
                                label21.Refresh();
                                //匯出excel
                                string msg = rm.ColorListToExcel(finalList, saveFilePath);

                                if ("OK".Equals(msg))
                                    MessageBox.Show("OOOK!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else
                                    MessageBox.Show("FFFFail!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            label21.Text = "STEP 4. 執行結束";
                            break;

                            #endregion

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        private void BtnBrowseRMSNew_Click(object sender, EventArgs e)
        {

            opfd.Filter = "RMS files (*.xlsx)|*.xlsx";
            opfd.Multiselect = true;
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string[] sFileName = opfd.FileNames;
                txtRmsPathNew.Text = string.Join(";", sFileName);
            }

        }

        private void BtnBrowseRMS_Click(object sender, EventArgs e)
        {

            opfd.Filter = "RMS files (*.xlsx)|*.xlsx";
            opfd.Multiselect = true;
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string[] sFileName = opfd.FileNames;
                txtRmsPath.Text = string.Join(";", sFileName);
            }

        }
        private void BtnBrowseDell_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;
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
        private void BtnBrowseSap_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtSapPath.Text = sFileName;
            }
        }
        private void BtnBrowsePo_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtPoPath.Text = sFileName;
            }

        }
        private void BtnBrowseTwoD_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtTowDPath.Text = sFileName;
            }

        }
        private void BtnBrowseSumTable_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txtSumTablePath.Text = sFileName;
            }
        }

        private void DgvSow_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            gdvSowIndex = e.RowIndex;
            DataGridViewRow row = dgvSow.Rows[gdvSowIndex];
            txtCategory.Text = row.Cells[0].Value.ToString();
            txtDpn.Text = row.Cells[1].Value.ToString();
            txtDescription.Text = row.Cells[2].Value.ToString();
            txtManu.Text = row.Cells[3].Value.ToString();

        }
        private void BtnInsert_Click(object sender, EventArgs e)
        {
            dt.Rows.Add(txtCategory.Text, txtDpn.Text, txtDescription.Text, txtManu.Text);

        }
        private void BtnUpdat_Click(object sender, EventArgs e)
        {
            DataGridViewRow newData = dgvSow.Rows[gdvSowIndex];
            newData.Cells[0].Value = txtCategory.Text;
            newData.Cells[1].Value = txtDpn.Text;
            newData.Cells[2].Value = txtDescription.Text;
            newData.Cells[3].Value = txtManu.Text;
        }
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            gdvSowIndex = dgvSow.CurrentCell.RowIndex;
            dgvSow.Rows.RemoveAt(gdvSowIndex);
        }

        private string GetSaveFilePath()
        {
            sfd.Filter = "Excel Files|*.xlsx"; ;//設定檔案型別
            sfd.FileName = "Inventory_report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss"));//設定預設檔名
            sfd.DefaultExt = "xlsx";//設定預設格式（可以不設）
            sfd.AddExtension = true;//設定自動在檔名中新增副檔名
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine(sfd.FileName);
                return sfd.FileName;
            }
            else return "";
        }
        private string GetSaveFilePathWithOutInventory()
        {
            sfd.Filter = "Excel Files|*.xlsx"; ;//設定檔案型別
            sfd.FileName = "Report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss"));//設定預設檔名
            sfd.DefaultExt = "xlsx";//設定預設格式（可以不設）
            sfd.AddExtension = true;//設定自動在檔名中新增副檔名
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            else return "";
        }
        private string GetSaveFilePathForDellRandomCheck()
        {
            sfd.Filter = "Excel Files|*.xlsx"; ;//設定檔案型別
            sfd.FileName = "Foxconn external audit report for " + DateTime.Now.ToString(("yyyy.MM"));//設定預設檔名
            sfd.DefaultExt = "xlsx";//設定預設格式（可以不設）
            sfd.AddExtension = true;//設定自動在檔名中新增副檔名
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }
            else return "";
        }
        private void NumberTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
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

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            //this.Close();
            //StartProgress();

            //DataTable dt = sm.WordSowToDataTable(@"D:\Desktop\Jennifer_test\YETI MAY FY23 SOW v0.4.docx");
            //List<SowModel> list = sm.WrodSowDataTableToList(dt);
            //sm.ListToExcel(list, @"D:\Desktop\Jennifer_test\xxxxx.xlsx", @"docs\Sample_polling_format.xlsx", "Allocation polling-project", 3);
            //sm.ExcelToDataTable(@"D:\Desktop\Jennifer_test\merge.xlsx", "Allocation polling-project", 3);
            //sm.DoCopyRange();
            //sm.DoInsertRow();
            //string str = dateTimePicker.Value.ToString("yyyy/MM/dd 00:00:00");
            //MessageBox.Show(str);
            //if ((DateTime.Parse(str) > DateTime.Parse("2021/8/31  09:51:27 AM")))
            //    MessageBox.Show(">");
            //else MessageBox.Show("<");
        }

        private void BtnBrowse2DTo1DOld_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txt2DTo1DOldPath.Text = sFileName;
            }

        }

        private void BtnBrowse2DTo1DNew_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = false;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                txt2DTo1DNewPath.Text = sFileName;
            }

        }
    }
}

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

namespace SyncFormsApp
{
    public partial class Form1 : Form
    {
        OpenFileDialog opfd;
        public Form1()
        {
            InitializeComponent();
            opfd = new OpenFileDialog();
            opfd.RestoreDirectory = true;
            opfd.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop);




        }
       
        private void BtnExe_Click(object sender, EventArgs e)
        {
            DellModelCell dmc = new DellModelCell();
            string filePaths = TxtDellMultiPath.Text.Trim();
            string[] tokens = filePaths.Split(';');

            foreach (string st in tokens)
            {
                dmc.CopySheetsToList(st);
            }
            
        }

        private void BtnBrowseRMS_Click(object sender, EventArgs e)
        {

            opfd.Filter = "CSV Files (*.csv)|*.csv";
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                TxtRmsPath.Text = sFileName;
            }

        }

        private void BtnBrowseDell_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string sFileName = opfd.FileName;
                TxtDellPath.Text = sFileName;
            }
        }

        private void BtnBrowseMultiDell_Click(object sender, EventArgs e)
        {
            opfd.Filter = "Excel Files|*.xlsx";
            opfd.Multiselect = true;

            if (opfd.ShowDialog() == DialogResult.OK)
            {
                string[] sFileName = opfd.FileNames;
                TxtDellMultiPath.Text = string.Join(";", sFileName);
            }

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            String FilePath = @"D:\Desktop\(Dell)Inventory Report_Mar.xlsx";
            IWorkbook workbook = null;
            FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);

            if (FilePath.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook(fs);

            else if (FilePath.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook(fs);

            ISheet sheet = workbook.GetSheet("Add Part");


        }
    }
}

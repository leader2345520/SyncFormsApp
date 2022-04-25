using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SyncFormsApp
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    class DellModelCell
    {
        public string Value { get; set; }
        public Byte[] FontColor { get; set; }


        public List<List<DellModelCell>> CopySheetsToList(string FilePath) {

            DellModelCell dmc;
            List<DellModelCell> cellList;
            List<List<DellModelCell>> rowList = new List<List<DellModelCell>>();

            try
            {
                //IWorkbook => excel file
                IWorkbook workbook = null;
                FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read);


                if (FilePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);

                else if (FilePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                //ISheet => sheet
                ISheet sheet = workbook.GetSheet("Add Part");               


                if (sheet != null)
                {
                    //LastRowNum => total count
                    int rowCount = sheet.LastRowNum;
                    for (int i = 5; i < rowCount + 1; i++)
                    {
                        // IRow => cursor
                        IRow curRow = sheet.GetRow(i);

                        cellList = new List<DellModelCell>();

                        for (int j = 0; j < curRow.LastCellNum; j++)
                        {
                            // ICell => column
                            ICell icell = curRow.GetCell(j);
                            icell.SetCellType(CellType.String);
                            var cellValue = icell.StringCellValue;


                            // IFont => colum style
                            IFont ifont = icell.CellStyle.GetFont(workbook);
                            Byte[] fontColor = (ifont as XSSFFont).GetXSSFColor().RGB;

                            dmc = new DellModelCell();
                            dmc.Value = cellValue;
                            dmc.FontColor = fontColor;
                            cellList.Add(dmc);

                        }
                        rowList.Add(cellList);
                    }           
                }           
            }

            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                MessageBox.Show(exception.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return rowList;
        }

    }

}

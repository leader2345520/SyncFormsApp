using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using static NPOI.HSSF.Util.HSSFColor;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;

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
        public string Value { get; set; } = "";
        public Byte[] FontColor { get; set; } = {0, 0, 0 };


        public string RowListToExcel(List<List<DellModelCell>> rowList)
        {

            //根據指定的檔案格式建立對應的類
            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin");
            using (FileStream templateFile = new FileStream("format.xlsx", FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.GetSheet("Add Part");
                int copyFormatRow = 5;


              
                //設定處存格 style(黑色)
                XSSFCellStyle cellStyle_black = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_black.BorderBottom = BorderStyle.Thin;
                cellStyle_black.BorderLeft = BorderStyle.Thin;
                cellStyle_black.BorderRight = BorderStyle.Thin;
                cellStyle_black.BorderTop = BorderStyle.Thin;

                //設定字體sytle
                XSSFFont font_black = (XSSFFont)workbook.CreateFont();
                font_black.FontName = "Calibri";//字體
                font_black.SetColor(new XSSFColor(new Byte[] { 0, 0, 0 }));
                cellStyle_black.SetFont(font_black);


                //設定處存格 style(藍色)
                XSSFCellStyle cellStyle_blue = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_blue.BorderBottom = BorderStyle.Thin;
                cellStyle_blue.BorderLeft = BorderStyle.Thin;
                cellStyle_blue.BorderRight = BorderStyle.Thin;
                cellStyle_blue.BorderTop = BorderStyle.Thin;

                //設定字體sytle
                XSSFFont font_blue = (XSSFFont)workbook.CreateFont();
                font_blue.FontName = "Calibri";//字體
                font_blue.SetColor(new XSSFColor( new Byte[]{ 0,0,255 }));

                cellStyle_blue.SetFont(font_blue);

                //設定處存格 style(紅色)
                XSSFCellStyle cellStyle_red = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_red.BorderBottom = BorderStyle.Thin;
                cellStyle_red.BorderLeft = BorderStyle.Thin;
                cellStyle_red.BorderRight = BorderStyle.Thin;
                cellStyle_red.BorderTop = BorderStyle.Thin;

                //設定字體sytle
                XSSFFont font_red = (XSSFFont)workbook.CreateFont();
                font_red.FontName = "Calibri";//字體
                font_red.SetColor(new XSSFColor(new Byte[] { 255, 0, 0 }));

                cellStyle_red.SetFont(font_red);



                IRow row = null;
                ICell cell = null;



                //寫入資料
                for (int i = 0; i < rowList.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;

                    foreach (DellModelCell dmc in rowList[i])
                    {
                        cell = row.CreateCell(cellIndex);
                        cell.SetCellValue(dmc.Value);

                        //逐格設定顏色, CellStyle超過65000個就會爆掉
                        //font.SetColor(new XSSFColor(dmc.FontColor));

                        if (BitConverter.ToString(dmc.FontColor) == BitConverter.ToString(new byte[]{ 0, 0, 0}))
                            cell.CellStyle = cellStyle_black;                        
                        else if ((BitConverter.ToString(dmc.FontColor) == BitConverter.ToString(new byte[] { 0, 0, 255 }))) 
                            cell.CellStyle = cellStyle_blue;                    
                        else 
                            cell.CellStyle = cellStyle_red;

                        cellIndex++;

                    }
                  
                }

                //寫檔
                using (FileStream file = new FileStream("Inventory report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss")) + ".xlsx", FileMode.Create))//產生檔案
                {
                    workbook.Write(file);
                }
            }


            return "OK";
        }

        public List<List<DellModelCell>> CopySheetsToRowList(string FilePath) {

            DellModelCell dmc;
            List<DellModelCell> cellList;
            List<List<DellModelCell>> rowList = new List<List<DellModelCell>>();

            string[] tokens = FilePath.Split(';');

            foreach (string fp in tokens)
            {

                try
                {
                    //IWorkbook => excel file
                    IWorkbook workbook = null;
                    FileStream fs = new FileStream(fp, FileMode.Open, FileAccess.Read);


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
                                dmc = new DellModelCell();
                                if (icell == null)
                                {
                                    cellList.Add(dmc); continue;
                                }

                                icell.SetCellType(CellType.String);
                                var cellValue = icell.StringCellValue;

                                // IFont => colum style
                                IFont ifont = icell.CellStyle.GetFont(workbook);
                                if ((ifont as XSSFFont).GetXSSFColor() == null)
                                {
                                    cellList.Add(dmc); continue;
                                }
                                Byte[] fontColor = (ifont as XSSFFont).GetXSSFColor().RGB;
                           

                                
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
                    break;
                }
            }

            return rowList;
        }

    }

}

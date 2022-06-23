using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;

namespace Anchor
{
    class DellMergeModel : BaseModel
    {
        #region getter/setter
        public string Value { get; set; } = "";
        public byte[] FontColor { get; set; } = { 0, 0, 0 };
        #endregion

        public List<List<DellMergeModel>> CopySheetsToRowList(string filePath)
        {

            DellMergeModel dmm;
            List<DellMergeModel> cellList;
            List<List<DellMergeModel>> rowList = new List<List<DellMergeModel>>();
            int copyFormatRow = 5;  //Head
            string[] tokens = filePath.Split(';');

            foreach (string fp in tokens)
            {

                try
                {
                    //IWorkbook => excel file
                    IWorkbook workbook = null;
                    FileStream fs = new FileStream(fp, FileMode.Open, FileAccess.Read);


                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);

                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    //ISheet => sheet
                    ISheet sheet = workbook.GetSheet("Add Part");

                    if (sheet != null)
                    {
                        //LastRowNum => total count
                        int rowCount = sheet.LastRowNum;
                        for (int i = copyFormatRow; i < rowCount + 1; i++)
                        {
                            // IRow => cursor
                            IRow curRow = sheet.GetRow(i);

                            //by pass 空白行
                            if (IsRowEmpty(curRow)) break;

                            cellList = new List<DellMergeModel>();

                            for (int j = 0; j < curRow.LastCellNum; j++)
                            {
                                // ICell => column
                                ICell icell = curRow.GetCell(j);
                                dmm = new DellMergeModel();
                                if (icell == null)
                                {
                                    cellList.Add(dmm);
                                    continue;
                                }

                                icell.SetCellType(CellType.String);
                                var cellValue = GetValueFromCell(icell);

                                // IFont => colum style
                                IFont ifont = icell.CellStyle.GetFont(workbook);
                                byte[] fontColor;

                                // 如果找不到顏色有可能就是 "自動", 當做是黑色處理
                                if ((ifont as XSSFFont).GetXSSFColor() == null)
                                    fontColor = new byte[] { 0, 0, 0 };
                                else
                                    fontColor = (ifont as XSSFFont).GetXSSFColor().RGB;

                                dmm.Value = cellValue;
                                dmm.FontColor = fontColor;
                                cellList.Add(dmm);

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
        public string RowListToExcel(List<List<DellMergeModel>> rowList, string saveFilePath)
        {
            //根據指定的檔案格式建立對應的類
            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin");

            // 使用'using' 自動做Flush()、Close()、Dispose() 釋放資源
            using (FileStream templateFile = new FileStream(@"docs\format.xlsx", FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.GetSheet("Add Part");
                int copyFormatRow = 5;  //Head                

                #region 藍色模板
                //設定處存格 style(藍色)
                XSSFCellStyle cellStyle_blue = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_blue.BorderBottom = BorderStyle.Thin;
                cellStyle_blue.BorderLeft = BorderStyle.Thin;
                cellStyle_blue.BorderRight = BorderStyle.Thin;
                cellStyle_blue.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_blue = (XSSFFont)workbook.CreateFont();
                font_blue.FontName = "Calibri";//字體
                font_blue.SetColor(new XSSFColor(new byte[] { 0, 0, 255 }));
                cellStyle_blue.SetFont(font_blue);
                #endregion

                #region 黑色模板
                //設定處存格 style(黑色)
                XSSFCellStyle cellStyle_black = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_black.BorderBottom = BorderStyle.Thin;
                cellStyle_black.BorderLeft = BorderStyle.Thin;
                cellStyle_black.BorderRight = BorderStyle.Thin;
                cellStyle_black.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_black = (XSSFFont)workbook.CreateFont();
                font_black.FontName = "Calibri";//字體
                font_black.SetColor(new XSSFColor(new byte[] { 0, 0, 0 }));
                cellStyle_black.SetFont(font_black);
                #endregion

                #region 紅色模板
                //設定處存格 style(紅色)
                XSSFCellStyle cellStyle_red = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_red.BorderBottom = BorderStyle.Thin;
                cellStyle_red.BorderLeft = BorderStyle.Thin;
                cellStyle_red.BorderRight = BorderStyle.Thin;
                cellStyle_red.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_red = (XSSFFont)workbook.CreateFont();
                font_red.FontName = "Calibri";//字體
                font_red.SetColor(new XSSFColor(new byte[] { 255, 0, 0 }));
                cellStyle_red.SetFont(font_red);
                #endregion

                IRow row = null;
                ICell cell = null;

                //寫入資料
                for (int i = 0; i < rowList.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;

                    foreach (DellMergeModel dmc in rowList[i])
                    {
                        //放值
                        cell = row.CreateCell(cellIndex);
                        cell.SetCellValue(dmc.Value);

                        //逐格設定顏色, CellStyle超過65000個就會爆掉
                        //黑藍共用模板
                        if (BitConverter.ToString(dmc.FontColor) == BitConverter.ToString(new byte[] { 0, 0, 0 }))
                            cell.CellStyle = cellStyle_black;
                        else if ((BitConverter.ToString(dmc.FontColor) == BitConverter.ToString(new byte[] { 0, 0, 255 })))
                            cell.CellStyle = cellStyle_blue;
                        else if ((BitConverter.ToString(dmc.FontColor) == BitConverter.ToString(new byte[] { 255, 0, 0 })))
                            cell.CellStyle = cellStyle_red;
                        else
                        {
                            //設定處存格 style(客製)
                            XSSFCellStyle cellStyle_cust = (XSSFCellStyle)workbook.CreateCellStyle();
                            cellStyle_cust.BorderBottom = BorderStyle.Thin;
                            cellStyle_cust.BorderLeft = BorderStyle.Thin;
                            cellStyle_cust.BorderRight = BorderStyle.Thin;
                            cellStyle_cust.BorderTop = BorderStyle.Thin;

                            //設定字體sytle
                            XSSFFont font_cust = (XSSFFont)workbook.CreateFont();
                            font_cust.FontName = "Calibri";//字體
                            font_cust.SetColor(new XSSFColor(dmc.FontColor));

                            cellStyle_cust.SetFont(font_cust);
                            cell.CellStyle = cellStyle_cust;
                        }

                        cellIndex++;

                    }

                }

                //寫檔
                string _saveFilePath = "".Equals(saveFilePath) ? "Inventory report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss")) + ".xlsx" : saveFilePath;
                SaveWorkbook(workbook, _saveFilePath);
            }


            return "OK";
        }
        public bool IsRowEmpty(IRow row)
        {
            if (row == null) return true;

            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell icell = row.GetCell(i);
                if (icell != null && icell.CellType != CellType.Blank)
                {
                    return false;
                }
            }
            return true;
        }



    }
}



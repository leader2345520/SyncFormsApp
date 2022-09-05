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

namespace Anchor.Model
{
    class TwoDimensionModel : BaseModel
    {
        public List<List<string>> XLists = new List<List<string>>();
        public List<List<string>> YLists = new List<List<string>>();
        public List<List<string>> ResultLists = new List<List<string>>();
        public int XWidth;
        public int XHeight;
        public int YWidth;
        public int YHeight;

        public TwoDimensionModel(int _XWidth, int _XHeight, int _YWidth, int _YHeight)
        {
            this.XWidth = _XWidth;
            this.XHeight = _XHeight;
            this.YWidth = _YWidth;
            this.YHeight = _YHeight;
        }


        public string ConvertTwoDimensionData(string filePath, string savePath)
        {

            try
            {
                //IWorkbook => excel file
                IWorkbook workbook = null;
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);

                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                //ISheet => sheet
                ISheet sheet = workbook.GetSheetAt(0);
                SetXList(sheet);
                SetYList(sheet);

                IRow curRow;
                ICell icell;

                for (int i = 0; i < XLists.Count(); i++)
                {
                    for (int j = 0; j < YLists.Count(); j++)
                    {
                        curRow = sheet.GetRow(XHeight + j); //往下取變動的數量
                        icell = curRow.GetCell(YWidth + i);
                        var cellValue = GetValueFromCell(icell);

                        List<string> resultList = YLists[j].Concat(XLists[i]).ToList<string>();
                        resultList.Add(cellValue);
                        ResultLists.Add(resultList);
                    }
                }

                ResultListsToExcel(ResultLists, filePath, savePath);

                return "OK";
            }

            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                MessageBox.Show(exception.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return exception.Message;
            }
        }
        public void SetXList(ISheet sheet)
        {

            if (sheet != null)
            {
                //LastRowNum => total count
                int rowCount = sheet.LastRowNum;
                for (int i = 0; i < XHeight; i++)
                {
                    // IRow => cursor
                    IRow curRow = sheet.GetRow(i);
                    int xCount = 0;
                    for (int j = YWidth; j < YWidth + XWidth; j++)
                    {
                        ICell icell = curRow.GetCell(j);
                        var cellValue = GetValueFromCell(icell);

                        if (i == 0)
                        {
                            List<string> list = new List<string>();
                            list.Add(cellValue);
                            XLists.Add(list);
                        }
                        else
                        {
                            XLists[xCount].Add(cellValue);
                            xCount++;
                        }
                    }
                }
            }
        }
        public void SetYList(ISheet sheet)
        {
            if (sheet != null)
            {
                //LastRowNum => total count
                int rowCount = sheet.LastRowNum;
                for (int i = XHeight; i < XHeight + YHeight; i++)
                {
                    // IRow => cursor
                    IRow curRow = sheet.GetRow(i);
                    List<string> list = new List<string>();
                    for (int j = 0; j < YWidth; j++)
                    {
                        ICell icell = curRow.GetCell(j);
                        var cellValue = GetValueFromCell(icell);
                        list.Add(cellValue);

                    }
                    YLists.Add(list);

                }
            }
        }
        public string ResultListsToExcel(List<List<string>> resultLists, string filePath, string saveFilePath)
        {
            // 使用'using' 自動做Flush()、Close()、Dispose() 釋放資源
            using (FileStream templateFile = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.CreateSheet("OneDimensionDataResult");

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

                IRow row = null;
                ICell cell = null;

                //寫入資料
                for (int i = 0; i < resultLists.Count; i++)
                {
                    row = sheet.CreateRow(i);
                    int cellIndex = 0;

                    foreach (string str in resultLists[i])
                    {
                        //放值
                        cell = row.CreateCell(cellIndex);
                        cell.SetCellValue(str);

                        //逐格設定顏色, CellStyle超過65000個就會爆掉
                        cell.CellStyle = cellStyle_black;
                        cellIndex++;
                    }
                }

                //寫檔
                SaveWorkbook(workbook, saveFilePath);

            }

            return "OK";
        }
    }

}

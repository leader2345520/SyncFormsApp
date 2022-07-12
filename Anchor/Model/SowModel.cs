using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Anchor
{
    class SowModel : BaseModel
    {
        #region getter/setter
        public string Category { get; set; }
        public string DPN { get; set; }
        public string Description { get; set; }
        public string Manufacturers { get; set; }

        #endregion

        //public DataTable WordSowToDataTable(string fileName)
        //{
        //    DataTable dt = new DataTable();
        //    DataColumnCollection columns;
        //    bool isFirst = true;

        //    //開啟word文件，fileName是路徑地址，需要副檔名
        //    //string fileName = @"D:\Desktop\Jennifer_test\YETI MAY FY23 SOW v0.4.docx";
        //    Aspose.Words.Document doc = new Document(fileName);

        //    //獲取word文件中的第3個表格
        //    Table docTable = (Table)doc.GetChildNodes(NodeType.Table, true)[2];
        //    //int c = doc.GetChildNodes(NodeType.Table, true).Count();

        //    for (int i = 0; i < docTable.Count; i++)
        //    {
        //        // IRow => cursor
        //        Row curRow = docTable.Rows[i];
        //        DataRow dr = dt.NewRow();

        //        for (int j = 0; j < curRow.Count; j++)
        //        {
        //            Cell cell = curRow.Cells[j];

        //            //用GetText()的方法來獲取cell中的值
        //            string cellValue = cell.GetText();

        //            if (isFirst)
        //            {
        //                columns = dt.Columns;

        //                if (columns.Contains(cellValue))
        //                    dt.Columns.Add(GetDistinctColumn(cellValue));
        //                else
        //                {
        //                    dt.Columns.Add(cellValue);
        //                    Console.WriteLine(cellValue);
        //                }
        //            }
        //            else if (j < dt.Columns.Count) //避免資料 cell 長度大於 column 的長度
        //                dr[j] = cellValue;
        //        }

        //        if (!isFirst)
        //            dt.Rows.Add(dr);

        //        isFirst = false;

        //    }
        //    return dt;

        //}
        public List<SowModel> WordSowDataTableToList(DataTable dt)
        {
            List<SowModel> sowModelList = new List<SowModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    SowModel sm = new SowModel
                    {
                        Category = row["Category"].ToString(),
                        DPN = row["DPN"].ToString(),
                        Description = row["Description of Change"].ToString(),

                    };
                    sowModelList.Add(sm);

                }
            }

            return sowModelList;
        }
        public string ListToExcel(List<SowModel> list, string saveFilePath, string formatPath, string sheetName, int copyFormatRow)
        {
            //根據指定的檔案格式建立對應的類
            using (FileStream templateFile = new FileStream(formatPath/*@"docs\format.xlsx"*/, FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.GetSheet(sheetName);/*"Add Part"*/
                //int copyFormatRow = 5;  //Head                          

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
                for (int i = 0; i < list.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;
                    SowModel sm = (SowModel)list[i];

                    //遍歷物件所有屬性
                    PropertyInfo[] propInfo = sm.GetType().GetProperties();
                    foreach (var p in propInfo)
                    {
                        // 放值
                        cell = row.CreateCell(cellIndex);
                        var val = p.GetValue(sm) == null ? "" : p.GetValue(sm).ToString();
                        cell.SetCellValue(val);
                        cell.CellStyle = cellStyle_black;

                        cellIndex++;
                    }
                }

                //寫檔
                string _saveFilePath = "".Equals(saveFilePath) ? "Inventory report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss")) + ".xlsx" : saveFilePath;
                SaveWorkbook(workbook, _saveFilePath);
            }


            return "OK";
        }
        public void DoInsertRow()
        {
            IWorkbook workbookTarget = OpenWorkbook(@"docs\polling_test.xlsx");
            ISheet destinationSheet = workbookTarget.GetSheet("Allocation polling-project");

            InsertRow(destinationSheet, 15);
            SaveWorkbook(workbookTarget, @"docs\polling_test.xlsx");
            Console.WriteLine("DONE!!!!!!!!!!!!!!!!!");
        }
        void InsertRow(ISheet sheet, int insertRow)
        {
            sheet.ShiftRows(insertRow, sheet.LastRowNum, 1);
            // 如果要插入的列數是0，複制下方列，反之，複制上方列(style)
            sheet.CopyRow(insertRow + 1, insertRow);
            // 清空插入列的值
            var row = sheet.GetRow(insertRow);
            for (int i = 0; i < row.LastCellNum; i++)
            {
                var cell = row.GetCell(i);
                if (cell != null)
                    cell.SetCellValue("");
            }
        }
        public void DoCopyRange()
        {
            IWorkbook workbook = OpenWorkbook(@"docs\Sample_polling_format.xlsx");
            IWorkbook workbookTarget = OpenWorkbook(@"docs\polling_test.xlsx");

            ISheet sourceSheet = workbook.GetSheet("Allocation polling-project");
            ISheet destinationSheet = workbookTarget.GetSheet("Allocation polling-project");

            int shiftNum = destinationSheet.GetRow(0).LastCellNum - 4;

            CopyRange(CellRangeAddress.ValueOf("E1:AT3"), sourceSheet, destinationSheet, shiftNum, workbookTarget);

            AddCellMergeRegion(destinationSheet, shiftNum);

            SaveWorkbook(workbookTarget, @"docs\polling_test.xlsx");
            Console.WriteLine("DONE!!!!!!!!!!!!!!!!!");
        }
        void CopyRange(CellRangeAddress range, ISheet sourceSheet, ISheet destinationSheet, int shiftNum, IWorkbook workbookTarget)
        {
            for (var rowNum = range.FirstRow; rowNum <= range.LastRow; rowNum++)
            {
                IRow sourceRow = sourceSheet.GetRow(rowNum);

                if (destinationSheet.GetRow(rowNum) == null)
                    destinationSheet.CreateRow(rowNum);

                if (sourceRow != null)
                {
                    IRow destinationRow = destinationSheet.GetRow(rowNum);

                    for (var col = range.FirstColumn; col < sourceRow.LastCellNum && col <= range.LastColumn; col++)
                    {
                        destinationRow.CreateCell(col + shiftNum);
                        CopyCell(sourceRow.GetCell(col), destinationRow.GetCell(col + shiftNum), workbookTarget);
                    }
                }
            }
            //copy cell width
            for (int i = 0; i < sourceSheet.GetRow(0).LastCellNum; i++)
            {
                destinationSheet.SetColumnWidth(i + shiftNum, sourceSheet.GetColumnWidth(i));
            }
        }
        void AddCellMergeRegion(ISheet destinationSheet, int shiftNum)
        {
            destinationSheet.AddMergedRegion(new CellRangeAddress(0, 2, 4 + shiftNum, 4 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(0, 2, 5 + shiftNum, 5 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(0, 0, 6 + shiftNum, 45 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 6 + shiftNum, 8 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 10 + shiftNum, 12 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 13 + shiftNum, 14 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 15 + shiftNum, 17 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 18 + shiftNum, 21 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 22 + shiftNum, 26 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 29 + shiftNum, 29 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 30 + shiftNum, 30 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 31 + shiftNum, 31 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 32 + shiftNum, 32 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 33 + shiftNum, 33 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 34 + shiftNum, 34 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 35 + shiftNum, 37 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 38 + shiftNum, 38 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 39 + shiftNum, 39 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 40 + shiftNum, 41 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 1, 42 + shiftNum, 44 + shiftNum));
            destinationSheet.AddMergedRegion(new CellRangeAddress(1, 2, 45 + shiftNum, 45 + shiftNum));
        }
        void CopyCell(ICell source, ICell destination, IWorkbook workbook)
        {
            if (destination != null && source != null)
            {
                //you can comment these out if you don't want to copy the style ...
                destination.CellComment = source.CellComment;
                destination.Hyperlink = source.Hyperlink;

                //destination.CellStyle = source.CellStyle;
                //注意新建單元格樣式的時候一定要克隆一下，否則單元格之前的樣式就沒了
                XSSFCellStyle cellStyle_cust = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_cust.CloneStyleFrom(source.CellStyle);
                destination.CellStyle = cellStyle_cust;

                switch (source.CellType)
                {
                    case CellType.Formula:
                        destination.CellFormula = source.CellFormula; break;
                    case CellType.Numeric:
                        destination.SetCellValue(source.NumericCellValue); break;
                    case CellType.String:
                        destination.SetCellValue(source.StringCellValue); break;
                }
            }
        }


    }
}

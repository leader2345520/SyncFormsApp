using Aspose.Words;
using Aspose.Words.Tables;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Table = Aspose.Words.Tables.Table;

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

        public DataTable WordSowToDataTable(string fileName)
        {
            DataTable dt = new DataTable();
            DataColumnCollection columns;
            bool isFirst = true;

            //開啟word文件，fileName是路徑地址，需要副檔名
            //string fileName = @"D:\Desktop\Jennifer_test\YETI MAY FY23 SOW v0.4.docx";
            Aspose.Words.Document doc = new Document(fileName);

            //獲取word文件中的第3個表格
            Table docTable = (Table)doc.GetChildNodes(NodeType.Table, true)[2];
            //int c = doc.GetChildNodes(NodeType.Table, true).Count();

            for (int i = 0; i < docTable.Count; i++)
            {
                // IRow => cursor
                Row curRow = docTable.Rows[i];
                DataRow dr = dt.NewRow();

                for (int j = 0; j < curRow.Count; j++)
                {
                    Cell cell = curRow.Cells[j];

                    //用GetText()的方法來獲取cell中的值
                    string cellValue = cell.GetText();

                    if (isFirst)
                    {
                        columns = dt.Columns;

                        if (columns.Contains(cellValue))
                            dt.Columns.Add(GetDistinctColumn(cellValue));
                        else
                        {
                            dt.Columns.Add(cellValue);
                            Console.WriteLine(cellValue);
                        }
                    }
                    else if (j < dt.Columns.Count) //避免資料 cell 長度大於 column 的長度
                        dr[j] = cellValue;
                }

                if (!isFirst)
                    dt.Rows.Add(dr);

                isFirst = false;

            }
            return dt;

        }
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
                using (FileStream file = new FileStream(_saveFilePath, FileMode.Create))//產生檔案             
                {
                    workbook.Write(file);
                }
            }


            return "OK";
        }
    }
}

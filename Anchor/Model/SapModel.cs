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

namespace Anchor
{
    class SapModel : BaseModel
    {
        #region getter/setter
        public string Supplier { get; set; }
        public string DPN { get; set; }
        public string Plnt { get; set; }
        public string StorLoc { get; set; }
        public string Sourcer { get; set; }
        public string ShortText { get; set; }
        public string Oun { get; set; }
        public string PriceUSD { get; set; }
        #endregion

        public List<SapModel> SapDataTableToList(DataTable dt)
        {
            List<SapModel> SapModelList = new List<SapModel>();
            if (dt != null)
            {
                foreach (DataRow row in dt.Rows)
                {
                    //PO 沒值才要放入表格
                    if (string.IsNullOrEmpty((string)row["PO"]))
                    {
                        SapModel sm = new SapModel
                        {
                            Supplier = row["Supplier"].ToString(),
                            DPN = row["DPN"].ToString(),
                            Plnt = "PIL6",
                            StorLoc = "LQ2N",
                            Sourcer = row["Sourcer"].ToString(),
                            ShortText = "", // row["Short text"].ToString(),
                            Oun = "", // row["Oun"].ToString(),
                            PriceUSD = row["Price-USD"].ToString()
                        };

                        SapModelList.Add(sm);
                    }
                }
            }

            return SapModelList;
        }
        public string ListToExcel(List<SapModel> sapList, string saveFilePath)
        {
            //根據指定的檔案格式建立對應的類
            using (FileStream templateFile = new FileStream(@"docs\SAP_format.xlsx", FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.GetSheetAt(0);
                int copyFormatRow = 1;  //Head 


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
                for (int i = 0; i < sapList.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;
                    SapModel sm = sapList[i];

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
    }
}

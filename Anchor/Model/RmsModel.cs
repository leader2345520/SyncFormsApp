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
    class RmsModel : BaseModel
    {
        public string Category { get; set; }

        public string DPN { get; set; }

        public string PartDescription { get; set; }

        public string Barcode { get; set; }

        public string VendorSN { get; set; }

        public string Project { get; set; }

        public string Borrower { get; set; }

        public string Status { get; set; }


        public List<RmsModel> RmsDataTableToModelList(DataTable dt)
        {
            List<RmsModel> rmsModelList = new List<RmsModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    RmsModel dm = new RmsModel()
                    {
                        DPN = row["客戶料號"].ToString(),
                        Barcode = row["條碼號"].ToString(),
                        VendorSN = row["來料SN"].ToString(),
                        Project = row["專案"].ToString(),
                        Borrower = row["借用人"].ToString(),
                        Status = row["物流狀態"].ToString(),
                        PartDescription = row["物料描述"].ToString(),

                        Category = GetCategory(row["物料大類"].ToString()),
                        //Category = row["物料大類"].ToString(),

                    };
                    rmsModelList.Add(dm);
                }
            }

            return rmsModelList;
        }

        public string GetCategory(string type1)
        {
            switch (type1)
            {
                case "CPU":
                    return "Processor";
                case "HDD":
                    return "Hard Drive";
                case "SSD":
                    return "Hard Drive";
                case "MEMORY":
                    return "DIMM";
                default:
                    return "";
            }
        }

        public string ColorListToExcel(List<RmsModel> modelList, string saveFilePath)
        {
            //根據指定的檔案格式建立對應的類
            using (FileStream templateFile = new FileStream(@"docs\Foxconn_external_audit_report.xlsx", FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new XSSFWorkbook(templateFile);
                ISheet sheet = workbook.GetSheet("Foxconn");
                int HeadRow = 1;  //Head                

                #region 藍色模板
                //設定儲存格 style(藍色)
                XSSFCellStyle cellStyle_blue = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_blue.BorderBottom = BorderStyle.Thin;
                cellStyle_blue.BorderLeft = BorderStyle.Thin;
                cellStyle_blue.BorderRight = BorderStyle.Thin;
                cellStyle_blue.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_blue = (XSSFFont)workbook.CreateFont();
                font_blue.FontName = "Calibri";//字體
                //font_blue.SetColor(new XSSFColor(new byte[] { 0, 0, 255 }));
                cellStyle_blue.SetFont(font_blue);
                #endregion

                // 把合併的 cell 置中
                XSSFCellStyle categoryStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                categoryStyle.Alignment = HorizontalAlignment.Center;
                categoryStyle.VerticalAlignment = VerticalAlignment.Center;
                categoryStyle.BorderBottom = BorderStyle.Thin;
                categoryStyle.BorderLeft = BorderStyle.Thin;
                categoryStyle.BorderRight = BorderStyle.Thin;
                categoryStyle.BorderTop = BorderStyle.Thin;
                categoryStyle.SetFont(font_blue);

                // 把合併的 cell 置中
                XSSFCellStyle dateStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                //dateStyle.Alignment = HorizontalAlignment.Center;
                dateStyle.VerticalAlignment = VerticalAlignment.Center;
                dateStyle.SetFont(font_blue);

                IRow row = null;
                ICell cell = null;

                string tempCategory = "Processor";
                int currentIndex = 1;
                int[] lastIndex = { 0, 0, 0, 0 };

                //寫入資料
                for (int i = 0; i < modelList.Count; i++)
                {
                    row = sheet.CreateRow(i + HeadRow);
                    int cellIndex = 0;
                    RmsModel rm = modelList[i];

                    //遍歷物件所有屬性
                    PropertyInfo[] propInfo = rm.GetType().GetProperties();
                    foreach (var p in propInfo)
                    {
                        if (p.Name.Equals("Color")) break;

                        // 放值
                        cell = row.CreateCell(cellIndex);
                        var val = p.GetValue(rm) == null ? "" : p.GetValue(rm).ToString();
                        cell.SetCellValue(val);
                        cell.CellStyle = cellStyle_blue;

                        cellIndex++;
                    }

                    // 最後一行固定 OK
                    cell = row.CreateCell(cellIndex);
                    cell.SetCellValue("OK");
                    cell.CellStyle = cellStyle_blue;

                    // 開始合併
                    if (!tempCategory.Equals(rm.Category) && currentIndex < i) {

                        // RowSpan.
                        var cra = new NPOI.SS.Util.CellRangeAddress(currentIndex, i, 0, 0);
                        sheet.AddMergedRegion(cra);
                        
                        // 合併完當下做 cell 設定.
                        ICell mergedCell = sheet.GetRow(currentIndex).GetCell(0);
                        mergedCell.CellStyle = categoryStyle;

                        currentIndex = i + 1;
                        tempCategory = rm.Category;
                    }
                }
                // 最後剩下的做合併
                if (currentIndex < modelList.Count) {
                    var cra2 = new NPOI.SS.Util.CellRangeAddress(currentIndex, modelList.Count, 0, 0);
                    sheet.AddMergedRegion(cra2);
                    ICell mergedCell2 = sheet.GetRow(currentIndex).GetCell(0);
                    mergedCell2.CellStyle = categoryStyle;
                }

                // 最後一列的 下一列的 I欄 ( index : 8 ) 輸出當月最後一日. format : Date:2023.08.30
                // 最後一列的下一列 在 index 表示 等於 modelList.Count + 1 ( IHeader )
                row = sheet.CreateRow(modelList.Count + 1);
                cell = row.CreateCell(8);
                int lastDay = DateTime.DaysInMonth(int.Parse(DateTime.Now.ToString("yyyy")), int.Parse(DateTime.Now.ToString("MM")));
                cell.SetCellValue("Date:" + DateTime.Now.ToString(("yyyy.MM")+ "." + lastDay));
                cell.CellStyle = dateStyle;
                

                //寫檔
                string _saveFilePath = "".Equals(saveFilePath) ? "Inventory report_" + DateTime.Now.ToString(("yyyyMMdd_HHmmss")) + ".xlsx" : saveFilePath;
                SaveWorkbook(workbook, _saveFilePath);
            }

            return "OK";
        }

        public static Random RandomGen = new Random();

        // 篩選
        // 1. 只要被借出的
        // 2. 只需要 25 條 ( 隨機 ) 
        // 3. 25 條須包含 3大 Category 並平均分布.
        public Dictionary<string, object> PorjectFilter(List<RmsModel> modelList)
        {
            try
            {
                // 只找借出狀態的
                modelList = modelList.FindAll(m => "借出".Equals(m.Status));
                // 只找符合 Processor / DIMM / Hard Drive.
                modelList = modelList.FindAll(m => ! "".Equals(m.Category));

                // Processor / DIMM / Hard Drive

                // sort by category ( 排序的時候 category 已經被分類過一輪了 ， 所以這裡直接跑 sort 沒問題. )
                // 但是 Module 要在 Memory 上 所以需要客製化 compare 方法.
                modelList.Sort((x, y) =>
                    CustomCompare(x.Category, y.Category)
                ); //升冪:a,b,c

                // 找到四大類的最後一筆 index.  ( 找到 Module 代表上一筆是 CPU ... 以此類推.
                int[] categoryCount = {0, 0, 0 };
                for (int i = 0; i < modelList.Count; i++)
                {
                    if (modelList[i].Category.Equals("Processor"))
                    {
                        categoryCount[0]++;
                    }
                    else if (modelList[i].Category.Equals("DIMM"))
                    {
                        categoryCount[1]++;
                    }
                    else if (modelList[i].Category.Equals("Hard Drive"))
                    {
                        categoryCount[2]++;
                    }
                }

                // max 8 8 9
                int[] pickList = { 0, 0, 0 };
                int need = 25;
                bool[] flag = { false, false, false }; 

                for (int i = 0; i < need; i++)
                {
                    // 無法滿足此條件就 break;
                    if (flag[0] && flag[1] && flag[2] ) {
                        break;
                    }

                    int randomValueBetween0And99 = RandomGen.Next(100);
                    if (randomValueBetween0And99 < 33)
                    {
                        if (pickList[0] + 1 <= categoryCount[0] && pickList[0] + 1 <= 8)
                        {
                            pickList[0] = pickList[0] + 1;
                        }
                        else
                        {
                            i--;
                            flag[0] = true;
                        }
                    }
                    else if (randomValueBetween0And99 < 67)
                    {
                        if (pickList[1] + 1 <= categoryCount[1] && pickList[1] + 1 <= 8 )
                        {
                            pickList[1] = pickList[1] + 1;
                        }
                        else
                        {
                            i--;
                            flag[1] = true;
                        }
                    }
                    else
                    {

                        if (pickList[2] + 1 <= categoryCount[2] && pickList[2] + 1 <= 9)
                        {
                            pickList[2] = pickList[2] + 1;
                        }
                        else
                        {
                            i--;
                            flag[2] = true;
                        }
                    }
                    
                }

                List<RmsModel> finalList = new List<RmsModel>();
                int alreadyUse = 0;
                for (int i = 0; i < pickList.Length; i++) {

                    int shouldPick = pickList[i];

                    HashSet<int> numbers = new HashSet<int>();
                    while (numbers.Count < shouldPick)
                    {
                        numbers.Add(RandomGen.Next(alreadyUse, alreadyUse + categoryCount[i]));
                    }
                    alreadyUse += categoryCount[i];

                    foreach (int item in numbers)
                    {
                        finalList.Add(modelList[item]);
                    }
                }

                return new Dictionary<string, object> { { "msg", "OK" }, { "returnObject", finalList } };
            }
            catch (Exception excption)
            {
                return new Dictionary<string, object> { { "msg", excption.Message }, { "returnObject", modelList } };
            }
        }

        // ref : https://learn.microsoft.com/zh-tw/dotnet/api/system.collections.generic.list-1.sort?view=net-7.0
        public int CustomCompare(string x, string y)
        {
            if (x == null)
            {
                if (y == null)
                {
                    // If x is null and y is null, they're
                    // equal.
                    return 0;
                }
                else
                {
                    // If x is null and y is not null, y
                    // is greater.
                    return -1;
                }
            }
            else
            {
                // If x is not null...
                //
                if (y == null)
                // ...and y is null, x is greater.
                {
                    return 1;
                }
                else
                {

                    //return x.CompareTo(y);

                    // 需要客製化的部分 ( 1. Processor / 2. DIMM / 3. Hard Drive )
                    if (x.Equals("Processor") && y.Equals("DIMM")) {
                        return -1;
                    } else if (x.Equals("Processor") && y.Equals("Hard Drive")) {
                        return -1;
                    }
                    else if (x.Equals("DIMM") && y.Equals("Processor"))
                    {
                        return 1;
                    }
                    else if (x.Equals("DIMM") && y.Equals("Hard Drive"))
                    {
                        return -1;
                    }
                    else if (x.Equals("Hard Drive") && y.Equals("DIMM"))
                    {
                        return 1;
                    }
                    else if (x.Equals("Hard Drive") && y.Equals("Processor"))
                    {
                        return 1;
                    }
                    else
                    {
                        return x.CompareTo(y);
                    }
                }
            }
        }
    }
}

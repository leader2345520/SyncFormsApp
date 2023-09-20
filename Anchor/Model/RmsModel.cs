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
                        DPN = row["客戶料號"].ToString()
                        .Replace("DELL:", "")
                        .Replace("DELL;", "")
                        .Replace("Dell:", "")
                        .Replace("Dell;", "")
                        .Trim(),
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
                        cell.SetCellType(CellType.String);

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
                modelList.Sort((x, y) => {
                    int result = x.Project.CompareTo(y.Project);
                    return result != 0 ? result : CustomCompare(x.Category, y.Category);
                }); //升冪:a,b,c

                // 找到各大類有幾筆 index.  ( 並依照所需排序 Processor > DIMM > Hard Drive )
                Dictionary<string, List<int>> projectAndCategoryMap = new Dictionary<string, List<int>>();
                for (int i = 0; i < modelList.Count; i++)
                {
                    if (! projectAndCategoryMap.ContainsKey(modelList[i].Project))
                    {
                        List<int> initList = new List<int>();
                        initList.Add(-1); // Processor index start in this Project
                        initList.Add(0); // Processor index end in this Project
                        initList.Add(-1); // DIMM index start in this Project
                        initList.Add(0); // DIMM index end in this Project
                        initList.Add(-1); // Hard Drive index start in this Project
                        initList.Add(0); // Hard Drive index end in this Project
                        projectAndCategoryMap.Add(modelList[i].Project, initList);
                    }
                    List<int> list = projectAndCategoryMap[modelList[i].Project];

                    if (modelList[i].Category.Equals("Processor"))
                    {
                        if (list[0] == -1) {
                            list[0] = i;
                        }
                        list[1]++;
                    }
                    else if (modelList[i].Category.Equals("DIMM"))
                    {
                        if (list[2] == -1)
                        {
                            list[2] = i;
                        }
                        list[3]++;
                    }
                    else if (modelList[i].Category.Equals("Hard Drive"))
                    {
                        if (list[4] == -1)
                        {
                            list[4] = i;
                        }
                        list[5]++;
                    }
                }

                // 計算每個 project 理論上要被抽幾筆.
                int need = 25;
                int averagePick = need / projectAndCategoryMap.Count;

                HashSet<int> numbers = new HashSet<int>();

                // 平均要抽的筆數，再平均分配到三大類
                foreach (KeyValuePair<string, List<int>> entry in projectAndCategoryMap)
                {
                    List<int> list = entry.Value;
                    if (list[0] == -1) list[0] = 0;
                    if (list[2] == -1) list[2] = 0;
                    if (list[4] == -1) list[4] = 0;
                    // do something with entry.Value or entry.Key
                    int[] pickNum = { 0, 0, 0 };
                    bool[] pickBoo = { false, false, false };

                    int everyCategory = averagePick / 3;

                    if (everyCategory > 0) {

                        while (pickNum[0] < everyCategory && list[0] + list[1] >= everyCategory) {
                            numbers.Add(RandomGen.Next(list[0], list[0] + list[1]));
                            pickNum[0]++;
                        }
                        while (pickNum[1] < everyCategory && list[2] + list[3] >= everyCategory)
                        {
                            numbers.Add(RandomGen.Next(list[2], list[2] + list[3]));
                            pickNum[1]++;
                        }
                        while (pickNum[2] < everyCategory && list[4] + list[5] >= everyCategory)
                        {
                            numbers.Add(RandomGen.Next(list[4], list[4] + list[5]));
                            pickNum[2]++;
                        }
                    }

                    int left = averagePick % 3;
                    while (left != 0 ) {

                        int randomCategory = RandomGen.Next(0, 3);

                        switch (randomCategory)
                        {
                            case 0:
                                if (list[0] + list[1] > pickNum[0])
                                {
                                    numbers.Add(RandomGen.Next(list[0], list[0] + list[1]));
                                    left--;
                                }
                                else {
                                    pickBoo[0] = true;
                                }
                                break;
                            case 1:
                                if (list[2] + list[3] > pickNum[1])
                                {
                                    numbers.Add(RandomGen.Next(list[2], list[2] + list[3]));
                                    left--;
                                }
                                else
                                {
                                    pickBoo[1] = true;
                                }
                                break;
                            case 2:
                                if (list[4] + list[5] > pickNum[2])
                                {
                                    numbers.Add(RandomGen.Next(list[4], list[4] + list[5]));
                                    left--;
                                }
                                else
                                {
                                    pickBoo[2] = true;
                                }
                                break;
                        }

                        if (pickBoo[0] && pickBoo[1] && pickBoo[2])
                        {
                            break;
                        }
                    }
                }
                // 如果能25個，就補到 25 個
                while (numbers.Count < 25 && modelList.Count >= 25) {
                    numbers.Add(RandomGen.Next(0, modelList.Count));
                }

                List<RmsModel> finalList = new List<RmsModel>();
                foreach (int item in numbers)
                {
                    finalList.Add(modelList[item]);
                }

                finalList.Sort((x, y) =>
                    CustomCompare(x.Category, y.Category)
                ); //升冪:a,b,c


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

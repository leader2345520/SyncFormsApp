using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static NPOI.HSSF.Util.HSSFColor;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;

namespace Anchor
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

        internal static string GetValueFromCell(ICell icell)
        {
            string output = "";
            if (icell != null)
            {
                switch (icell.CellType)
                {
                    case CellType.String:
                        output = icell.StringCellValue.Trim();
                        break;
                    case CellType.Numeric:
                        output = icell.NumericCellValue.ToString();
                        break;
                    case CellType.Boolean:
                        output = icell.BooleanCellValue.ToString();
                        break;
                    default:
                        output = "";
                        break;
                }
            }
            return output;
            throw new NotImplementedException();
        }
    }

    //BaseModel
    class BaseModel { }

    //RMS To Dwll
    class DellModel : BaseModel
    {
        #region getter/setter
        public string Location_s { get; set; }
        public string SerialNumber_s { get; set; }
        public string Part_s { get; set; }
        public string Project_s { get; set; }
        public string PartLocation_s { get; set; }
        public string UserDefinedDescription_s { get; set; }
        public string Manufacturer_s { get; set; }
        public string TestStatus { get; set; }
        public string AgilePartType { get; set; }
        public string Revision_s { get; set; }
        public string Category_s { get; set; }
        public string Asset { get; set; }
        public string FrozenCost_s { get; set; }
        public string UserDefinedOriginalCost_s { get; set; }
        public string CountryofOrigin_s { get; set; }
        public string ManufacturerPartNumber { get; set; }
        public string AllocatedBorrower { get; set; }
        public string CurrentOwner_s { get; set; }
        public string PurchasingCostCenter_s { get; set; }
        public string ShippingCarrierInformation { get; set; }
        public string BondingInformation { get; set; }
        public string PONumber_s { get; set; }
        public string PartComment { get; set; }
        public string RecallDate { get; set; }
        public string TDCHWL { get; set; }
        public string Comment { get; set; }

        public byte[] Color { get; set; }
        #endregion

        public DataTable CsvRmsToDataTable(string filename, string seperator)
        {
            DataTable dt = new DataTable();
            bool isFirst = true;
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sr = new StreamReader(fs, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        var eachLineStr = sr.ReadLine();

                        string[] rows = eachLineStr.Split(new string[] { seperator }, StringSplitOptions.None);

                        char[] ch = eachLineStr.ToCharArray();
                        for (int i = 0; i < ch.Length; i++)
                        {
                            if (i < ch.Length - 1 && ch[i] == ch[i + 1] && ch[i] == ',')
                            {
                                ch[i] = '換';
                            }

                        }

                        string charsStr = new string(ch).Replace("換", ", ");

                        var result = new List<string>();
                        var filter = @"([^\""\,]*[^\""\,])|[\""]([^\""]*)[\""]";
                        Match match = Regex.Match(charsStr, filter, RegexOptions.IgnoreCase);

                        while (match.Success)
                        {
                            if (!string.IsNullOrEmpty(match.Groups[2].Value))
                            {
                                result.Add(match.Groups[2].Value);
                            }
                            else
                            {
                                if (result.Contains(match.Value) && isFirst)
                                    result.Add(match.Value + "1");
                                else
                                    result.Add(match.Value);
                            }
                            match = match.NextMatch();
                        }

                        //exclude first line
                        if (isFirst)
                        {
                            for (int i = 0; i < result.Count; i++)
                            {
                                dt.Columns.Add(result[i]);
                            }
                            isFirst = false;
                            continue;
                        }
                        //split string to array of string use seperator
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < result.Count; i++)
                        {
                            dr[i] = result[i];
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }

            return dt;
        }
        public List<DellModel> CsvRmsDataTableToList(DataTable dt)
        {

            List<DellModel> DellModelList = new List<DellModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    DellModel dm = new DellModel()
                    {
                        //Location_s = row["*Location"].ToString(),
                        SerialNumber_s = row["Manuf. Shipment S/N"].ToString(),
                        Part_s = string.IsNullOrEmpty(row["Cust. PN"].ToString()) || row["Cust. PN"].ToString().Split(':').Length < 2 ? row["Cust. PN"].ToString() : row["Cust. PN"].ToString().Split(':')[1].Trim(),
                        Project_s = row["Purchase Pur."].ToString(),
                        PartLocation_s = GetPartLocation(row["Status"].ToString(), row["Fixed Location"].ToString()),
                        UserDefinedDescription_s = row["Description"].ToString(),
                        Manufacturer_s = row["Manufacturer"].ToString(),
                        //TestStatus = row["Test Status"].ToString(),
                        //AgilePartType = row["Agile Part Type"].ToString(),
                        Revision_s = row["HW/SW Version"].ToString(),
                        Category_s = GetCategory(row["type1"].ToString(), row["type2"].ToString()),
                        //Asset = row["Asset #"].ToString(),
                        //FrozenCost_s = row["*Frozen Cost"].ToString(),
                        UserDefinedOriginalCost_s = row["Price"].ToString(),
                        //CountryofOrigin_s = row["*Country of Origin"].ToString(),
                        //ManufacturerPartNumber = row["Manufacturer Part Number"].ToString(),
                        //AllocatedBorrower = row["Allocated Borrower"].ToString(),
                        //CurrentOwner_s = row["*Current Owner"].ToString(),
                        //PurchasingCostCenter_s = row["*Purchasing Cost Center"].ToString(),
                        //ShippingCarrierInformation = row["Shipping / Carrier Information"].ToString(),
                        //BondingInformation = row["Bonding Information"].ToString(),
                        //PONumber_s = row["*PO Number"].ToString(),
                        //PartComment = row["Part Comment"].ToString(),
                        //RecallDate = row["Recall Date"].ToString(),
                        //TDCHWL = row["TDC HWL"].ToString(),
                        Comment = GetComment(row["Status"].ToString(), "TPE"),
                        Color = new byte[] { 0, 0, 255 }  //New Data 預設字體顏色藍色
                    };
                    DellModelList.Add(dm);
                }


            }

            return DellModelList;
        }
        public List<DellModel> XlsRmsDataTableToList(DataTable dt)
        {

            List<DellModel> DellModelList = new List<DellModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    DellModel dm = new DellModel()
                    {//Location_s = row["*Location"].ToString(),
                        SerialNumber_s = row["Manuf. Shipment S/N"].ToString(),
                        Part_s = row["Customer Part Number"].ToString(),
                        Project_s = row["Purchase Purpose"].ToString(),
                        PartLocation_s = "TDC->FOXCONN TJ->",//GetPartLocation(row["Status"].ToString(), row["Fixed Location"].ToString()),
                        UserDefinedDescription_s = row["Other Description"].ToString(),
                        Manufacturer_s = row["Manufacturer"].ToString(),
                        //TestStatus = row["Test Status"].ToString(),
                        //AgilePartType = row["Agile Part Type"].ToString(),
                        Revision_s = row["HW/SW Version"].ToString(),
                        Category_s = GetCategory(row["type 1"].ToString(), row["type 2"].ToString()),
                        //Asset = row["Asset #"].ToString(),
                        //FrozenCost_s = row["*Frozen Cost"].ToString(),
                        UserDefinedOriginalCost_s = row["Price"].ToString(),
                        //CountryofOrigin_s = row["*Country of Origin"].ToString(),
                        //ManufacturerPartNumber = row["Manufacturer Part Number"].ToString(),
                        //AllocatedBorrower = row["Allocated Borrower"].ToString(),
                        //CurrentOwner_s = row["*Current Owner"].ToString(),
                        //PurchasingCostCenter_s = row["*Purchasing Cost Center"].ToString(),
                        //ShippingCarrierInformation = row["Shipping / Carrier Information"].ToString(),
                        //BondingInformation = row["Bonding Information"].ToString(),
                        //PONumber_s = row["*PO Number"].ToString(),
                        //PartComment = row["Part Comment"].ToString(),
                        //RecallDate = row["Recall Date"].ToString(),
                        //TDCHWL = row["TDC HWL"].ToString(),
                        Comment = GetComment(row["Status"].ToString(), "TJ"),
                        Color = new byte[] { 0, 0, 255 }  //New Data 預設字體顏色藍色

                    };
                    DellModelList.Add(dm);
                }


            }

            return DellModelList;
        }
        public List<DellModel> RMSListInputColor(List<DellModel> list)
        {
            foreach (DellModel dm in list)
            {
                if (String.IsNullOrEmpty(dm.SerialNumber_s.Trim()))
                {
                    dm.Color = new byte[] { 0, 128, 0 }; //暗綠色
                    dm.SerialNumber_s = "NA";
                }
                else if (String.IsNullOrEmpty(dm.Part_s.Trim()) || String.IsNullOrEmpty(dm.Project_s.Trim()) ||
                         String.IsNullOrEmpty(dm.PartLocation_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedDescription_s.Trim()) ||
                         String.IsNullOrEmpty(dm.Manufacturer_s.Trim()) || String.IsNullOrEmpty(dm.Revision_s.Trim()) ||
                         String.IsNullOrEmpty(dm.Category_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedOriginalCost_s.Trim()))

                    dm.Color = new byte[] { 255, 0, 0 };

            }

            return list;
        }




        public DataTable ExcelToDataTable(string filePath, object sheetPage, int copyFormatRow)
        {
            DataTable dt = new DataTable();
            bool isFirst = true;


            //IWorkbook => excel file
            IWorkbook workbook = null;

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {

                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);

                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                ISheet sheet = null;

                //ISheet => sheet
                if (sheetPage is int)
                {
                    sheet = workbook.GetSheetAt((int)sheetPage);
                    copyFormatRow = 1;
                }
                else sheet = workbook.GetSheet((string)sheetPage);

                if (sheet != null)
                {
                    //LastRowNum => total count
                    int rowCount = sheet.LastRowNum;
                    for (int i = copyFormatRow - 1; i < rowCount + 1; i++)
                    {
                        // IRow => cursor
                        IRow curRow = sheet.GetRow(i);
                        DataRow dr = dt.NewRow();

                        for (int j = 0; j < curRow.LastCellNum; j++)
                        {

                            ICell icell = curRow.GetCell(j);
                            string cellValue = Program.GetValueFromCell(icell);


                            if (isFirst)
                                dt.Columns.Add(cellValue);
                            else
                                dr[j] = cellValue;

                        }

                        if (!isFirst)
                            dt.Rows.Add(dr);

                        isFirst = false;

                    }
                }
            }

            return dt;

        }
        public List<DellModel> DellDataTableToList(DataTable dt)
        {
            List<DellModel> DellModelList = new List<DellModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    DellModel dm = new DellModel
                    {
                        Location_s = row["*Location"].ToString(),
                        SerialNumber_s = row["*Serial Number"].ToString(),
                        Part_s = row["*Part #"].ToString(),
                        Project_s = row["*Project"].ToString(),
                        PartLocation_s = row["*Part Location"].ToString(),
                        UserDefinedDescription_s = row["*User Defined Description"].ToString(),
                        Manufacturer_s = row["*Manufacturer"].ToString(),
                        TestStatus = row["Test Status"].ToString(),
                        AgilePartType = row["Agile Part Type"].ToString(),
                        Revision_s = row["*Revision"].ToString(),
                        Category_s = row["*Category"].ToString(),
                        Asset = row["Asset #"].ToString(),
                        FrozenCost_s = row["*Frozen Cost"].ToString(),
                        UserDefinedOriginalCost_s = row["*User Defined Original Cost"].ToString(),
                        CountryofOrigin_s = row["*Country of Origin"].ToString(),
                        ManufacturerPartNumber = row["Manufacturer Part Number"].ToString(),
                        AllocatedBorrower = row["Allocated Borrower"].ToString(),
                        CurrentOwner_s = row["*Current Owner"].ToString(),
                        PurchasingCostCenter_s = row["*Purchasing Cost Center"].ToString(),
                        ShippingCarrierInformation = row["Shipping / Carrier Information"].ToString(),
                        BondingInformation = row["Bonding Information"].ToString(),
                        PONumber_s = row["*PO Number"].ToString(),
                        PartComment = row["Part Comment"].ToString(),
                        RecallDate = row["Recall Date"].ToString(),
                        TDCHWL = row["TDC HWL"].ToString(),
                        Comment = row["Comment"].ToString(),
                        Color = new byte[] { 0, 0, 0 }  //New Data 預設字體顏色黑色
                    };
                    DellModelList.Add(dm);

                }
            }

            return DellModelList;
        }
        public List<DellModel> DellListInputColor(List<DellModel> list)
        {
            {
                foreach (DellModel dm in list)
                {
                    if (String.IsNullOrEmpty(dm.SerialNumber_s.Trim()))
                    {
                        dm.Color = new byte[] { 0, 255, 0 };  //亮綠色
                        dm.SerialNumber_s = "NA";
                    }
                    else if (String.IsNullOrEmpty(dm.Part_s.Trim()) || String.IsNullOrEmpty(dm.Project_s.Trim()) ||
                             String.IsNullOrEmpty(dm.PartLocation_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedDescription_s.Trim()) ||
                             String.IsNullOrEmpty(dm.Manufacturer_s.Trim()) || String.IsNullOrEmpty(dm.Revision_s.Trim()) ||
                             String.IsNullOrEmpty(dm.Category_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedOriginalCost_s.Trim()))

                        dm.Color = new byte[] { 255, 0, 255 };  //粉紅色
                }

                return list;
            }
        }


        public void StautsFilter(List<DellModel> rmsColorList)
        {
            for (int i = 0; i < rmsColorList.Count; i++)
            {
                if (("FilterOut").Contains(rmsColorList[i].Comment))
                {
                    rmsColorList.RemoveAt(i);
                    i--; //移除後物件為往上移動, 需 i-- 留在原位再判斷一次
                }
            }
        }
        public Dictionary<string, object> PorjectFilter(List<DellModel> rmsColorList, string txtFilter)
        {
            try
            {
                if (!"None".Equals(txtFilter) && !string.IsNullOrEmpty(txtFilter.Trim()))
                {
                    string[] project = txtFilter.Trim().Split(',');
                    List<DellModel> finalList = new List<DellModel>();
                    foreach (string p in project)
                    {
                        for (int i = 0; i < rmsColorList.Count; i++)
                        {
                            if (p.Equals(rmsColorList[i].Project_s))
                                finalList.Add(rmsColorList[i]);
                        }
                    }
                    return new Dictionary<string, object> { { "msg", "OK" }, { "returnObject", finalList } };
                }
                else
                    return new Dictionary<string, object> { { "msg", "OK" }, { "returnObject", rmsColorList } };
            }
            catch (Exception excption)
            {
                return new Dictionary<string, object> { { "msg", excption.Message }, { "returnObject", rmsColorList } };
            }
        }
        public string ColorListToExcel(List<DellModel> colorList, string saveFilePath)
        {
            //根據指定的檔案格式建立對應的類
            string basePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin");
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

                #region 紅色模板
                //設定處存格 style(粉紅色)
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

                #region 暗綠色模板
                //設定處存格 style(暗綠色)
                XSSFCellStyle cellStyle_darkGreen = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_darkGreen.BorderBottom = BorderStyle.Thin;
                cellStyle_darkGreen.BorderLeft = BorderStyle.Thin;
                cellStyle_darkGreen.BorderRight = BorderStyle.Thin;
                cellStyle_darkGreen.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_darkGreen = (XSSFFont)workbook.CreateFont();
                font_darkGreen.FontName = "Calibri";//字體
                font_darkGreen.SetColor(new XSSFColor(new byte[] { 0, 128, 0 }));
                cellStyle_darkGreen.SetFont(font_darkGreen);
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

                #region 粉紅色模板
                //設定處存格 style(粉紅色)
                XSSFCellStyle cellStyle_pink = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_pink.BorderBottom = BorderStyle.Thin;
                cellStyle_pink.BorderLeft = BorderStyle.Thin;
                cellStyle_pink.BorderRight = BorderStyle.Thin;
                cellStyle_pink.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_pink = (XSSFFont)workbook.CreateFont();
                font_pink.FontName = "Calibri";//字體
                font_pink.SetColor(new XSSFColor(new byte[] { 255, 0, 255 }));
                cellStyle_pink.SetFont(font_pink);
                #endregion

                #region 亮綠色模板
                //設定處存格 style(亮綠色)
                XSSFCellStyle cellStyle_lightGreen = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_lightGreen.BorderBottom = BorderStyle.Thin;
                cellStyle_lightGreen.BorderLeft = BorderStyle.Thin;
                cellStyle_lightGreen.BorderRight = BorderStyle.Thin;
                cellStyle_lightGreen.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_lightGreen = (XSSFFont)workbook.CreateFont();
                font_lightGreen.FontName = "Calibri";//字體
                font_lightGreen.SetColor(new XSSFColor(new byte[] { 0, 255, 0 }));
                cellStyle_lightGreen.SetFont(font_lightGreen);
                #endregion

                IRow row = null;
                ICell cell = null;

                //寫入資料
                for (int i = 0; i < colorList.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;
                    DellModel dm = colorList[i];

                    //遍歷物件所有屬性
                    PropertyInfo[] propInfo = dm.GetType().GetProperties();
                    foreach (var p in propInfo)
                    {
                        if (p.Name.Equals("Color")) break;
                        // 放值
                        cell = row.CreateCell(cellIndex);
                        var val = p.GetValue(dm) == null ? "" : p.GetValue(dm).ToString();
                        cell.SetCellValue(val);

                        if (BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 0, 0, 0 }))
                            cell.CellStyle = cellStyle_black;
                        else if ((BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 0, 0, 255 })))
                            cell.CellStyle = cellStyle_blue;
                        else if ((BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 255, 0, 255 })))
                            cell.CellStyle = cellStyle_pink;
                        else if ((BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 255, 0, 0 })))
                            cell.CellStyle = cellStyle_red;
                        else if ((BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 0, 255, 0 })))
                            cell.CellStyle = cellStyle_lightGreen;
                        else if ((BitConverter.ToString(dm.Color) == BitConverter.ToString(new byte[] { 0, 128, 0 })))
                            cell.CellStyle = cellStyle_darkGreen;

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

        #region common function
        public string GetCategory(string type1, string type2)
        {
            switch (type1)
            {
                case "CPU":
                    return "PROCESSOR";
                case "MEMORY":
                    return "DIMM";
                case "HDD":
                    return "HARD DRIVE";
                case "Assembly":
                    return "ASSY,PWA";
                case "Add on card":
                    if (new string[] { "LAN CARD", "Control Board", "TPM", "Boss module", "Other" }.Contains(type2))
                        return "ASSY,PWA";
                    else if (new string[] { "Perc Card", "Storage Card", "RAID CARD" }.Contains(type2))
                        return "PWA,CONTROLLER";
                    else if ("IDSDM".Equals(type2))
                        return "CARD,MEMORY RISER";
                    else if ("Fiber Channel".Equals(type2))
                        return "ASSY,CARD,INTERPOSER";
                    else return "";
                case "Other":
                    if ("SD Card".Equals(type2))
                        return "HARD DRIVE";
                    else
                        return "";
                default:
                    return "";
            }

        }
        public string GetPartLocation(string status, string fixedLocation)
        {
            if (new string[] { "available", "borrowed" }.Contains(status))
                return "TDC->FOXCONN TPE->";
            else if ("returned".Equals(status))
            {
                if (fixedLocation.Contains("轉天津"))
                    return "TDC->FOXCONN TJ->";
                else
                    return "";
            }
            else
                return "";
        }
        public string GetComment(string status, string location)
        {
            status = status.ToLower();
            switch (location)
            {
                case "TJ":
                    if ("available".Equals(status.ToLower()))
                        return "Idle";
                    else if ("borrowed".Equals(status.ToLower()))
                        return "In Use";
                    else if (new string[] { "callback", "returned" }.Contains(status))
                        return "Check Out";
                    else if (new string[] { "sold", "broken" }.Contains(status))
                        return "FilterOut";
                    else
                        return "";

                case "TPE":
                    if (new string[] { "available", "returned" }.Contains(status))
                        return "Idle";
                    else if ("borrowed".Equals(status))
                        return "In Use";
                    else if ("callback".Equals(status))
                        return "Check Out";
                    else if (new string[] { "sold", "broken" }.Contains(status))
                        return "FilterOut";
                    else
                        return "";

                default:
                    return "";
            }
        }

        #endregion

    }

    //Dell Merge
    class DellCellModel : BaseModel
    {
        public string Value { get; set; } = "";
        public byte[] FontColor { get; set; } = { 0, 0, 0 };


        public List<List<DellCellModel>> CopySheetsToRowList(string filePath)
        {

            DellCellModel dmc;
            List<DellCellModel> cellList;
            List<List<DellCellModel>> rowList = new List<List<DellCellModel>>();
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

                            cellList = new List<DellCellModel>();

                            for (int j = 0; j < curRow.LastCellNum; j++)
                            {
                                // ICell => column
                                ICell icell = curRow.GetCell(j);
                                dmc = new DellCellModel();
                                if (icell == null)
                                {
                                    cellList.Add(dmc);
                                    continue;
                                }

                                icell.SetCellType(CellType.String);
                                var cellValue = Program.GetValueFromCell(icell);

                                // IFont => colum style
                                IFont ifont = icell.CellStyle.GetFont(workbook);
                                if ((ifont as XSSFFont).GetXSSFColor() == null)
                                {
                                    cellList.Add(dmc);
                                    continue;
                                }
                                byte[] fontColor = (ifont as XSSFFont).GetXSSFColor().RGB;



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
        public string RowListToExcel(List<List<DellCellModel>> rowList, string saveFilePath)
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

                IRow row = null;
                ICell cell = null;

                //寫入資料
                for (int i = 0; i < rowList.Count; i++)
                {
                    row = sheet.CreateRow(i + copyFormatRow);
                    int cellIndex = 0;

                    foreach (DellCellModel dmc in rowList[i])
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
                using (FileStream file = new FileStream(_saveFilePath, FileMode.Create))//產生檔案
                {
                    workbook.Write(file);
                }
            }


            return "OK";
        }
        public bool IsRowEmpty(IRow row)
        {
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

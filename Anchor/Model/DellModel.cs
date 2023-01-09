using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Anchor
{
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
        public string Barcode { get; set; }

        public byte[] Color { get; set; }
        public string InputtingTime { get; set; }
        #endregion

        public List<DellModel> NewRmsDataTableToList(DataTable dt)
        {

            List<DellModel> dellModelList = new List<DellModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    DellModel dm = new DellModel()
                    {
                        //Location_s = row["*Location"].ToString(),
                        SerialNumber_s = row["來料SN"].ToString(),
                        Part_s = string.IsNullOrEmpty(row["客戶料號"].ToString()) || row["客戶料號"].ToString().Split(':').Length < 2 ? row["客戶料號"].ToString() : row["客戶料號"].ToString().Split(':')[1].Trim(),
                        Project_s = row["專案"].ToString(),
                        PartLocation_s = GetPartLocation(row["物流狀態"].ToString(), row["條碼號"].ToString()),
                        UserDefinedDescription_s = row["物料描述"].ToString(),
                        Manufacturer_s = row["製造商"].ToString(),
                        //TestStatus = row["Test Status"].ToString(),
                        //AgilePartType = row["Agile Part Type"].ToString(),
                        Revision_s = row["製造商版次"].ToString(),
                        Category_s = GetCategory(row["物料大類"].ToString()),
                        //Asset = row["Asset #"].ToString(),
                        //FrozenCost_s = row["*Frozen Cost"].ToString(),
                        UserDefinedOriginalCost_s = row["價格"].ToString(),
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
                        Comment = GetComment(row["物流狀態"].ToString()),
                        Barcode = row["條碼號"].ToString(),

                        Color = new byte[] { 0, 0, 255 }  //New Data 預設字體顏色藍色
                    };
                    dellModelList.Add(dm);
                }


            }

            return dellModelList;
        }
        public List<DellModel> NewRmsToDataTable(DataTable dt)
        {

            List<DellModel> dellModelList = new List<DellModel>();
            if (dt != null)
            {

                foreach (DataRow row in dt.Rows)
                {
                    DellModel dm = new DellModel()
                    {//Location_s = row["*Location"].ToString(),
                        SerialNumber_s = row["Manuf. S/N"].ToString(),
                        Part_s = row["Customer Part Number"].ToString(),
                        Project_s = row["Purchase Purpose"].ToString(),
                        PartLocation_s = "TDC->FOXCONN TJ->",//GetPartLocation(row["Status"].ToString(), row["Fixed Location"].ToString()),
                        UserDefinedDescription_s = row["Other Description"].ToString(),
                        Manufacturer_s = row["Manufacturer"].ToString(),
                        //TestStatus = row["Test Status"].ToString(),
                        //AgilePartType = row["Agile Part Type"].ToString(),
                        Revision_s = GetRevision(row["Other Description"].ToString()),
                        Category_s = GetCategory(row["type 1"].ToString()),
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
                        Comment = GetComment(row["Status"].ToString()),
                        Barcode = row["Barcode"].ToString(),

                        Color = new byte[] { 0, 0, 255 },  //New Data 預設字體顏色藍色
                        InputtingTime = row["Inputting Time"].ToString(), // TJ 要存輸入時間
                    };
                    dellModelList.Add(dm);
                }


            }

            return dellModelList;
        }
        public List<DellModel> RMSListInputColor(List<DellModel> list)
        {
            foreach (DellModel dm in list)
            {
                if (String.IsNullOrEmpty(dm.SerialNumber_s.Trim()))
                {
                    dm.Color = new byte[] { 0, 128, 0 }; //暗綠色                   
                }
                else if (String.IsNullOrEmpty(dm.Part_s.Trim()) || String.IsNullOrEmpty(dm.Project_s.Trim()) ||
                         String.IsNullOrEmpty(dm.PartLocation_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedDescription_s.Trim()) ||
                         String.IsNullOrEmpty(dm.Manufacturer_s.Trim()) || String.IsNullOrEmpty(dm.Revision_s.Trim()) ||
                         String.IsNullOrEmpty(dm.Category_s.Trim()) || String.IsNullOrEmpty(dm.UserDefinedOriginalCost_s.Trim()))

                    dm.Color = new byte[] { 255, 0, 0 };

            }

            return list;
        }


        public List<DellModel> DellDataTableToList(DataTable dt)
        {
            List<DellModel> dellModelList = new List<DellModel>();
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
                        Barcode = row["Barcode"].ToString(),
                        Color = new byte[] { 0, 0, 0 }  //New Data 預設字體顏色黑色
                    };
                    dellModelList.Add(dm);

                }
            }

            return dellModelList;
        }
        public List<DellModel> DellListInputColor(List<DellModel> list)
        {
            {
                foreach (DellModel dm in list)
                {
                    if (String.IsNullOrEmpty(dm.SerialNumber_s.Trim()))
                    {
                        dm.Color = new byte[] { 0, 255, 0 };  //亮綠色                        
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
                SaveWorkbook(workbook, _saveFilePath);
            }


            return "OK";
        }

        #region common function
        public string GetCategory(string type1)
        {
            switch (type1)
            {
                case "CPU":
                    return "PROCESSOR";
                case "GPU":
                    return "GRAPHICS CARD";
                case "HDD":
                    return "HARD DRIVE";
                case "SSD":
                    return "HARD DRIVE";
                case "MEMORY":
                    return "DIMM";
                case "HBA_RAID":
                    return "ASSY,CARD,INTERPOSER";
                case "NIC CARD":
                    return "ASSY,CARD,INTERPOSER";
                case "PSU":
                    return "POWER SUPPLY";
                case "Fixed_Asset":
                    return "SYSTEM - SYSTEM";
                case "Server":
                    return "SYSTEM - SYSTEM";
                default:
                    return "";
            }

        }
        public string GetPartLocation(string status, string barcode)
        {
            //TD,SN --> TPE
            //TJ --> TJ
            // status 如果是 return location 就互換
            if (new string[] { "良品入库", "借出" }.Contains(status))
                if (barcode.Contains("TD") || barcode.Contains("SN"))
                    return "TDC->FOXCONN TPE->";
                else if (barcode.Contains("TJ"))
                    return "TDC->FOXCONN TJ->";
                else return "";
            else if ("轉出".Equals(status))
            {
                if (barcode.Contains("TD") || barcode.Contains("SN"))
                    return "TDC->FOXCONN TJ->";
                else if (barcode.Contains("TJ"))
                    return "TDC->FOXCONN TPE->";
                else return "";
            }
            else
                return "";
        }
        public string GetComment(string status)
        {

            if ("良品入库".Equals(status))
                return "Idle";
            else if ("借出".Equals(status))
                return "In Use";
            else if ("Call Back".Equals(status))
                return "Check Out";
            else if (new string[] { "待報廢", "遺失" }.Contains(status))
                return "FilterOut";
            else if ("轉出".Equals(status))
                //if (fixedLocation.Contains("天津"))
                return "In Use";
            //else if (fixedLocation.Contains("鴻佰"))
            //    return "FilterOut";
            //else
            //    return "Check Out";
            else
                return "";

        }

        public string GetRevision(string description)
        {
            string result = "";
            if (description != null)
            {
                string[] tokens = description.Split('/');

                foreach (string str in tokens)
                {
                    if (str.ToUpper().Contains("REV:"))
                    {
                        result = str.ToUpper().Replace("REV:", "").Trim();
                        break;
                    }
                }
            }
            return result;
        }
        public Dictionary<string, DellModel> GetDellColorListDict(List<DellModel> dellColorList)
        {
            string barcode;
            Dictionary<string, DellModel> dict = new Dictionary<string, DellModel>();
            foreach (DellModel dm in dellColorList)
            {
                barcode = string.IsNullOrEmpty(dm.Barcode) ? "RND_" + (Guid.NewGuid().ToString("N")).Substring(8) : dm.Barcode;
                Console.WriteLine(dm.SerialNumber_s + "_" + barcode);

                dict.Add(barcode, dm);

            }
            return dict;
        }
        public List<DellModel> CreateAndUpdateDellFile(List<DellModel> rmsColorList, List<DellModel> dellColorList, string tjValidDate)
        {
            Dictionary<string, DellModel> dellDict = GetDellColorListDict(dellColorList);

            foreach (DellModel dm in rmsColorList)
            {
                //mapping barcode
                if (dellDict.ContainsKey(dm.Barcode))
                    dellDict[dm.Barcode].Comment = dm.Comment;

                // TJ 要大於 valid date 才能新增
                else if ((dm.Barcode.Contains("TJ") && DateTime.Parse(dm.InputtingTime) > (DateTime.Parse(tjValidDate)))
                    || !dm.Barcode.Contains("TJ"))
                    dellDict.Add(dm.Barcode, dm);
            }

            return dellDict.Values.ToList();
        }
        #endregion
    }
}

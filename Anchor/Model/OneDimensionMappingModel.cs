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
    class OneDimensionMappingModel : BaseModel
    {
        public List<List<string>> OldLists = new List<List<string>>();
        public List<List<string>> NewLists = new List<List<string>>();
        public List<List<string>> ResultLists = new List<List<string>>();

        public string MappingOneDimensionData(string oldFilePath, string newFilePath, string savePath)
        {
            try
            {
                #region 舊2轉1檔案
                //IWorkbook => excel file
                IWorkbook oldFileWorkbook = null;
                using (FileStream oldFileFs = new FileStream(oldFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (oldFilePath.IndexOf(".xlsx") > 0)
                        oldFileWorkbook = new XSSFWorkbook(oldFileFs);

                    else if (oldFilePath.IndexOf(".xls") > 0)
                        oldFileWorkbook = new HSSFWorkbook(oldFileFs);

                    //ISheet => sheet
                    ISheet oldFileWorkbookSheet = oldFileWorkbook.GetSheetAt(0);
                    if (oldFileWorkbookSheet != null)
                    {
                        // 取得資料筆數
                        int oldFileLastRowNum = oldFileWorkbookSheet.LastRowNum;

                        // 取得最後一個欄位位置
                        int oldFileLastCellNum = oldFileWorkbookSheet.GetRow(0).LastCellNum;

                        for (int i = 0; i < (oldFileLastRowNum + 1); i++)
                        {
                            List<string> oldFileResultList = new List<string>();
                            for (int j = 0; j < oldFileLastCellNum; j++)
                            {
                                IRow curRow = oldFileWorkbookSheet.GetRow(i);
                                ICell icell = curRow.GetCell(j);
                                var cellValue = GetValueFromCell(icell);
                                oldFileResultList.Add(cellValue);
                            }
                            OldLists.Add(oldFileResultList);
                        }
                    }
                }
                #endregion

                #region 新2轉1檔案
                IWorkbook newFileWorkbook = null;
                using (FileStream newFileFs = new FileStream(newFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (newFilePath.IndexOf(".xlsx") > 0)
                        newFileWorkbook = new XSSFWorkbook(newFileFs);

                    else if (newFilePath.IndexOf(".xls") > 0)
                        newFileWorkbook = new HSSFWorkbook(newFileFs);

                    //ISheet => sheet
                    ISheet newFileWorkbookSheet = newFileWorkbook.GetSheetAt(0);
                    if (newFileWorkbookSheet != null)
                    {
                        // 取得資料筆數
                        int newFileLastRowNum = newFileWorkbookSheet.LastRowNum;

                        // 取得最後一個欄位位置
                        int newFileLastCellNum = newFileWorkbookSheet.GetRow(0).LastCellNum;

                        for (int i = 0; i < (newFileLastRowNum + 1); i++)
                        {
                            List<string> newFileResultList = new List<string>();
                            for (int j = 0; j < newFileLastCellNum; j++)
                            {
                                IRow curRow = newFileWorkbookSheet.GetRow(i);
                                ICell icell = curRow.GetCell(j);
                                var cellValue = GetValueFromCell(icell);
                                newFileResultList.Add(cellValue);
                            }
                            NewLists.Add(newFileResultList);
                        }
                    }
                }
                #endregion

                ResultLists = MappingToResultList(OldLists, NewLists);

                ResultListsToExcel(ResultLists, oldFilePath, newFilePath, savePath);

                return "OK";
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                MessageBox.Show(exception.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return exception.Message;
            }
        }

        public List<List<string>> MappingToResultList(List<List<string>> oldFileResultList, List<List<string>> newFileResultList)
        {
            // 以新資料為基底，比對舊資料
            for (int i = 0; i < newFileResultList.Count; i++)
            {
                List<string> resultList = new List<string>();
                List<string> newFileCellList = newFileResultList[i];
                String newFileKey = "";
                String newFileValue = "";
                
                for (int k = 0; k < (newFileCellList.Count - 1); k++)
                {
                    // 組新資料的Key值
                    newFileKey = newFileKey + newFileCellList[k];
                    resultList.Add(newFileCellList[k]);
                }
                // 取得新資料的Value值
                newFileValue = newFileCellList[(newFileCellList.Count - 1)];
                resultList.Add(newFileValue);

                // 新資料的Key比對不到舊資料的計數
                int newCount = 0;
                for (int j = 0; j < oldFileResultList.Count; j++)
                {
                    List<string> oldFileCellList = oldFileResultList[j];
                    String oldFileKey = "";
                    String oldFileValue = "";

                    for (int k = 0; k < (oldFileCellList.Count - 1); k++)
                    {
                        // 組舊資料的Key值
                        oldFileKey = oldFileKey + oldFileCellList[k];
                    }
                    // 取得舊資料的Value值
                    oldFileValue = oldFileCellList[(oldFileCellList.Count - 1)];
                    
                    if (newFileKey.Equals(oldFileKey))
                    {
                        // Key 相同、Value相同，以黑色顯示 表示 未更動
                        if (newFileValue.Equals(oldFileValue))
                        {
                            resultList.Add("O"); // O 代表 Original
                            break;
                        }
                        // Key 相同、Value不同，以紅色顯示 表示 有更動
                        else
                        {
                            resultList.Add("M"); // M 代表 Modify
                            break;
                        }
                    }
                    else
                    {
                        // 計數 + 1
                        newCount++;
                        // 當計數 等於 舊資料的筆數時，以藍色顯示 表示 新增
                        if (oldFileResultList.Count == newCount)
                        {
                            resultList.Add("A"); // A 代表 Add
                        }
                    }
                }
                ResultLists.Add(resultList);
            }

            // 以舊資料為基底，比對新資料有無刪除
            // 舊資料被刪除的筆數
            int deleteCount = 0;
            for (int i = 0; i < oldFileResultList.Count; i++)
            {
                List<string> resultList = new List<string>();
                List<string> oldFileCellList = oldFileResultList[i];
                String oldFileKey = "";
                String oldFileValue = "";
                for (int k = 0; k < (oldFileCellList.Count - 1); k++)
                {
                    // 組舊資料的Key值
                    oldFileKey = oldFileKey + oldFileCellList[k];
                    resultList.Add(oldFileCellList[k]);
                }
                // 取得舊資料的Value值
                oldFileValue = oldFileCellList[(oldFileCellList.Count - 1)];
                resultList.Add(oldFileValue);

                // 舊資料的Key比對不到新資料的計數
                int oldCount = 0;
                for (int j = 0; j < newFileResultList.Count; j++)
                {
                    List<string> newFileCellList = newFileResultList[j];
                    String newFileKey = "";

                    for (int k = 0; k < (newFileCellList.Count - 1); k++)
                    {
                        // 組新資料的Key值
                        newFileKey = newFileKey + newFileCellList[k];
                    }

                    if (!newFileKey.Equals(oldFileKey))
                    {
                        // 計數 + 1
                        oldCount++;
                        // 當計數 等於 新資料的筆數時，以灰色 + 刪除線顯示 表示 刪除
                        if (newFileResultList.Count == oldCount)
                        {
                            // 舊資料被刪除的筆數 + 1
                            deleteCount++;
                            resultList.Add("D"); // D 代表 Delete
                            // 從新資料當前筆數插入被刪除的那一筆資料
                            ResultLists.Insert((i + deleteCount), resultList);
                            break;
                        }
                    }
                }
            }
            return ResultLists;
        }

        public string ResultListsToExcel(List<List<string>> resultLists, string oldFilePath, string newFilePath, string saveFilePath)
        {
            // 先取得新資料的Data
            IWorkbook newFileWorkbook = null;
            ISheet newFileWorkbookSheet = null;
            using (FileStream newFileFs = new FileStream(newFilePath, FileMode.Open, FileAccess.Read))
            {
                if (newFilePath.IndexOf(".xlsx") > 0)
                    newFileWorkbook = new XSSFWorkbook(newFileFs);

                else if (newFilePath.IndexOf(".xls") > 0)
                    newFileWorkbook = new HSSFWorkbook(newFileFs);
                //ISheet => sheet
                newFileWorkbookSheet = newFileWorkbook.GetSheetAt(0);
            }

            IWorkbook workbook = null;
            using (FileStream templateFile = new FileStream(oldFilePath, FileMode.Open, FileAccess.Read))
            {
                if (oldFilePath.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook(templateFile);
                }
                else if (oldFilePath.IndexOf(".xls") > 0)
                {
                    workbook = new HSSFWorkbook(templateFile);
                }
                // 第一張Sheet改名稱為 Old_OneDimensionDataResult
                workbook.SetSheetName(0, "Old_OneDimensionDataResult");

                // 第二張Sheet為新資料的 Sheet，名稱為 New_OneDimensionDataResult
                newFileWorkbookSheet.CopyTo(workbook, "New_OneDimensionDataResult", true, true);

                // 第三張Sheet為新舊資料比對的結果
                ISheet sheet = workbook.CreateSheet("MappingResult");
                // 設定比對結果的sheet為active
                workbook.SetActiveSheet(2);
                workbook.SetSelectedTab(2);
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

                #region 灰色 + 刪除線模板
                //設定處存格 style(灰色)
                XSSFCellStyle cellStyle_gray_strikeout = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_gray_strikeout.BorderBottom = BorderStyle.Thin;
                cellStyle_gray_strikeout.BorderLeft = BorderStyle.Thin;
                cellStyle_gray_strikeout.BorderRight = BorderStyle.Thin;
                cellStyle_gray_strikeout.BorderTop = BorderStyle.Thin;
                //設定字體sytle
                XSSFFont font_gray_strikeout = (XSSFFont)workbook.CreateFont();
                font_gray_strikeout.FontName = "Calibri";//字體
                font_gray_strikeout.SetColor(new XSSFColor(new byte[] { 128, 128, 128 }));
                font_gray_strikeout.IsStrikeout = true;//刪除線
                cellStyle_gray_strikeout.SetFont(font_gray_strikeout);
                #endregion

                IRow row = null;
                ICell cell = null;
                //寫入資料
                for (int i = 0; i < resultLists.Count; i++)
                {
                    row = sheet.CreateRow(i);
                    int cellIndex = 0;

                    for (int j = 0; j < (resultLists[i].Count - 1); j++)
                    {
                        // O 代表 Original
                        if ("O".Equals(resultLists[i][resultLists[i].Count - 1]))
                        {
                            //放值
                            cell = row.CreateCell(j);
                            cell.SetCellValue(resultLists[i][j]);
                            //逐格設定顏色, CellStyle超過65000個就會爆掉
                            cell.CellStyle = cellStyle_black;
                            cellIndex++;
                        }
                        // M 代表 Modify
                        if ("M".Equals(resultLists[i][resultLists[i].Count - 1]))
                        {
                            //放值
                            cell = row.CreateCell(j);
                            cell.SetCellValue(resultLists[i][j]);
                            //逐格設定顏色, CellStyle超過65000個就會爆掉
                            cell.CellStyle = cellStyle_red;
                            cellIndex++;
                        }
                        // A 代表 Add
                        if ("A".Equals(resultLists[i][resultLists[i].Count - 1]))
                        {
                            //放值
                            cell = row.CreateCell(j);
                            cell.SetCellValue(resultLists[i][j]);
                            //逐格設定顏色, CellStyle超過65000個就會爆掉
                            cell.CellStyle = cellStyle_blue;
                            cellIndex++;
                        }
                        // D 代表 Delete
                        if ("D".Equals(resultLists[i][resultLists[i].Count - 1]))
                        {
                            //放值
                            cell = row.CreateCell(j);
                            cell.SetCellValue(resultLists[i][j]);
                            //逐格設定顏色, CellStyle超過65000個就會爆掉
                            cell.CellStyle = cellStyle_gray_strikeout;
                            cellIndex++;
                        }
                    }
                }
                //寫檔
                SaveWorkbook(workbook, saveFilePath);
            }
            return "OK";
        }
    }
}

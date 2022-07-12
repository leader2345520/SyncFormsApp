using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Anchor.Model
{
    class PoModel : BaseModel
    {
        #region getter/setter
        public string Project { get; set; }
        public int ClosedCount { get; set; }
        public int OpenCount { get; set; }
        public int Over2monthsCount { get; set; }
        #endregion

        public string AnalysisWeeklyReport(string filePath, string savePath)
        {
            try
            {
                IWorkbook workbook = OpenWorkbook(filePath);

                ISheet weeklySheet = workbook.GetSheet("Status of All Projects");

                CopyRange(CellRangeAddress.ValueOf("B6:M8"), weeklySheet);

                AddCellMergeRegion(weeklySheet);

                DictToExcel(PoDataTableToDict(ExcelToDataTable(filePath, "總表", 1)), weeklySheet);

                weeklySheet.ForceFormulaRecalculation = true;

                SaveWorkbook(workbook, savePath);

                return "OK";
            }
            catch (Exception e)
            {
                return e.Message;

            }

        }
        public Dictionary<string, PoModel> PoDataTableToDict(DataTable dt)
        {
            Dictionary<string, PoModel> dict = new Dictionary<string, PoModel>();

            if (dt != null)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (!dict.ContainsKey(row["Project"].ToString()))
                    {
                        PoModel em = new PoModel
                        {
                            Project = row["Project"].ToString(),
                            ClosedCount = 0,
                            OpenCount = 0,
                            Over2monthsCount = 0
                        };
                        dict.Add(row["Project"].ToString(), em);
                    }

                    if ("Close".Equals(row["Status"].ToString()))
                        dict[row["Project"].ToString()].ClosedCount++;
                    else if (Convert.ToInt32(row["Delay days"].ToString()) >= 60)
                        dict[row["Project"].ToString()].Over2monthsCount++;
                    else
                        dict[row["Project"].ToString()].OpenCount++;




                }

            }
            return dict;
        }
        public void DictToExcel(Dictionary<string, PoModel> dict, ISheet weeklySheet)
        {
            int rowCount = weeklySheet.LastRowNum;
            string week = GetWeek(weeklySheet, rowCount);
            List<int> list = GetCountList(dict);
            int summaryRowNum = GetSummaryLastRowNum(weeklySheet);

            for (int i = rowCount - 2; i < rowCount + 1; i++)
            {
                IRow row = weeklySheet.GetRow(i);
                if (i == rowCount - 2)
                {
                    for (int j = 1; j < 12; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (j == 1)
                            cell.SetCellValue(week);
                        else if (j == 2)
                        {
                            cell.SetCellValue(DateTime.Today);

                            //右側summary for chart
                            weeklySheet.GetRow(summaryRowNum).CreateCell(14);
                            CopyCell(weeklySheet.GetRow(2).GetCell(14), weeklySheet.GetRow(summaryRowNum).GetCell(14));
                            weeklySheet.GetRow(summaryRowNum).GetCell(14).SetCellValue(DateTime.Today);
                        }
                        else
                        {
                            cell.SetCellValue(list[j - 3]);

                            //右側summary for chart
                            weeklySheet.GetRow(summaryRowNum).CreateCell(j + 12);
                            CopyCell(weeklySheet.GetRow(2).GetCell(j + 12), weeklySheet.GetRow(summaryRowNum).GetCell(j + 12));
                            weeklySheet.GetRow(summaryRowNum).GetCell(j + 12).SetCellValue(list[j - 3]);
                        }
                    }
                    row.GetCell(12).SetCellValue("");
                }
                else if (i == rowCount - 1)
                {
                    for (int k = 3; k < 12; k++)
                    {
                        if (row.GetCell(k).CellType == CellType.Formula)
                        {
                            if (k < 6)
                                row.GetCell(k).SetCellFormula("SUM(D" + i + ":F" + i + ")");
                            else if (k < 9)
                                row.GetCell(k).SetCellFormula("SUM(G" + i + ":I" + i + ")");
                            else
                                row.GetCell(k).SetCellFormula("SUM(J" + i + ":L" + i + ")");
                        }
                    }
                }
                else
                {
                    int preRow = i - 1;
                    for (int l = 3; l < 12; l++)
                    {
                        if (row.GetCell(l).CellType == CellType.Formula)
                        {
                            if (l < 6)
                                row.GetCell(l).SetCellFormula("D" + preRow + "/D" + i + "*100%");
                            else if (l < 9)
                                row.GetCell(l).SetCellFormula("G" + preRow + "/G" + i + "*100%");
                            else
                                row.GetCell(l).SetCellFormula("J" + preRow + "/J" + i + "*100%");
                        }
                    }
                }
            }
        }
        public int GetSummaryLastRowNum(ISheet sheet)
        {
            int i = 0;

            while (sheet.GetRow(i).GetCell(14) != null)
            {
                i++;
            }
            return i;
        }
        public List<int> GetCountList(Dictionary<string, PoModel> dict)
        {
            List<int> list = new List<int>
            {
                dict["Acqua"].ClosedCount,
                dict["Acqua"].OpenCount,
                dict["Acqua"].Over2monthsCount,
                dict["Mineral Wells"].ClosedCount,
                dict["Mineral Wells"].OpenCount,
                dict["Mineral Wells"].Over2monthsCount,
                dict["Patras"].ClosedCount,
                dict["Patras"].OpenCount,
                dict["Patras"].Over2monthsCount
            };

            return list;
        }
        public string GetWeek(ISheet weeklySheet, int rowCount)
        {
            int num = Convert.ToInt32(weeklySheet.GetRow(rowCount - 5).GetCell(1).ToString().Replace("week", "")) + 1;
            string result = "week " + num;

            return result;
        }
        void CopyRange(CellRangeAddress range, ISheet weeklySheet)
        {
            int lastRowNum = weeklySheet.LastRowNum;

            for (var rowNum = range.FirstRow; rowNum <= range.LastRow; rowNum++)
            {
                IRow sourceRow = weeklySheet.GetRow(rowNum);

                weeklySheet.CreateRow(++lastRowNum);

                if (sourceRow != null)
                {
                    IRow destinationRow = weeklySheet.GetRow(weeklySheet.LastRowNum);

                    for (var col = range.FirstColumn; col < sourceRow.LastCellNum && col <= range.LastColumn; col++)
                    {
                        destinationRow.CreateCell(col);
                        CopyCell(sourceRow.GetCell(col), destinationRow.GetCell(col));
                    }
                }
            }
            //copy cell width
            /*for (int i = 0; i < sourceSheet.GetRow(0).LastCellNum; i++)
            {
                destinationSheet.SetColumnWidth(i + shiftNum, sourceSheet.GetColumnWidth(i));
            }*/
        }
        void AddCellMergeRegion(ISheet destinationSheet)
        {
            int date = destinationSheet.LastRowNum - 2;
            int subtotal = destinationSheet.LastRowNum - 1;
            int ach = destinationSheet.LastRowNum;
            destinationSheet.AddMergedRegion(new CellRangeAddress(date, ach, 1, 1));
            destinationSheet.AddMergedRegion(new CellRangeAddress(subtotal, subtotal, 3, 5));
            destinationSheet.AddMergedRegion(new CellRangeAddress(ach, ach, 3, 5));
            destinationSheet.AddMergedRegion(new CellRangeAddress(subtotal, subtotal, 6, 8));
            destinationSheet.AddMergedRegion(new CellRangeAddress(ach, ach, 6, 8));
            destinationSheet.AddMergedRegion(new CellRangeAddress(subtotal, subtotal, 9, 11));
            destinationSheet.AddMergedRegion(new CellRangeAddress(ach, ach, 9, 11));
            destinationSheet.AddMergedRegion(new CellRangeAddress(date, ach, 12, 12));

        }
        void CopyCell(ICell source, ICell destination)
        {
            if (destination != null && source != null)
            {
                //you can comment these out if you don't want to copy the style ...
                destination.CellComment = source.CellComment;
                destination.Hyperlink = source.Hyperlink;

                destination.CellStyle = source.CellStyle;
                //注意新建單元格樣式的時候一定要克隆一下，否則單元格之前的樣式就沒了
                /*XSSFCellStyle cellStyle_cust = (XSSFCellStyle)workbook.CreateCellStyle();
                cellStyle_cust.CloneStyleFrom(source.CellStyle);
                destination.CellStyle = cellStyle_cust;*/

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


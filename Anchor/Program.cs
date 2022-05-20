using Aspose.Words;
using Aspose.Words.Tables;
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
using Table = Aspose.Words.Tables.Table;

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

    }

    //BaseModel
    class BaseModel
    {
        public DataTable ExcelToDataTable(string filePath, object sheetPage, int copyFormatRow)
        {
            DataTable dt = new DataTable();
            DataColumnCollection columns;
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
                            string cellValue = GetValueFromCell(icell);


                            if (isFirst)
                            {
                                columns = dt.Columns;

                                //如果欄位重複值
                                if (columns.Contains(cellValue))
                                    dt.Columns.Add(GetDistinctColumn(cellValue));
                                //如果欄位空值, 往上找看有無合併儲存格
                                else if (string.IsNullOrEmpty(cellValue))
                                    dt.Columns.Add(GetMergeColumn(sheet, i, j));
                                else
                                    dt.Columns.Add(cellValue);
                            }
                            else if (j < dt.Columns.Count) //避免資料 cell 長度大於 column 的長度
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
        public string GetDistinctColumn(string cellValue)
        {
            return cellValue + DateTime.Now.ToString(("_HHmmss"));
        }
        public string GetValueFromCell(ICell icell)
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
        public string GetMergeColumn(ISheet sheet, int i, int j)
        {

            if (i < 0) return "";
            if (string.IsNullOrEmpty(GetValueFromCell(sheet.GetRow(i).GetCell(j))))
                return GetMergeColumn(sheet, --i, j);
            else
                return GetValueFromCell(sheet.GetRow(i).GetCell(j));

        }
    }


}


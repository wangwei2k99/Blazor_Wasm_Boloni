using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace AppCfg.Npoi
{
    public static class NPOI
    {
        public static DataTable DtempToDt(DataTable Dt_Temp, int StartCol, int EndCol)//StartCol,EndCol从1开始
        {
            for (int i = Dt_Temp.Columns.Count-1; i >=0; i--)
            {
                if (i<StartCol-1||i>=EndCol)
                {
                    Dt_Temp.Columns.RemoveAt(i);
                }
            }
            return Dt_Temp;
        }
        public static DataSet MemstreamToDataSet(MemoryStream ms)
        {
            DataSet ds = new DataSet();
            IWorkbook workbook;
            //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
            workbook = new XSSFWorkbook(ms);
            if (workbook == null) { return null; }
            Dictionary<int, string> SheetNames = new Dictionary<int, string>();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                SheetNames.Add(i, workbook.GetSheetName(i));
            }
            foreach (KeyValuePair<int, string> kvp in SheetNames)
            {
                DataTable dt = new DataTable();
                ISheet Sheet = workbook.GetSheetAt(kvp.Key);
                IRow header = Sheet.GetRow(Sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = Sheet.FirstRowNum; i <= header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString().Trim() == "")
                    {
                        dt.Columns.Add(new DataColumn("Col" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据  
                for (int i = Sheet.FirstRowNum + 1; i <= Sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        if (Sheet.GetRow(i) != null)
                        {
                            dr[j] = GetValueType(Sheet.GetRow(i).GetCell(j));
                            if (dr[j] != null && dr[j].ToString() != string.Empty)
                            {
                                hasValue = true;
                            }
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
                //表头  
                dt.TableName = kvp.Value.Trim();
                ds.Tables.Add(dt);
            }
            return ds;
        }
        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="File">上传的Excel路径</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string File, string SheetName, int StartCol = 0, int EndCol = 0)//StartCol,EndCol从1开始
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(File).ToLower();
            using (FileStream fs = new FileStream(File, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx" || fileExt == ".xlsm") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                Dictionary<int, string> SheetNames = new Dictionary<int, string>();
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    SheetNames.Add(i, workbook.GetSheetName(i));
                }
                foreach (KeyValuePair<int, string> kvp in SheetNames)
                {
                    if (kvp.Value.Trim() == SheetName)
                    {
                        ISheet Sheet = workbook.GetSheetAt(kvp.Key);
                        IRow header = Sheet.GetRow(Sheet.FirstRowNum);
                        List<int> columns = new List<int>();
                        if (StartCol == 0 && EndCol == 0)
                        {
                            StartCol = 1;
                            EndCol = header.LastCellNum + 1;
                        }
                        else if (StartCol != 0 && EndCol == 0)
                        {
                            EndCol = header.LastCellNum + 1;
                        }
                        for (int i = StartCol - 1; i <= EndCol - 1; i++)
                        {
                            object obj = GetValueType(header.GetCell(i));
                            if (obj == null || obj.ToString().Trim() == "")
                            {
                                dt.Columns.Add(new DataColumn("Col" + i.ToString()));
                            }
                            else
                                dt.Columns.Add(new DataColumn(obj.ToString()));
                            columns.Add(i);
                        }
                        //数据  
                        for (int i = Sheet.FirstRowNum + 1; i <= Sheet.LastRowNum; i++)
                        {
                            DataRow dr = dt.NewRow();
                            bool hasValue = false;
                            foreach (int j in columns)
                            {
                                if (Sheet.GetRow(i) != null)
                                {
                                    dr[j] = GetValueType(Sheet.GetRow(i).GetCell(j));
                                    if (dr[j] != null && dr[j].ToString() != string.Empty)
                                    {
                                        hasValue = true;
                                    }
                                }
                            }
                            if (hasValue)
                            {
                                dt.Rows.Add(dr);
                            }
                        }
                        //表头  
                    }
                }
            }
            return dt;
        }
        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                    //return "=" + cell.CellFormula;
            }
        }
        /// <summary>
        /// Datable导出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void TableToExcel(DataTable dt, string file, string ShtName = "ExcelToMes")
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet(ShtName) : workbook.CreateSheet(dt.TableName);
            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            //数据  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }
            AutoColumnWidth(sheet, 50);
            //转为字节数组  
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();
            //保存为Excel文件  
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }

        }
        /// <summary>
        /// 批量设置Col值
        /// </summary>
        public static void SetColValue(ref DataTable dt, string str, string ColName)
        {
            foreach (DataRow row in dt.Rows)
            {
                row[ColName] = str;
            }
        }
        /// <summary>
        /// 批量设置Col宽
        /// </summary>
        public static void AutoColumnWidth(ISheet sheet, int Cols)
        {
            for (int col = 0; col <= Cols; col++)
            {
                sheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
                int columnWidth = sheet.GetColumnWidth(col) / 256;//获取当前列宽度
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    ICell cell = row.GetCell(col);
                    if (cell != null)
                    {
                        int contextLength = Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度
                        columnWidth = columnWidth < contextLength ? contextLength : columnWidth;
                    }
                }
                sheet.SetColumnWidth(col, columnWidth * 200);//

            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;

namespace ExcelTool
{
    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            SheetNames = new List<string>();
            disposed = false;
            loadInit();
        }

        private void getWorkbookInstance()
        {
            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.ToLower().IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook(fs);
            else if (fileName.ToLower().IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook(fs);
        }

        private void loadInit()
        {
            getWorkbookInstance();
            if (workbook != null)
            {
                int num = workbook.NumberOfSheets;
                for (int k = 0; k < num; k++)
                {
                    SheetNames.Add(workbook.GetSheetName(k));
                }
            }
        }

        public List<String> GetAllSheetName()
        {
            return SheetNames;
        }
        public List<string> SheetNames { get; set; }


        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception: " + ex.Message);
                log(ex);
             
            }
            return -1;
        }

        /// <summary>
        /// 获取第<paramref name="idx"/>的sheet的数据
        /// </summary>
        /// <param name="idx">Excel文件的第几个sheet表</param>
        /// <param name="isFirstRowCoumn">是否将第一行作为列标题</param>
        /// <returns></returns>
        public DataTable GetTable(int idx, int headIndex)
        {
            if (idx >= SheetNames.Count || idx < 0)
                throw new Exception("Do not Get This Sheet");
            return ExcelToDataTable(SheetNames[idx], headIndex);
        }


        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="headIndex">第N行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, int headIndex)
        {
            ISheet sheet = null;
            var data = new DataTable();
            data.TableName = sheetName;
            int startRow = 0;
            try
            {
                sheet = sheetName != null ? workbook.GetSheet(sheetName) : workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    var firstRow = sheet.GetRow(headIndex);
                    if (firstRow == null)
                        return data;
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                    startRow = headIndex;// isFirstRowColumn ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;

                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        var column = new DataColumn(Convert.ToChar(((int)'A') + i).ToString());
                        //if (isFirstRowColumn)
                        //{
                        //    var columnName = firstRow.GetCell(i).StringCellValue;
                        //    column = new DataColumn(columnName);
                        //}
                        data.Columns.Add(column);
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                else throw new Exception("Don not have This Sheet");
                return data;
            }
            catch (Exception ex)
            {
                log(ex);
                throw ex;
            }
        }



        public DataSet ExcelToDataSet(int headIndex)
        {
            DataSet dst = new DataSet();
            try
            {
                foreach (string sname in SheetNames)
                {
                    dst.Tables.Add(ExcelToDataTable(sname, headIndex));
                }
                return dst;
            }
            catch (Exception ex)
            {
               // Console.WriteLine("Exception: " + ex.Message);
                log(ex);
                throw ex;
            }
        }

        private DataTable setToDataTable(int startRow, ISheet sheet)
        {
            DataTable data = new DataTable();
            IRow firstRow = sheet.GetRow(startRow);
            int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

            for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
            {
                ICell cell = firstRow.GetCell(i);
                if (cell != null)
                {
                    string cellValue = cell.StringCellValue;
                    if (cellValue != null)
                    {
                        DataColumn column = new DataColumn(cellValue);
                        data.Columns.Add(column);
                    }
                }
            }
            //   startRow = sheet.FirstRowNum + 1;

            //最后一列的标号
            int rowCount = sheet.LastRowNum;
            for (int i = startRow; i <= rowCount; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue; //没有数据的行默认是null　　　　　　　

                DataRow dataRow = data.NewRow();
                for (int j = row.FirstCellNum; j < cellCount; ++j)
                {
                    if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        dataRow[j] = row.GetCell(j).ToString();
                }
                data.Rows.Add(dataRow);
            }
            return data;
        }
    

        
        /// <summary>
        /// 导出EXCEL,可以导出多个sheet
        /// </summary>
        /// <param name="dtSources">原始数据数组类型</param>
        /// <param name="strFileName">路径</param>
        public void ExportToExcel(List<DataTable> dtSources)
        {
            for (int k = 0; k < dtSources.Count; k++)
            {
                String name = dtSources[k].TableName.ToString();
                if (name.Length > 20)
                {
                    name = name.Substring(0, 20);
                }
                ISheet sheet = workbook.CreateSheet(name);
                //填充表头
                IRow dataRow = sheet.CreateRow(0);
                foreach (DataColumn column in dtSources[k].Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                }

                //填充内容
                for (int i = 0; i < dtSources[k].Rows.Count; i++)
                {
                    dataRow = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dtSources[k].Columns.Count; j++)
                    {
                        dataRow.CreateCell(j).SetCellValue(dtSources[k].Rows[i][j].ToString());
                    }
                }
            }
            workbook.Write(fs);
        }
        public void SaveToExcel(DataSet dtSources)
        {
            foreach (DataTable dt in dtSources.Tables)
            {
                ExportToExcel(dt);
            }
        }

        /// <summary>
        /// 导出单个EXCEL
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="strFileName"></param>
        public void ExportToExcel(DataTable dtSource)
        {
            ISheet sheet = workbook.CreateSheet();
            //填充表头
            IRow dataRow = sheet.CreateRow(0);
            foreach (DataColumn column in dtSource.Columns)
            {
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            }
            //填充内容
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dtSource.Rows[i][j].ToString());
                }
            }
            workbook.Write(fs);
             
            //workbook.Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }

        private static void log(string message)
        {
            LogHelper.Log(typeof(ExcelHelper), message);
        }
        private static void log(Exception ex)
        {
            LogHelper.Log(typeof(ExcelHelper), ex);
        }

    }
}


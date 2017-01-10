using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace ExcelTool
{
    public class oledbExcel
    {

        public static DataTable GetDataFromExcelByConn(String filePath, bool hasTitle = false)
        {
            using (DataSet ds = new DataSet())
            {
                using (OleDbConnection conn = GetOleDbConnection(filePath,hasTitle))
                {
                    System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim();
                    string strCom = " SELECT * FROM  [" + sheetName + "]";//[Sheet1$]

                    using (OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, conn))
                    {
                        if (conn.State == ConnectionState.Broken || conn.State == ConnectionState.Closed)
                        {
                            conn.Open();
                        }
                        myCommand.Fill(ds);
                    }
                }
                if (ds == null || ds.Tables.Count <= 0) return null;
                return ds.Tables[0];
            }
        }


        public static OleDbConnection GetOleDbConnection(String filePath, Boolean hasTitle)
        {
            string fileType = System.IO.Path.GetExtension(filePath);
            if (!File.Exists(filePath))
            {
                return null;
            }
            if (string.IsNullOrEmpty(fileType))
            {
                return null;
            }

            string strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                               "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                               "data source={3};",
                              (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
           // string strCon = string.Format("Provider=Microsoft.ACE.OLEDB.{0}.0;" +
           //                "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
           //                "data source={3};",
            //              12, (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
            OleDbConnection conn = new OleDbConnection(strCon);
            return conn;

        }

        /// <summary> 
        /// 将Excel读取到DataSet 
        /// </summary> 
        /// <param name="path">Excel 路径</param> 
        /// <param name="excelversion">12.0</param> 
        /// <returns></returns> 
        public static DataSet ExcelToDataSet(string path, bool hasTitle = false)
        {
            try
            {
                // 拼写连接字符串，打开连接 
               // string strConn = "Provider=Microsoft.ACE.OLEDB." + excelversion + ";" + "Data Source=" + path + ";Extended Properties='Excel " + excelversion + "; HDR=YES; IMEX=1'";

                OleDbConnection objConn = GetOleDbConnection(path, hasTitle);
                if(objConn==null)
                {
                    return null;
                }
                objConn.Open();
                // 取得Excel工作簿中所有工作表 
                DataTable schemaTable = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                OleDbDataAdapter sqlada = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                // 遍历工作表取得数据并存入Dataset 
                foreach (DataRow dr in schemaTable.Rows)
                {
                    string strSql = "Select * From [" + dr[2].ToString().Trim() + "]";
                    OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
                    sqlada.SelectCommand = objCmd;
                    sqlada.Fill(ds, dr[2].ToString().Trim());
                }
                objConn.Close();
                return ds;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

        }


        /// <summary>
        /// 从System.Data.DataTable导入数据到数据库
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        //public int InsetData(System.Data.DataTable dt)
        //{
        //    int i = 0;
        //    string lng = "";
        //    string lat = "";
        //    string offsetLNG = "";
        //    string offsetLAT = "";

        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        lng = dr["LNG"].ToString().Trim();
        //        lat = dr["LAT"].ToString().Trim();
        //        offsetLNG = dr["OFFSET_LNG"].ToString().Trim();
        //        offsetLAT = dr["OFFSET_LAT"].ToString().Trim();

        //        //sw = string.IsNullOrEmpty(sw) ? "null" : sw;
        //        //kr = string.IsNullOrEmpty(kr) ? "null" : kr;

        //        string strSql = string.Format("Insert into DBToExcel (LNG,LAT,OFFSET_LNG,OFFSET_LAT) Values ('{0}','{1}',{2},{3})", lng, lat, offsetLNG, offsetLAT);

        //        string strConnection = ConfigurationManager.ConnectionStrings["ConnectionStr"].ToString();
        //        SqlConnection sqlConnection = new SqlConnection(strConnection);
        //        try
        //        {
        //            // SqlConnection sqlConnection = new SqlConnection(strConnection);
        //            sqlConnection.Open();
        //            SqlCommand sqlCmd = new SqlCommand();
        //            sqlCmd.CommandText = strSql;
        //            sqlCmd.Connection = sqlConnection;
        //            SqlDataReader sqlDataReader = sqlCmd.ExecuteReader();
        //            i++;
        //            sqlDataReader.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            throw ex;
        //        }
        //        finally
        //        {
        //            sqlConnection.Close();

        //        }
        //        //if (opdb.ExcSQL(strSql))
        //        //    i++;
        //    }
        //    return i;
        //}

    }
}

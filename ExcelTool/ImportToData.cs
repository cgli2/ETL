using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;
　



namespace ExcelTool
{
    public class ImportToData
    {
      　
        public static int saveBase(string taxId,string taxName,string certificate,string declareDate,string fillingDate,string reportDate)
        {
            String sql = "insert into cys_customer_report_base(taxpayer_id,taxpayer_name,certificate,declare_date,filling_date,report_date,create_date) values('" + taxId + "','" + taxName + "','" + certificate + "','" + declareDate + "','" + fillingDate + "','" + reportDate + "',NOW())";
            LogHelper.Log(typeof(ImportToData), sql);
            return MysqlHelper.ExecuteSql(sql);
        }

        public static int saveBalance(int baseId, String assetName, int lineNo, Decimal beginBalance, Decimal endingBalance)
        {
            String sql = "insert into cys_customer_balance(report_id,asset_name,line_no,begin_balance,ending_balance)values('"+baseId+"','"+assetName+"','"+lineNo+"','"+beginBalance+"','"+endingBalance+"')";
            return MysqlHelper.ExecuteSql(sql);
        }


        public static int saveProfit(int baseId, String profitItem, int lineNo, Decimal yearAmount, Decimal monthAmount)
        {
            String sql = "insert into cys_customer_balance(report_id,profit_item,line_no,year_amount,month_amount)values('" + baseId + "','" + profitItem + "','" + lineNo + "','" + yearAmount + "','" + monthAmount + "')";
            return MysqlHelper.ExecuteSql(sql);
        }


        private static int insertBatchProfit(String sqlValues)
        {
            String sql = "insert into cys_customer_balance(report_id,profit_item,line_no,year_amount,month_amount)values";
            sql += sqlValues;
            return MysqlHelper.ExecuteSql(sql);
        }


        private static int insertBatchBalance(String sqlValues)
        {
            String sql = "insert into cys_customer_balance(report_id,asset_name,line_no,begin_balance,ending_balance)values";
            sql += sqlValues;
            return MysqlHelper.ExecuteSql(sql);
        }


        private static int insertBatchCashFlow(String sqlValues)
        {
            String sql = "insert into cys_customer_cash_flow(report_id,cash_item,line_no,amount)values";
            sql += sqlValues;
            return MysqlHelper.ExecuteSql(sql);
        }



        private static int insertBatchCompany(String sqlValues)
        {
            String sql = "insert into hh_company(id,name,linkman,phone,address,reg_date)values";
            sql += sqlValues;
            return MysqlHelper.ExecuteSql(sql);
        }

        private static void OperateCompanyData(DataTable data)
        {
            if (data == null) return;
            StringBuilder sb = new StringBuilder();
            int batchCount = 0;

            for (int i = 0; i < data.Rows.Count; ++i)
            {
                string id = Guid.NewGuid().ToString().Replace("-", "");
                string name = data.Rows[i][0].ToString();
                string linkman = data.Rows[i][1].ToString();
                string phone = data.Rows[i][2].ToString();
                string address = data.Rows[i][3].ToString();
                string regDate = data.Rows[i][4].ToString();
                sb.Append("(");
                sb.Append("'" + id + "',");
                sb.Append("'" + name + "',");
                sb.Append("'" + linkman + "',");
                sb.Append("'" + phone + "',");
                sb.Append("'" + address + "',");
                sb.Append("'" + regDate + "'),");

                if (batchCount >= 1000)
                {
                    if (sb.Length > 0)
                    {
                        string sql = sb.ToString().TrimEnd(',') + ";";
                        log(sql);
                        insertBatchCompany(sql);
                        sb.Length = 0;
                    }
                }
                batchCount++;
            }
            if (sb.Length > 0)
            {
                string sql = sb.ToString().TrimEnd(',') + ";";
                log(sql);
                insertBatchCompany(sql);
                sb.Length = 0;
            }
        }
        public static void ImportAllCompany(string file)
        {
            try
            {
                using (ExcelHelper excelHelper = new ExcelHelper(file))
                {
                    DataSet dst = excelHelper.ExcelToDataSet(0);
                    foreach (DataTable dt in dst.Tables)
                    {
                        OperateCompanyData(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                log(ex);
            }
        }
        private static void OperateProfitData(DataTable data)
        {
            if (data == null) return;

            int id = getBaseId(data);
            if (id <= 0)
            {
               log("The invalide base id=" + id);
                return;
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < data.Rows.Count; ++i)
            {
                if (i > 6)
                {
                    string name = data.Rows[i][0].ToString();
                    int line = int.Parse(data.Rows[i][1].ToString());
                    Decimal yearAmount = 0;
                    Decimal monthAmount = 0;
                    if (data.Rows[i][2] != null)
                    {
                        Decimal.TryParse(data.Rows[i][2].ToString(), out yearAmount);
                    }
                    if (data.Rows[i][ 3] != null)
                    {
                        Decimal.TryParse(data.Rows[i][3].ToString(), out monthAmount);
                    }
                    sb.Append("(");
                    sb.Append("'" + id + "',");
                    sb.Append("'" + name + "',");
                    sb.Append("'" + line + "',");
                    sb.Append("'" + yearAmount + "',");
                    sb.Append("'" + monthAmount + "'),");
                    //if (i == data.Rows.Count - 1)
                    //{
                    //    sb.Append("'" + monthAmount + "');");
                    //}
                    //else
                    //{
                    //    sb.Append("'" + monthAmount + "'),");
                    //}
                }
            }
          //  log(sb.ToString());
          // log.info("insert--------------------->" + sb.ToString());
          //  insertBatchProfit(sb.ToString());
          //  sb.Length = 0;

            if (sb.Length > 0)
            {
                string sql = sb.ToString().TrimEnd(',') + ";";
                log(sql);
                insertBatchProfit(sql);
                sb.Length = 0;
            }
        }



        private static void log(string message){
             LogHelper.Log(typeof(ImportToData),message);
        }
        private static void log(Exception ex)
        {
            LogHelper.Log(typeof(ImportToData), ex);
        }


        private static void OperateBalanceData(DataTable data)
        {
            if (data == null) return;
            int id = getBaseId(data);
            if (id <= 0)
            {
              log("The invalide base id="+id);
                return;
            }
            StringBuilder sb = new StringBuilder();
            int maxCell = 8;
            Decimal yearAmount = 0;
            Decimal monthAmount = 0;
            for (int j = 0; j < maxCell; )
            {
                //start from line 8
                for (int i = 7; i < data.Rows.Count; ++i)
                {
                    if (data.Rows[i][j+1] == null || data.Rows[i][j+1].ToString().Length == 0)
                    {
                        continue;
                    }
                    string name = data.Rows[i][j].ToString();
                    int line = int.Parse(data.Rows[i][j+1].ToString());
                    yearAmount = 0;
                    monthAmount = 0;
                    if (data.Rows[i][j + 2] != null)
                    {
                        Decimal.TryParse(data.Rows[i][j + 2].ToString(), out yearAmount);
                    }
                    if (data.Rows[i][j + 3] != null)
                    {
                        Decimal.TryParse(data.Rows[i][j+3].ToString(),out monthAmount);
                    }
                    sb.Append("(");
                    sb.Append("'" + id + "',");
                    sb.Append("'" + name + "',");
                    sb.Append("'" + line + "',");
                    sb.Append("'" + yearAmount + "',");
                    sb.Append("'" + monthAmount + "'),");

                    //if (i == (data.Rows.Count - 1) && (j == maxCell - 1))
                    //{
                    //    sb.Append("'" + monthAmount + "');");
                    //}
                    //else
                    //{
                    //    sb.Append("'" + monthAmount + "'),");
                    //}
                }
                j += maxCell / 2;
            }
           
            if (sb.Length > 0)
            {
                string sql = sb.ToString().TrimEnd(',') + ";";
                log(sql);
                insertBatchBalance(sql);
                sb.Length = 0;
            }
          
        }

        private static int getBaseId(DataTable data)
        {
            for (int i = 0; i < 6; i++)
            {
                if (data.Rows[i][1] == null || data.Rows[i][1].ToString().Length == 0)
                {
                    Console.WriteLine("表头[" + data.Rows[i][0] + "]必填信息为空!");
                    return 0;
                }
            }
            object obj = data.Rows[0][1];
            if (null == obj)
            {
                log("提交的数据缺失！" + data.TableName);
                return 0;
            }
            string taxpayer_id = data.Rows[0][1].ToString();
            string taxpayer_name = data.Rows[1][1].ToString();
            string certificate = data.Rows[2][1].ToString();

            DateTime dt = DateTime.Now;
            string declare_date = DateTime.Parse(data.Rows[3][1].ToString()).ToString("yyyy-MM-dd");
            string filling_date = DateTime.Parse(data.Rows[4][1].ToString()).ToString("yyyy-MM-dd");
            string report_date = DateTime.Parse(data.Rows[5][1].ToString()).ToString("yyyy-MM-dd");

            int id = ImportToData.saveBase(taxpayer_id, taxpayer_name, certificate, declare_date, filling_date, report_date);
            return id;
        }

        private static void OperateCashFlowData(DataTable data)
        {
            if (data == null) return;
            int id = getBaseId(data);
            if (id <= 0)
            {
                log("The invalide base id=" + id);
                return;
            }
            StringBuilder sb = new StringBuilder();
            int maxCell = 6;
            for (int j = 0; j < maxCell; )
            {
                //start from line 8
                for (int i = 7; i < data.Rows.Count; ++i)
                {
                    if (data.Rows[i][j+1] == null || data.Rows[i][j+1].ToString().Length == 0)
                    {
                        continue;
                    }
                    string name = data.Rows[i][j].ToString();
                    int line = int.Parse(data.Rows[i][j+1].ToString());
                    Decimal amount = 0;
                    if(data.Rows[i][j+2]!=null){
                         Decimal.TryParse(data.Rows[i][j+2].ToString(),out amount);
                    }
                       
                    sb.Append("(");
                    sb.Append("'" + id + "',");
                    sb.Append("'" + name + "',");
                    sb.Append("'" + line + "',");
                    sb.Append("'" + amount + "'),");
                    //if (i == (data.Rows.Count - 1) && (j == maxCell - 1))
                    //{
                    //    sb.Append("'" + amount + "');");
                    //}
                    //else
                    //{
                    //    sb.Append("'" + amount + "'),");
                    //}
                }
                j += maxCell/2;
            }
            if (sb.Length > 0)
            {
                string sql = sb.ToString().TrimEnd(',') + ";";
                log(sql);
                insertBatchProfit(sql);
                sb.Length = 0;
            }
           }

        public static string checkCell(Object obj)
        {
            if (obj == null||obj.ToString()==string.Empty) return string.Empty;
            return obj.ToString();
        }
       public static void ImportAllProfit(string file)
        {
            try
            {
                using (ExcelHelper excelHelper = new ExcelHelper(file))
                {
                    DataSet dst = excelHelper.ExcelToDataSet(1);
                    foreach (DataTable dt in dst.Tables)
                    {
                        OperateProfitData(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                log(ex);
               // Console.WriteLine("Exception: " + ex.Message);
            }
        }

       public static void ImportAllBalance(string file)
       {
           try
           {
               using (ExcelHelper excelHelper = new ExcelHelper(file))
               {
                   DataSet dst = excelHelper.ExcelToDataSet(1);
                   foreach (DataTable dt in dst.Tables)
                   {
                       OperateBalanceData(dt);
                   }
               }
           }
           catch (Exception ex)
           {
               log(ex);
               //Console.WriteLine("Exception: " + ex.Message);
           }
       }

       public static void ImportAllCashFlow(string file)
       {
           try
           {
               using (ExcelHelper excelHelper = new ExcelHelper(file))
               {
                   DataSet dst = excelHelper.ExcelToDataSet(1);
                   foreach (DataTable dt in dst.Tables)
                   {
                       OperateCashFlowData(dt);
                   }
               }
           }
           catch (Exception ex)
           {
               log(ex);
              // Console.WriteLine("Exception: " + ex.Message);
           }
       }
    }
}

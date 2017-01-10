using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Data;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace ExcelTool
{
   public class MysqlHelper
    {
        private static readonly string connectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

        public static int ExecuteSql(string cmdText)
        {
            int ret = 0;
            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
                using (MySqlCommand command = new MySqlCommand(cmdText, conn))
                {
                    ret = command.ExecuteNonQuery();
                    conn.Close();
                }
            }
            return ret;
        }
    }
}

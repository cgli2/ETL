using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace ExcelTool
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainFrame());
        }
        //private static string saveFolder = @"D:\temp";
        //private static string file = @"D:\10.zip";
       // static void Main(string[] args)
      //  {
            //Console.WriteLine("Begin to unzip");
            //String tempFolder = file.Substring(file.LastIndexOf("\\") + 1);
            //tempFolder = saveFolder + "\\" + tempFolder.Remove(tempFolder.IndexOf("."));
            //Console.WriteLine("------------------->tempFolder:" + tempFolder);
            //LogHelper.Log(typeof(Program),"testing log...");

            //ZipHelper.UnZip(file, tempFolder);
            //Console.WriteLine("Un zip success !");
            //String excelFile = tempFolder + "\\利润表.xls";
            //ImportToData.ImportAllProfit(excelFile);
            //String jj = "";
            //Decimal dd = 0;
            //    Decimal.TryParse(jj,out dd);
            //Console.WriteLine("dd parse ok="+dd);

           //  ExcelHelper excel = new ExcelHelper(excelFile);
           ////  String excelFile = saveFolder + "\\利润表.xls";
           // DataSet ds = excel.ExcelToDataSet(true);
           // if (ds != null)
           // {
           //     foreach (DataTable dt in ds.Tables)
           //     {
           //         foreach (DataRow dr in dt.Rows)
           //         {
           //             for (int i = 0; i < dt.Columns.Count; i++)
           //             {
           //                 Console.WriteLine(dr[i]);
           //             }
           //         }
           //     }

           // }
            // TestExcelWrite(excelFile);
            // TestExcelRead(excelFile);
            // TestExcelReadAll(excelFile);
            //try
            //{
            //    JobManager.Initialize(new MyRegistry());
            //}
            //catch (Exception ex)
            //{
            //    LogHelper.Log(typeof(Program), ex);
            //}
            //Console.ReadKey();
        //}


        #region testing 
        //static DataTable GenerateData()
        //{
        //    DataTable data = new DataTable();
        //    for (int i = 0; i < 5; ++i)
        //    {
        //        data.Columns.Add("Columns_" + i.ToString(), typeof(string));
        //    }

        //    for (int i = 0; i < 10; ++i)
        //    {
        //        DataRow row = data.NewRow();
        //        row["Columns_0"] = "item0_" + i.ToString();
        //        row["Columns_1"] = "item1_" + i.ToString();
        //        row["Columns_2"] = "item2_" + i.ToString();
        //        row["Columns_3"] = "item3_" + i.ToString();
        //        row["Columns_4"] = "item4_" + i.ToString();
        //        data.Rows.Add(row);
        //    }
        //    return data;
        //}

        //static void PrintData(DataTable data)
        //{
        //    if (data == null) return;

        //    string taxpayer_id = data.Rows[0][1].ToString();
        //    string taxpayer_name = data.Rows[1][1].ToString();
        //    string certificate = data.Rows[2][1].ToString();
        //    string declare_date = data.Rows[3][1].ToString();
        //    string filling_date = data.Rows[4][1].ToString();
        //    string report_date = data.Rows[5][1].ToString();
        //    int id = ImportToData.saveBase(taxpayer_id, taxpayer_name, certificate, declare_date, filling_date, report_date);

        //    StringBuilder sb = new StringBuilder();
        //    for (int i = 0; i < data.Rows.Count; ++i)
        //    {
        //        if (i > 6)
        //        {
        //            string name = data.Rows[i][0].ToString();
        //            int line = int.Parse(data.Rows[i][1].ToString());
        //            Decimal yearAmount = Decimal.Parse(data.Rows[i][2].ToString());
        //            Decimal monthAmount = Decimal.Parse(data.Rows[i][3].ToString());
        //            // ImportToData.saveProfit(id, name, line, monthAmount, yearAmount);
        //            sb.Append("(");
        //            sb.Append("'" + name + "',");
        //            sb.Append("'" + line + "',");
        //            sb.Append("'" + yearAmount + "',");
        //            sb.Append("'" + monthAmount + "'),");

        //            //for (int j = 0; j < data.Columns.Count; ++j)
        //            //{
        //            //    Console.Write("{0} ", data.Rows[i][j]);
        //            //}
        //        }
        //        Console.Write("\n");
        //    }
        //}

        //static void TestExcelWrite(string file)
        //{
        //    try
        //    {
        //        using (ExcelHelper excelHelper = new ExcelHelper(file))
        //        {
        //            DataTable data = GenerateData();
        //            int count = excelHelper.DataTableToExcel(data, "MySheet", true);
        //            if (count > 0)
        //                Console.WriteLine("Number of imported data is {0} ", count);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Exception: " + ex.Message);
        //    }
        //}

        //static void TestExcelRead(string file)
        //{
        //    try
        //    {
        //        using (ExcelHelper excelHelper = new ExcelHelper(file))
        //        {
        //            DataTable dt = excelHelper.ExcelToDataTable("MySheet", 1);
        //            PrintData(dt);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Exception: " + ex.Message);
        //    }
        //}

        //static void TestExcelReadAll(string file)
        //{
        //    try
        //    {
        //        using (ExcelHelper excelHelper = new ExcelHelper(file))
        //        {
        //            DataSet dst = excelHelper.ExcelToDataSet(1);
        //            foreach (DataTable dt in dst.Tables)
        //            {
        //                // PrintData(dt);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Exception: " + ex.Message);
        //    }
        //}
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Configuration;

using FluentScheduler;

namespace ExcelTool
{
    class MyRegistry : Registry
    {
        private static string unZipTempFolder = @"D:\temp";
        private static string finishBackup = ConfigurationManager.AppSettings["DataBackupPath"];// @"D:\databackup";
        private static string path =  ConfigurationManager.AppSettings["ExcelDataPath"];//@"D:\www\cyshu\zt\declare\data";
        public MyRegistry()
        {
            //NonReentrantAsDefault();
           // Welcome();
           // NonReentrant();
            Reentrant();
            //OnceIn();
            OnceAt();
            //Sleepy();
            //Faulty();
            //Removed();
            //TenMinutes();
            //Hour();
            //Day();
            //Weekday();
            //Week();
        }

        //private void Welcome()
        //{
        //    Schedule(() => Console.Write("3, "))
        //        .WithName("[welcome]")
        //        .AndThen(() => Console.Write("2, "))
        //        .AndThen(() => Console.Write("1, "))
        //        .AndThen(() => Console.WriteLine("Live!"))
        //        .AndThen(() => Console.WriteLine("{0}You can check what's happening in the log file at \"{1}\"",
        //            Environment.NewLine, L.Directory));
        //}
        //private void NonReentrant()
        //{
        //  //  L.Register("[non reentrant]");

        //    Schedule(() =>
        //    {
        //        //L.Log("[non reentrant]", "Sleeping a minute...");
        //        Thread.Sleep(TimeSpan.FromMinutes(1));
        //    }).NonReentrant().WithName("[non reentrant]").ToRunEvery(1).Seconds();
        //}
        private static void log(string message)
        {
            LogHelper.Log(typeof(MyRegistry), message);
        }
        private void Reentrant()
        {
            log("---->Begin to reentrant the Job!");
            Schedule(() =>
            {
                string[] txtFiles = Directory.GetFiles(path, "*.zip", SearchOption.AllDirectories);
                string fileName = string.Empty;
                string tempFolder = string.Empty;
                foreach (string file in txtFiles)
                {
                    fileName = file.Substring(file.LastIndexOf("\\") + 1);
                    tempFolder = unZipTempFolder + "\\" + fileName.Remove(fileName.IndexOf("."));
                    log("------------------->tempFolder:" + tempFolder);
                    ZipHelper.UnZip(file, tempFolder);
                    //Console.WriteLine("Un zip success !");

                    String excelFile1 = tempFolder + "\\利润表.xls";
                    String excelFile2 = tempFolder + "\\现金流量表.xls";
                    String excelFile3 = tempFolder + "\\资产表.xls";

                    ImportToData.ImportAllProfit(excelFile1);
                    ImportToData.ImportAllBalance(excelFile2);
                    ImportToData.ImportAllCashFlow(excelFile3);
                    File.Move(file, finishBackup + fileName);
                }
                // Thread.Sleep(TimeSpan.FromSeconds(3));
            }).WithName("[reentrant]").ToRunEvery(30).Minutes();//.ToRunNow().AndEvery(30).Minutes();
        }

        //private void OnceIn()
        //{
        //    L.Register("[once in]");

        //    Schedule(() =>
        //    {
        //        JobManager.RemoveJob("[reentrant]");
        //        JobManager.RemoveJob("[non reentrant]");
        //        L.Log("[once in]", "Disabled the reentrant and non reentrant jobs.");
        //    }).WithName("[once in]").ToRunOnceIn(3).Minutes();
        //}

        private void OnceAt()
        {
            log("----------->[once at]");

            Schedule(
                () =>{
                    //JobManager.RemoveJob("[reentrant]");
                    log("[once at] It's almost midnight.");
                    Directory.Delete(unZipTempFolder,true);
                   // JobManager..AddJob("[reentrant]");
                }
                
                ).WithName("[once at]").ToRunOnceAt(23, 50);
        }

        //private void Sleepy()
        //{
        //    L.Register("[sleepy]");

        //    Schedule(() =>
        //    {
        //        L.Log("[sleepy]", "Sleeping...");
        //        Thread.Sleep(new TimeSpan(0, 7, 30));
        //    }).WithName("[sleepy]").ToRunEvery(15).Minutes();
        //}

        //private void Faulty()
        //{
        //    L.Register("[faulty]");

        //    Schedule(() =>
        //    {
        //        L.Register("[faulty]", "I'm going to raise an exception!");
        //        throw new Exception("I warned you.");
        //    }).WithName("[faulty]").ToRunEvery(20).Minutes();
        //}

        //private void Removed()
        //{
        //    L.Register("[removed]");

        //    Schedule(() =>
        //    {
        //        L.Register("[removed]", "SOMETHING WENT WRONG.");
        //    }).WithName("[removed]").ToRunOnceAt(0, 2);
        //}

        //private void TenMinutes()
        //{
        //    L.Register("[ten minutes]");

        //    Schedule(() => L.Log("[ten minutes]", "Ten minutes has passed."))
        //        .WithName("[ten minutes]").ToRunEvery(10).Minutes();
        //}

        //private void Hour()
        //{
        //    L.Register("[hour]");

        //    Schedule(() => L.Log("[hour]", "A hour has passed."))
        //        .WithName("[hour]").ToRunEvery(1).Hours();
        //}

        //private void Day()
        //{
        //    L.Register("[day]");

        //    Schedule(() => L.Log("[day]", "A day has passed."))
        //        .WithName("[day]").ToRunEvery(1).Days();
        //}

        //private void Weekday()
        //{
        //    L.Register("[weekday]");

        //    Schedule(() => L.Log("[weekday]", "A new weekday has started."))
        //        .WithName("[weekday]").ToRunEvery(1).Weekdays();
        //}

        //private void Week()
        //{
        //    L.Register("[week]");

        //    Schedule(() => L.Log("[week]", "A new week has started."))
        //        .WithName("[week]").ToRunEvery(1).Weeks();
        //}
    }
}

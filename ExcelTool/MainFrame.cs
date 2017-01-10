using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FluentScheduler;

namespace ExcelTool
{
    public partial class MainFrame : Form
    {
        private NotifyIcon notifyIcon = null;  
        public MainFrame()
        {
            InitializeComponent();
            InitialTray();
        }

        private void MainFrame_Load(object sender, EventArgs e)
        {
            try
            {
                JobManager.Initialize(new MyRegistry());
            }
            catch (Exception ex)
            {
                LogHelper.Log(typeof(Program), ex);
            }

            this.ShowInTaskbar = false;
            this.Hide();
        }

        private void MainFrame_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            //通过这里可以看出，这里的关闭其实不是真正意义上的“关闭”，而是将窗体隐藏，实现一个“伪关闭”  
            this.Hide();  
        }
        private void InitialTray()
        {
            //隐藏主窗体  
            this.Hide();
            //实例化一个NotifyIcon对象  
            notifyIcon = new NotifyIcon();
            //托盘图标气泡显示的内容  
            notifyIcon.BalloonTipText = "正在后台运行";
            //托盘图标显示的内容  
            notifyIcon.Text = "ETL后台监控程序";
            notifyIcon.Icon = new System.Drawing.Icon("small.ico");
            //true表示在托盘区可见，false表示在托盘区不可见  
            notifyIcon.Visible = true;
            //气泡显示的时间（单位是毫秒）  
            notifyIcon.ShowBalloonTip(2000);
            notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(notifyIcon_MouseClick);
            ////设置二级菜单  
            //MenuItem setting1 = new MenuItem("二级菜单1");  
            //MenuItem setting2 = new MenuItem("二级菜单2");  
            //MenuItem setting = new MenuItem("一级菜单", new MenuItem[]{setting1,setting2});  
            //帮助选项，这里只是“有名无实”在菜单上只是显示，单击没有效果，可以参照下面的“退出菜单”实现单击事件  

            MenuItem help = new MenuItem("帮助");
            //关于选项  
            MenuItem about = new MenuItem("关于");
            //退出菜单项  
            MenuItem exit = new MenuItem("退出");
            exit.Click += new EventHandler(exit_Click);
            ////关联托盘控件  
            //注释的这一行与下一行的区别就是参数不同，setting这个参数是为了实现二级菜单  
            //MenuItem[] childen = new MenuItem[] { setting, help, about, exit };  
            MenuItem[] childen = new MenuItem[] { help, about, exit };
            notifyIcon.ContextMenu = new ContextMenu(childen);
            //窗体关闭时触发  
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFrame_FormClosing);

        }



        /// <summary>  
        /// 鼠标单击  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void notifyIcon_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                //如果窗体是可见的，那么鼠标左击托盘区图标后，窗体为不可见  
                if (this.Visible == true)
                {
                    this.Visible = false;
                }
                else
                {
                    this.Visible = true;
                    this.Activate();
                }

            }

        }



        /// <summary>  
        /// 退出选项  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void exit_Click(object sender, EventArgs e)
        {
            //退出程序  
            System.Environment.Exit(0);

        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            String excelFile = @"D:\project\hongxian\厦门14-4-12.xls";
            ImportToData.ImportAllCompany(excelFile);//.ImportAllProfit(excelFile);
            MessageBox.Show("完成！");
        }  
    }
}

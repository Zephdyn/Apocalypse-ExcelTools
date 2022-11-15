using excelTools.data;
using excelTools.WPF;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelTools.window
{
    /// <summary>
    /// Main.xaml 的交互逻辑
    /// </summary>
    public partial class Main : Window
    {
        // 数据绑定
        private static readonly WpfInfo WIfon = new();

        // 数据库路径
        private static string DatabasePath;
        // 软件支持的数据库版本
        private static readonly int DatabaseVersion = 3030001;
        // 当前数据库版本
        private static int DatabaseReadVersion;
        public Main()
        {
            // 窗口居中显示
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            // 初始化
            InitializeComponent();
            // 数据绑定
            this.DataContext = WIfon;

            // 设置窗口名称
            this.Title = @"天启幻境II - Excel Tools";
            WIfon.TabControl = 0;
            // 获取版本
            WIfon.Version = Application.ResourceAssembly.GetName().Version.ToString();

            // 数据初始化            
            string configPath = "config.db";
            string filePath = "config.json";
            if (File.Exists(filePath))
            {
                //StreamReader file = File.OpenText(filePath);
                string file = File.ReadAllText(filePath);
                try
                {
                    JObject json = JObject.Parse(file);
                    // 读取数据库路径
                    configPath = json["db_path"]["config"].ToString();
                }
                catch
                {
                    // 无json或json错误
                }
            }
            GetConfigInfo(configPath);
        }
        #region 窗口 事件
        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                this.WindowState = this.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            }
        }
        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        private void 顶层_缩小_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        private void 顶层_最大_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = this.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }
        private void 顶层_关闭_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        #endregion
        /// <summary>
        /// 获取数据库信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetConfigInfo(string filePath)
        {
            if (!File.Exists(filePath))// 路径不存在
            {
                return false;
            }
            File.Copy(filePath, filePath + ".back", true);
            if (Config.Open(filePath) == false)// 未打开数据库
            {
                WIfon.DbConfigInfo = "当前数据库状态：连接失败";
                return false;
            }
            string[] ver = Config.Ver();
            if (ver[2] != null)
            {
                Config.Close();// 关闭数据库连接
                MessageBox.Show(ver[2], "读取数据库错误");
                return false;
            }
            WIfon.DbConfigInfo = $"当前数据库信息：{ver[0]}({ver[1]})";
            DatabasePath = filePath;
            if (!int.TryParse(ver[1], out DatabaseReadVersion)) DatabaseReadVersion = 0;      
            // 读取版本低于 软件版本
            if (DatabaseReadVersion < DatabaseVersion)
            {
                MessageBox.Show($"当前数据库版本低于本工具, 无法使用本工具, 请将数据库版本升级后使用{System.Environment.NewLine}当前数据库版本：{DatabaseReadVersion} {System.Environment.NewLine} 当前工具支持版本：{DatabaseVersion}", "警告");
                return false;
            }
            // 读取版本超过 软件版本
            if (DatabaseReadVersion > DatabaseVersion)
                if (MessageBox.Show(
                    $"当前数据库版本大于本工具, 推荐使用与数据库版本匹配的工具{System.Environment.NewLine}否则可能出现导入数据后异常的情况{System.Environment.NewLine}当前数据库版本：{DatabaseReadVersion} {System.Environment.NewLine} 当前工具支持版本：{DatabaseVersion}{System.Environment.NewLine} 是否要继续使用?",
                    "警告",
                    MessageBoxButton.OKCancel) != MessageBoxResult.OK)
                    return false;
            return true;
        }

    }
}

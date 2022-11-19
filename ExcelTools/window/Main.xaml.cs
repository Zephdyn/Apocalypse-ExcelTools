using excelTools.data;
using excelTools.excel;
using excelTools.WPF;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
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

        // 打开保存文件
        readonly OpenFileDialog Odialog = new();
        readonly SaveFileDialog Sdialog = new();
        // 写入类型
        private static int WriteType = 0;
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

            Debug.WriteLine("数据库已连接", filePath);
            return true;
        }


        private void 功能_栏_SelectionChanged(object sender, SelectionChangedEventArgs e) // 选择夹选中
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var data = this.功能_栏.SelectedItem as TabItem;
                var datas = data.Header; //主要是在后端获取到当前的TabItem的Heade 
                Debug.WriteLine("content:" + this.功能_栏.SelectedContent + "; index = " + this.功能_栏.SelectedIndex + "; datas" + datas);

            }));
        }

        private void Button_Excel_Click(object sender, RoutedEventArgs e) // 数据导入
        {
            Button bt = sender as Button;
            int command = Int32.Parse(bt.CommandParameter.ToString());
            string content = bt.Content.ToString();
            Debug.WriteLine(content);
            Debug.WriteLine(Config.Status());
            if (Config.Status() != (int)System.Data.ConnectionState.Open)
            {
                if (content == "数据导入")
                {
                    MessageBox.Show("还未连接数据库");
                    return;
                }
            }
            string[] Rtype = Array.Empty<string>();
            string name = "";
            byte[] template = Array.Empty<byte>();
            switch (command)
            {
                case 1:
                    Rtype = ExcelTypeRead_v1.equip;
                    name = "装备示例.xlsx";
                    template = Properties.Resources.装备示例;
                    break;
                case 2:
                    Rtype = ExcelTypeRead_v1.map;
                    name = "地图示例.xlsx";
                    template = Properties.Resources.地图示例;
                    break;
                case 3:
                    Rtype = ExcelTypeRead_v1.monster;
                    name = "怪物示例.xlsx";
                    template = Properties.Resources.怪物示例;
                    break;
                case 4:
                    Rtype = ExcelTypeRead_v1.drop;
                    name = "掉落示例.xlsx";
                    template = Properties.Resources.掉落示例;
                    break;
                case 5:
                    Rtype = ExcelTypeRead_v1.goods;
                    name = "物品示例.xlsx";
                    template = Properties.Resources.物品示例;
                    break;
                case 6:
                    Rtype = ExcelTypeRead_v1.shop;
                    name = "商店示例.xlsx";
                    template = Properties.Resources.商店示例;
                    break;
                case 7:
                    Rtype = ExcelTypeRead_v1.duties;
                    name = "任务示例.xlsx";
                    template = Properties.Resources.任务示例;
                    break;
                case 8:
                    Rtype = ExcelTypeRead_v1.copy;
                    name = "副本示例.xlsx";
                    template = Properties.Resources.副本示例;
                    break;

            }
            if (content=="数据导入")
            {
                // 选择excel
                Odialog.Filter = "Xlsx|*.xlsx|Xls|*.xls";
                if (Odialog.ShowDialog(this) == false) return;
                var _fileName = Odialog.FileName;
                // 进行数据读取
                WriteType = command;
                //Tab_2_Get(1, Rtype, _fileName, 0);
            }
            else
            {
                // 导出excel示例
                Sdialog.FileName = name;
                Sdialog.Filter = "Excel(*xlsx)|*.xlsx";
                if (Sdialog.ShowDialog(this) == false) return;
                var _fileName = Sdialog.FileName;

                //byte[] template = Properties.Resources.装备示例;
                File.WriteAllBytes(_fileName, template);
                MessageBox.Show("已写出示例文件");
            }
        }
    }
}

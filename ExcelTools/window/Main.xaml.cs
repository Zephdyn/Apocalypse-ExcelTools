using excelTools.data;
using excelTools.excel;
using excelTools.WPF;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Schema;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
using System.Xml.Linq;

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
        // 软件支持的最低数据库版本(结构要求: 低于此版本没有 4.0.0 新增的自定义属性表)
        private const int DatabaseVersionMin = 4000000;
        // 软件兼容的最高数据库版本(4000001 只新增了 group_notice, 不影响本工具)
        private const int DatabaseVersionMax = 4000001;
        // 当前数据库版本
        private static int DatabaseReadVersion;
        // 数据库是否已通过 版本 + 结构 校验(通过后才允许导入)
        private static bool DbReady;

        // 打开保存文件
        readonly OpenFileDialog Odialog = new();
        readonly SaveFileDialog Sdialog = new();
        // 写入类型
        private static int WriteType = 0;

        // 配置文件
        private static JObject Infojson = [];
        public Main()
        {
            // 窗口居中显示
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            // 初始化
            InitializeComponent();
            // 数据绑定
            this.DataContext = WIfon;

            // 设置窗口名称
            this.Title = "天启幻境II - Excel Tools";
            WIfon.TabControl = 0;
            // 获取版本
            WIfon.Version = Application.ResourceAssembly.GetName().Version.ToString()[..^2];
            WIfon.DbSupportInfo = $"支持：4.0.0 - 4.0.1 ({DatabaseVersionMin} - {DatabaseVersionMax})";
            // 数据初始化            
            string configPath = "config.db";
            string filePath = "config.json";
            Debug.WriteLine("config = " + filePath);
            if (File.Exists(filePath))
            {
                //StreamReader file = File.OpenText(filePath);
                string file = File.ReadAllText(filePath);
                try
                {
                    Infojson = JObject.Parse(file);

                    if (Infojson.Property("db") != null) configPath = Infojson["db"]["path"].ToString();// 读取数据库路径
                    if (Infojson.Property("window") != null) WIfon.TabControl = int.Parse(Infojson["window"]["tab"].ToString());// 记录上次保存的
                }
                catch
                {
                    // 无json或json错误
                }
            }
            GetConfigInfo(configPath);
        }
        #region 窗口 基础事件
        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
            }
        }
        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
        private void 顶层_缩小_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
        private void 顶层_最大_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }
        private void 顶层_关闭_Click(object sender, RoutedEventArgs e)
        {
            // 保存上次界面
            if (Infojson.Property("window") == null)
                Infojson.Add(new JProperty("window", new JObject(new JProperty("tab"))));
            Infojson["window"]["tab"] = 功能_栏.SelectedIndex;
            File.WriteAllText("config.json", Infojson.ToString());
            Close();
        }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Key key = (e.Key == Key.System ? e.SystemKey : e.Key);
            if (key == Key.Tab || key == Key.LeftAlt || key == Key.RightAlt)
            {
                e.Handled = true;
            }
        }
        private void 功能_栏_SelectionChanged(object sender, SelectionChangedEventArgs e) // 选择夹选中
        {
            //Dispatcher.BeginInvoke(new Action(() =>
            //{
            //    var data = this.功能_栏.SelectedItem as TabItem;
            //    var datas = data.Header; //主要是在后端获取到当前的TabItem的Heade 
            //    Debug.WriteLine("content:" + this.功能_栏.SelectedContent + "; index = " + this.功能_栏.SelectedIndex + "; datas" + datas);

            //}));
        }

        #endregion

        #region 窗口 功能事件
        private void Button_Click_Excel(object sender, RoutedEventArgs e) // 数据导入
        {
            Button bt = sender as Button;
            int command = Int32.Parse(bt.CommandParameter.ToString());
            string content = bt.Content.ToString();
            Debug.WriteLine(content);
            Debug.WriteLine(Config.Status());
            // 数据导入必须先通过数据库的版本与结构校验, 否则会写坏数据或导到一半报错
            if (content == "数据导入" && (!DbReady || Config.Status() != (int)ConnectionState.Open))
            {
                MessageBox.Show("还未连接数据库，或数据库未通过校验，请先点击「打开数据库」选择 config.db", "提示");
                return;
            }
            string[] Rtype = [];
            string name = "";
            byte[] template = [];
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
                case 9:
                    Rtype = ExcelTypeRead_v1.skill;
                    name = "技能示例.xlsx";
                    template = Properties.Resources.技能示例;
                    break;

            }
            if (content == "数据导入")
            {
                // 选择excel
                Odialog.Filter = "Xlsx|*.xlsx|Xls|*.xls";
                if (Odialog.ShowDialog(this) == false) return;
                // 创建新窗口
                var _fileName = Odialog.FileName;
                WriteType = command;
                Window_Excel windows = new(WriteType, 1, Rtype, _fileName, 0);
                windows.ShowDialog();
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
                MessageBox.Show("已写出示例文件", "提示");
            }
        }

        private void Button_Click_OpenDB(object sender, RoutedEventArgs e) // 打开数据库
        {
            Odialog.Filter = "config(.db)|*.db";
            if (Odialog.ShowDialog(this) == false) return;
            var _fileName = Odialog.FileName;
            // config.db
            if (_fileName.Substring(_fileName.Length - 9, 9) != "config.db")
            {
                MessageBox.Show("请选择名为config.db的数据库", "打开数据库失败");
                return;
            }
            if (_fileName == DatabasePath)
            {
                MessageBox.Show("当前已是此数据库");
                return;
            };
            // 判断数据库是否已打开
            int state = (int)Config.Status();
            if (state == (int)System.Data.ConnectionState.Open) Config.Close();// 关闭已打开的数据库

            // 打开新数据库
            if (!GetConfigInfo(_fileName)) return;

            // Json生成
            if (Infojson.Property("db") == null)
                Infojson.Add(new JProperty("db", new JObject(new JProperty("path"))));
            Infojson["db"]["path"] = _fileName;
            // 写到新文件
            File.WriteAllText("config.json", Infojson.ToString());
            MessageBox.Show("数据库连接成功!");
        }// 数据库切换
        #endregion
        /// <summary>
        /// 获取数据库信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetConfigInfo(string filePath)
        {
            DbReady = false;
            WIfon.DbConfigInfo = "未连接数据库";
            WIfon.DbConfigPath = filePath;
            WIfon.DbConnectButtonBackground = "#F44336";
            if (!File.Exists(filePath))// 路径不存在
            {
                return false;
            }
            try
            {
                // 打开前备份一份, 备份失败(例如被服务端占用)不阻断使用
                File.Copy(filePath, filePath + ".back", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("备份数据库失败: " + ex.Message);
            }
            if (Config.Open(filePath) == false)// 未打开数据库
            {
                WIfon.DbConfigInfo = "当前数据库状态：连接失败";
                return false;
            }
            string[] ver = Config.Ver();
            if (ver[2] != null)
            {
                Config.Close();// 关闭数据库连接
                WIfon.DbConfigInfo = "当前数据库状态：读取错误";
                //MessageBox.Show(ver[2], "读取数据库错误");
                return false;
            }
            if (!int.TryParse(ver[1], out DatabaseReadVersion)) DatabaseReadVersion = 0;
            // 读取版本低于 软件要求
            if (DatabaseReadVersion < DatabaseVersionMin)
            {
                // 必须先断开连接: 否则界面仍显示已连接, 用户还能继续往旧库里导入
                Config.Close();
                WIfon.DbConfigInfo = $"当前数据库版本过低：{ver[0]}({ver[1]})";
                MessageBox.Show(
                    $"当前数据库版本低于本工具要求, 无法使用本工具, 请将数据库版本升级后使用{System.Environment.NewLine}" +
                    $"当前数据库版本：{DatabaseReadVersion}{System.Environment.NewLine}" +
                    $"本工具要求版本：{DatabaseVersionMin} 及以上{System.Environment.NewLine}" +
                    $"当前数据库：{filePath}",
                    "警告");
                return false;
            }
            // 读取版本超过 软件兼容
            if (DatabaseReadVersion > DatabaseVersionMax)
            {
                if (MessageBox.Show(
                    $"当前数据库版本高于本工具兼容范围, 推荐使用与数据库版本匹配的工具{System.Environment.NewLine}否则可能出现导入数据后异常的情况{System.Environment.NewLine}" +
                    $"当前数据库版本：{DatabaseReadVersion}{System.Environment.NewLine}" +
                    $"本工具兼容版本：{DatabaseVersionMin} - {DatabaseVersionMax}{System.Environment.NewLine}" +
                    $"当前数据库：{filePath}{System.Environment.NewLine} 是否要继续使用?",
                    "警告",
                    MessageBoxButton.OKCancel) != MessageBoxResult.OK)
                {
                    Config.Close();
                    return false;
                }
                // 显示黄色
                WIfon.DbConnectButtonBackground = "#FF9800";
            }
            else
            {
                // 显示绿色
                WIfon.DbConnectButtonBackground = "#4CAF50";
            }
            // 4.0.0 结构检查: 缺少自定义属性等表时, 装备/怪物导入会在第 1 条数据就失败
            List<string> missing = DbSchema.Check();
            if (missing.Count > 0)
            {
                MessageBoxResult result = MessageBox.Show(
                    $"当前数据库缺少 4.0.0 新增的表：{string.Join("、", missing)}{System.Environment.NewLine}" +
                    "缺少时装备/怪物的自定义属性无法写入。" + System.Environment.NewLine + System.Environment.NewLine +
                    "是否按 4.0.0 结构自动创建缺失的表?" + System.Environment.NewLine +
                    $"当前数据库：{filePath}",
                    "数据库结构不完整",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (result != MessageBoxResult.Yes)
                {
                    Config.Close();
                    WIfon.DbConfigInfo = "当前数据库状态：结构不完整";
                    return false;
                }
                if (!DbSchema.CreateMissingTables(out string error))
                {
                    Config.Close();
                    WIfon.DbConfigInfo = "当前数据库状态：结构不完整";
                    MessageBox.Show("创建缺失的表失败：" + System.Environment.NewLine + error, "错误");
                    return false;
                }
            }
            WIfon.DbConfigInfo = $"当前版本：{ver[0]}({ver[1]})";
            DatabasePath = filePath;
            DbReady = true;
            return true;
        }




    }
}

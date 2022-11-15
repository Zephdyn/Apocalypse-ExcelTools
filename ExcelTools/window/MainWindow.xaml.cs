using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

using System.Threading;

using Newtonsoft.Json.Linq;
using Microsoft.Win32;

using excelTools.WPF;
using excelTools.data;
using excelTools.excel;
using System.Windows.Input;

namespace ExcelTools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 软件支持的数据库版本
        private static readonly int DbVersion = 3030001;

        public delegate void UpdateUI(int step);
        public int Readint = 0;

        private static int DbReadVersion;
        private static readonly WpfInfo WIfon = new();

        // 记录
        private static int GetDataVersion, GetDataSheetId;
        private static string[] GetDataReadType;
        private static string GetDataFileName;

        private static string DbConfigPath = "";
        private static int WriteType = 0;
        readonly OpenFileDialog Odialog = new();
        readonly SaveFileDialog Sdialog = new();
        public MainWindow()
        {
            InitializeComponent();
            // 数据初始化            
            string configPath = "config.db";
            WIfon.Version = "Ver：1.3.1";
            WIfon.DbConfigInfo = "未连接数据库, 请先选择要连接的数据库";
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
            if (File.Exists(configPath))
            {
                GetConfigInfo(configPath);
                DbConfigPath = configPath;
            }
            // Visibility="Collapsed"

            // 设置窗口名称
            this.Title = @"天启幻境II - Excel Tools";
            // 隐藏表头
            this.TabItem_1.Visibility = Visibility.Collapsed;
            this.TabItem_2.Visibility = Visibility.Collapsed;
            this.TabItem_3.Visibility = Visibility.Collapsed;
            // 显示软件版本
            this.DataContext = WIfon;
            // 注册可否被关闭事件
            this.Closing += WpfClosing;
            // 窗口居中显示
            WindowStartupLocation = WindowStartupLocation.CenterScreen;

            var t1 = new Task(() => Test());
            t1.Start();
        }
        private void WpfClosing(object o, System.ComponentModel.CancelEventArgs e)
        {
            if (Config.Status() == (int)System.Data.ConnectionState.Open) Config.Close(); // 如果数据库已连接,则关闭数据库
            e.Cancel = false;
        }
        private void Test()
        {
            string temp;
            while (true)
            {
                if (Readint == 2)
                {

                    switch (excelTools.data.V1Write.state)
                    {
                        case 1:
                            temp = "正在写入数据:" + excelTools.data.V1Write.Numwrite.ToString() + "/" + excelTools.excel.Program.arrWrite.Count.ToString();
                            this.Dispatcher.BeginInvoke(new Action(() => WIfon.TabWriteDataNum = temp));
                            break;
                        case 2:
                            temp = "正在清理数据:" + excelTools.data.V1Write.Numwrite.ToString() + "/" + excelTools.excel.Program.arrWrite.Count.ToString();
                            this.Dispatcher.BeginInvoke(new Action(() => WIfon.TabWriteDataNum = temp));
                            break;
                    }
                }
                Thread.Sleep(100);
            }
        }
        /// <summary>
        /// [按钮] 切换数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_OpenDb(object sender, RoutedEventArgs e)
        {
            if (Readint == 1)
            {
                MessageBox.Show("请等待数据读取完毕");
                return;
            }
            if (Readint == 2)
            {
                MessageBox.Show("正在写入数据,请耐心等待");
                return;
            }
            Odialog.Filter = "config(.db)|*.db";
            if (Odialog.ShowDialog(this) == false) return;
            var _fileName = Odialog.FileName;
            // config.db
            if (_fileName.Substring(_fileName.Length - 9, 9) != "config.db")
            {
                MessageBox.Show("请选择名为config.db的数据库", "打开数据库失败");
                return;
            }
            if (_fileName == DbConfigPath)
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
            JObject rss =
                new(new JProperty("db_path", new JObject(new JProperty("config", _fileName))));

            // 写到新文件
            File.WriteAllText("config.json", rss.ToString());
            MessageBox.Show("数据库连接成功!");
        }
        /// <summary>
        /// [按钮]切换到 选择夹1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Tab_1(object sender, RoutedEventArgs e)
        {
            if (Readint == 1)
            {
                MessageBox.Show("请等待数据读取完毕");
                return;
            }
            if (Readint == 2)
            {
                MessageBox.Show("请等待数据写入完毕");
                return;
            }
            WIfon.TabControl = 0;
        }
        /// <summary>
        /// [按钮]切换到 选择夹2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Tab_2(object sender, RoutedEventArgs e)
        {
            if (Config.Status() != (int)System.Data.ConnectionState.Open)
            {
                MessageBox.Show("请先进行数据库连接");
                return;
            }
            // 读取版本低于 软件版本
            if (DbReadVersion < DbVersion)
            {
                MessageBox.Show($"当前数据库版本低于本工具, 无法使用本工具, 请将数据库版本升级后使用{System.Environment.NewLine}当前数据库版本：{DbReadVersion} {System.Environment.NewLine} 当前工具支持版本：{DbVersion}", "警告");
                return;
            }
            // 读取版本超过 软件版本
            if (DbReadVersion > DbVersion)
                if (MessageBox.Show(
                    $"当前数据库版本大于本工具, 推荐使用与数据库版本匹配的工具{System.Environment.NewLine}否则可能出现导入数据后异常的情况{System.Environment.NewLine}当前数据库版本：{DbReadVersion} {System.Environment.NewLine} 当前工具支持版本：{DbVersion}{System.Environment.NewLine} 是否要继续使用?",
                    "警告",
                    MessageBoxButton.OKCancel) != MessageBoxResult.OK)
                    return;
            WIfon.TabControl = 1;
        }
        private void Button_Tab_2_equip(object sender, RoutedEventArgs e)
        {
            WindowExcelData(1);
        }
        private void Button_Tab_2_map(object sender, RoutedEventArgs e)
        {
            WindowExcelData(2);
        }
        private void Button_Tab_2_monster(object sender, RoutedEventArgs e)
        {
            WindowExcelData(3);
        }
        private void Button_Tab_2_drop(object sender, RoutedEventArgs e)
        {
            WindowExcelData(4);
        }
        private void Button_Tab_2_goods(object sender, RoutedEventArgs e)
        {
            WindowExcelData(5);
        }
        private void Button_Tab_2_shop(object sender, RoutedEventArgs e)
        {
            WindowExcelData(6);
        }

        private void Button_Tab_2_duties(object sender, RoutedEventArgs e)
        {
            WindowExcelData(7);
        }
        private void Button_Tab_2_copyOne(object sender, RoutedEventArgs e)
        {
            WindowExcelData(8);
        }
        /// <summary>
        /// [按钮] 写入数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Tab_3_Writedata(object sender, RoutedEventArgs e)
        {
            if (Readint == 1)
            {
                MessageBox.Show("请等待数据读取完毕");
                return;
            }
            if (Readint == 2)
            {
                MessageBox.Show("正在写入数据,请耐心等待");
                return;
            }
            if (excelTools.excel.Program.arrWrite.Count == 0)
            {
                MessageBox.Show("无数据可写入");
                return;
            }
            if (MessageBox.Show("提示：在导入数据前请先备份数据库\r\n你是否要开始导入", "提示", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                return;
            Readint = 2;
            try
            {
                var t1 = new Task(() => V1Write.Equip());
                switch (WriteType)
                {
                    case 0: // 未读取数据
                        MessageBox.Show("请先进行excel读取");
                        return;
                    case 1:
                        t1 = new Task(() => V1Write.Equip());
                        break;
                    case 2:
                        t1 = new Task(() => V1Write.Map());
                        break;
                    case 3:
                        t1 = new Task(() => V1Write.Monster());
                        break;
                    case 4:
                        t1 = new Task(() => V1Write.Drop());
                        break;
                    case 5:
                        t1 = new Task(() => V1Write.Goods());
                        break;
                    case 6:
                        t1 = new Task(() => V1Write.Shop());
                        break;
                    case 7:
                        t1 = new Task(() => V1Write.Duties());
                        break;
                    case 8:
                        t1 = new Task(() => V1Write.Copy());
                        break;
                    default:
                        MessageBox.Show("异常类型");
                        return;
                }
                t1.Start();
                t1.ContinueWith((task) =>
                {
                    // 委托
                    Dispatcher.Invoke(
                    new Action(
                            delegate
                            {
                                Readint = 0;
                                WIfon.TabWriteDataNum = "";
                            }
                        ));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误");
                return;
            }

        }


        public void WindowExcelData(int type)
        {
            string[] Rtype = Array.Empty<string>();
            string name = "";
            byte[] template = Array.Empty<byte>();
            switch (type)
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
            if ((bool)RadioTab2Mode.IsChecked)
            {
                // 选择excel
                Odialog.Filter = "xlsx(推荐|*.xlsx|xls(不推荐|*.xls";
                if (Odialog.ShowDialog(this) == false) return;
                var _fileName = Odialog.FileName;
                // 进行数据读取
                WriteType = type;
                Tab_2_Get(1, Rtype, _fileName, 0);
            }
            else if ((bool)RadioTab2Mode_2.IsChecked)
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
            else
            {
                MessageBox.Show("这是个异次元");
            }
        }
        /// <summary>
        /// 获取数据库信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetConfigInfo(string filePath)
        {
            File.Copy(filePath, filePath + ".back", true);
            if (Config.Open(filePath) == false)// 未打开数据库
            {
                WIfon.DbConfigInfo = "当前数据库状态：连接失败";
            }
            else
            {
                string[] ver = Config.Ver();
                if (ver[2] != null)
                {
                    Config.Close();// 关闭数据库连接
                    MessageBox.Show(ver[2], "读取数据库错误");
                }
                else
                {
                    WIfon.DbConfigInfo = $"当前数据库信息：{ver[0]}({ver[1]})";
                    if (!int.TryParse(ver[1], out DbReadVersion)) DbReadVersion = 0;
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 传入要打开的excel,进行读取
        /// </summary>
        /// <param name="version"></param>
        /// <param name="sutype"></param>
        /// <param name="_fileName"></param>
        /// <param name="sheetId"></param>
        /// <returns></returns>
        private void Tab_2_Get(int version, string[] readType, string _fileName, int sheetId)
        {
            if (!File.Exists(_fileName)) return;
            int _excelType = _fileName[^5..] == ".xlsx" ? 2 : 1;
            // 记录耗时
            Int64 startTime = Stopwatch.GetTimestamp();
            // 线程启动读取
            if (Readint != 0)
            {
                MessageBox.Show("正在读取, 请耐心等待读取完毕");
                return;
            }
            Readint = 1;
            GetDataVersion = version;
            GetDataReadType = readType;
            GetDataFileName = _fileName;
            GetDataSheetId = sheetId;
            Task t = new(() =>
           {
               TrdRead info = new()
               {
                   version = version,
                   type = _excelType,
                   sutype = readType,
                   fileName = _fileName,
                   sheetId = sheetId
               };
               info.Excel();
           });
            t.Start();

            this.Tab_3_DataGrid.Columns.Clear();
            WIfon.TabControl = 2;
            this.Tab_3_TextBlock.Text = $"正在读取数据,请耐心等待";
            t.ContinueWith((task) =>
            {
                // 委托
                Dispatcher.Invoke(
                new Action(
                        delegate
                        {
                            Readint = 0;
                            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
                            DataGridHs(1, readType);
                            this.Tab_3_TextBlock.Text = $"当前数据数量：{excelTools.excel.Program.arrWrite.Count}, 耗时：{stime} s";
                        }
                    ));
            });
            return;
        }

        /// <summary>
        /// 将已读取的数据显示到页面
        /// </summary>
        /// <param name="page"></param>
        /// <param name="sArr"></param>
        public void DataGridHs(int page, string[] sArr)
        {
            this.Tab_3_DataGrid.Columns.Clear();
            this.Tab_3_DataGrid.Columns.Add(new DataGridTextColumn() { Header = "ID", Binding = new Binding(string.Format("[{0}]", 0)) });
            for (var i = 0; i < sArr.Length; i++)
            {
                this.Tab_3_DataGrid.Columns.Add(new DataGridTextColumn() { Header = sArr[i], Binding = new Binding(string.Format("[{0}]", i + 1)) });
            }
            int pageSize = 50;
            List<string[]> list = excelTools.excel.Program.arrWrite;
            list = list.Skip((page - 1) * pageSize).Take(pageSize).ToList();
            this.Tab_3_DataGrid.ItemsSource = list;
            this.Tab_3_DataGrid.IsReadOnly = true;

            List<object[]> oList = excelTools.excel.Program.arrShee;
            List<WpfInfo.TabShee> shee = new();
            foreach (var item in oList)
            {
                shee.Add(new WpfInfo.TabShee { ID = (int)item.GetValue(0), Name = (string)item.GetValue(1) });
            }
            Tab_3_ComboBox_Shhet.ItemsSource = shee;
            Tab_3_ComboBox_Shhet.SelectedIndex = 0;
        }

        private void Button_Tab_3_brake(object sender, RoutedEventArgs e)
        {
            WIfon.TabControl = 1;
        }

        private void Button_Tab_3_GetShee(object sender, RoutedEventArgs e)
        {
            int sheeId = Tab_3_ComboBox_Shhet.SelectedIndex;
            Tab_2_Get(GetDataVersion, GetDataReadType, GetDataFileName, sheeId);
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

    }

    class TrdRead
    {
        public int version;
        public int type;
        public string[] sutype;
        public string fileName;
        public int sheetId;
        public void Excel()
        {
            Program.OpenData(version, type, sutype, fileName, sheetId);
            return;
        }

    }

}

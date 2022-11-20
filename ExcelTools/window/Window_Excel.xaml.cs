using excelTools.data;
using excelTools.excel;
using excelTools.WPF;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Window_Excel.xaml 的交互逻辑
    /// </summary>
    public partial class Window_Excel : Window
    {
        // 数据绑定
        private static readonly WpfInfo WIfon = new();
        // 记录
        private static int GetDataVersion, GetDataSheetId;
        private static string[] GetDataReadType;
        private static string GetDataFileName;
        public int Readint = 0;
        private static int WriteType = 0;
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="type">写入类型</param>
        /// <param name="version">版本,暂时无用</param>
        /// <param name="readType">读取字段</param>
        /// <param name="_fileName">读取路径</param>
        /// <param name="sheetId">工作表</param>
        public Window_Excel(int type, int version, string[] readType, string _fileName, int sheetId)
        {
            // 窗口居中显示
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            // 数据绑定
            this.DataContext = WIfon;
            // 设置窗口名称
            this.Title = "Excel Tools - 数据导入";

            // 注册可否被关闭事件
            this.Closing += WpfClosing;

            // 读取数据
            WriteType = type;
            Read_Excel_Data(version, readType, _fileName, sheetId);

            // 实时显示写入状态
            var t1 = new Task(() => Display_Status());
            t1.Start();
        }
        #region 窗口 事件
        private void WpfClosing(object o, System.ComponentModel.CancelEventArgs e)
        {
            if (Readint == 1)
            {
                MessageBox.Show("请等待数据读取完毕后退出", "提示");
                e.Cancel = true;
                return;
            }
            if (Readint == 2)
            {
                MessageBox.Show("请等待数据写入完毕后退出", "提示");
                e.Cancel = true;
                return;
            }
            if (Config.Status() == (int)System.Data.ConnectionState.Open) Config.Close(); // 如果数据库已连接,则关闭数据库
            e.Cancel = false;
        }
        private void Grid_MouseDown(object sender, MouseButtonEventArgs e) // 双击放大缩小
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
        private void 顶层_最大_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = this.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        }
        private void 顶层_关闭_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            Key key = (e.Key == Key.System ? e.SystemKey : e.Key);
            if (key == Key.Tab)
            {
                e.Handled = true;
            }
        }
        #endregion

        /// <summary>
        /// 显示写入状态
        /// </summary>
        private void Display_Status()
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

        private void Button_Click_Writedata(object sender, RoutedEventArgs e)
        {
            if (Readint == 1)
            {
                MessageBox.Show("请等待数据读取完毕", "提示");
                return;
            }
            if (Readint == 2)
            {
                MessageBox.Show("正在写入数据,请耐心等待", "提示");
                return;
            }
            if (excelTools.excel.Program.arrWrite.Count == 0)
            {
                MessageBox.Show("无数据可写入", "提示");
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
                        MessageBox.Show("异常类型", "错误");
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

        /// <summary>
        /// 切换Excel 工作表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_GetShee(object sender, RoutedEventArgs e)
        {
            int sheeId = Tab_3_ComboBox_Shhet.SelectedIndex;
            Read_Excel_Data(GetDataVersion, GetDataReadType, GetDataFileName, sheeId);
        }

        /// <summary>
        /// 读取Excel数据
        /// </summary>
        /// <param name="version"></param>
        /// <param name="readType"></param>
        /// <param name="_fileName"></param>
        /// <param name="sheetId"></param>
        private void Read_Excel_Data(int version, string[] readType, string _fileName, int sheetId)
        {
            if (!File.Exists(_fileName)) return;
            int _excelType = _fileName[^5..] == ".xlsx" ? 2 : 1;
            // 记录耗时
            Int64 startTime = Stopwatch.GetTimestamp();
            // 线程启动读取
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
            this.Tab_3_TextBlock.Text = $"正在读取数据,请耐心等待";
            t.ContinueWith((task) =>
            {
                // 委托
                Dispatcher.Invoke(
                new Action(
                        delegate
                        {
                            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
                            DataGridHs(1, readType);
                            this.Tab_3_TextBlock.Text = $"当前数据数量：{excelTools.excel.Program.arrWrite.Count}, 耗时：{stime} s";
                        }
                    ));
            });
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
}

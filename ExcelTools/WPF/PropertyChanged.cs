using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace excelTools.WPF
{
    public class WpfInfo : INotifyPropertyChanged
    {
        private static string _Version; // 版本
        
        private static int _TabControl; // 选择
        private static string _TabWriteDataNum;


        private static string _DbConfigInfo;// 数据库信息
        private static string _DbConnectButtonBackground= "F44336";// 数据库连接按钮背景颜色

        public string Version
        {
            get { return _Version; }
            set { _Version = value; OnPropertyChanged(); }
        }
        public string DbConfigInfo
        {
            get { return _DbConfigInfo; }
            set { _DbConfigInfo = value; OnPropertyChanged(); }
        }
        public string DbConnectButtonBackground
        {
            get { return _DbConnectButtonBackground; }
            set { _DbConnectButtonBackground = value; OnPropertyChanged(); }
        }
        public int TabControl
        {

            get { return _TabControl; }
            set { _TabControl = value; OnPropertyChanged(); }
        }
        public string TabWriteDataNum
        {

            get { return _TabWriteDataNum; }
            set { _TabWriteDataNum = value; OnPropertyChanged(); }
        }
        public class TabShee
        {
            public int ID { get; set; }
            public string Name { get; set; }
        }


        /// <summary>
        /// 进行数据绑定更改
        /// </summary>
        /// <param name="name"></param>
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

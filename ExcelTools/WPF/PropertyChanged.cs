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
        private static string _Version;
        private static string _DbConfigInfo;
        private static int _TabControl;
        private static string _TabWriteDataNum;
        /// <summary>
        /// 软件版本信息
        /// </summary>
        public string Version
        {
            get { return _Version; }
            set { _Version = value; OnPropertyChanged(); }
        }
        /// <summary>
        /// 数据库信息
        /// </summary>
        public string DbConfigInfo
        {
            get { return _DbConfigInfo; }
            set { _DbConfigInfo = value; OnPropertyChanged(); }
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
        public event PropertyChangedEventHandler PropertyChanged;
        /// <summary>
        /// 进行数据绑定更改
        /// </summary>
        /// <param name="name"></param>
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

    }
}

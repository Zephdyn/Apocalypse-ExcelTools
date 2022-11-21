using excelTools.assist;
using excelTools.data;
using excelTools.excel;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelTools.data.write
{
    internal class Copy
    {
        public static int V1()
        {
            var arrWrite = excelTools.excel.Program.arrWrite;
            if (arrWrite.Count == 0) return -1;
            // 获取连接
            var con = Config.connection;
            // 开启事务
            var transaction = con.BeginTransaction();
            // sql操作
            SqliteCommand write = con.CreateCommand();
            // 获取当前时间戳
            Int64 startTime = Stopwatch.GetTimestamp();

            string copy, sqlTxt;
            int  errorCheck = 0;
            try
            {
                V1Write.state = 2;
                // 先进行处理, 将副本地图清理
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    copy = arr[(int)ExcelType_v1.Copy.副本].Trim();
                    if (copy == "") continue;
                    // 每3000提交一次事务


                    V1Write.Numwrite++; // 执行的次数
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }
                    List<object[]> list = new()
                    {
                        new object[] { "@copy", copy }
                    };

                    if (Sqlite.Select("SELECT `ID` FROM `副本地图` WHERE `copy` = @copy", list))
                        Sqlite.NonQuery("DELETE FROM `副本地图` WHERE `copy` = @copy", list);
                }
                // 清理完毕提交一次事务
                transaction.Commit();
                transaction = con.BeginTransaction();

                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("副本配置", "SELECT `ID`,`copy` FROM `副本配置`", null, "copy");

                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    copy = arr[(int)ExcelType_v1.Copy.副本].Trim();
                    if (copy == "") continue;

                    V1Write.Numwrite++; // 执行的次数
                                      // 每1000提交一次事务
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.副本类型], out int type)) type = 0;
                    if (type < 0 || type > 4) type = 0;
                    object entries; // 进入次数
                    if (type == 0 || type == 4) entries = null;
                    else
                    {
                        if (!int.TryParse(arr[(int)ExcelType_v1.Copy.进入次数], out int tnum)) entries = 0;
                        else if (tnum < 0) entries = 0;
                        else entries = tnum;
                    }

                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.奖励类型], out int rewardType)) rewardType = 0;
                    if (rewardType < 0 || rewardType > 1) rewardType = 0;

                    string rewards = arr[(int)ExcelType_v1.Copy.通过奖励];
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.副本排序], out int copySort)) copySort = 0;

                    // 限制人数,默认为0
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.限制人数], out int limitUser)) limitUser = 0;
                    // 限制消耗
                    if (arr[(int)ExcelType_v1.Copy.限制消耗] != "")
                    {
                        string ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Copy.限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.单人副本, arr[(int)ExcelType_v1.Copy.副本], arr[(int)ExcelType_v1.Copy.限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.单人副本, arr[(int)ExcelType_v1.Copy.副本], arr[(int)ExcelType_v1.Copy.限制消耗], true);


                    List<object[]> list = new()
                    {
                        new object[] { "@copy", copy }
                    };
                    if (selectArr.Exists(t => t == copy))
                    {
                        sqlTxt = "UPDATE `副本配置` SET `type`=@type, `t_number`=@entries, `reward_1`=@rewardType, `c_drop`=@rewards, `sor`=@copySort,`limit_user`=@limitUser WHERE `copy` = @copy;";
                    }
                    else
                    {
                        sqlTxt = "INSERT INTO `副本配置` (`copy`,`type`,`t_number`,`reward_1`,`c_drop`,`sor`,`limit_user`) VALUES (@copy,@type,@entries,@rewardType,@rewards,@copySort,@limitUser);";
                        selectArr.Add(copy);
                    }
                    list.Add(new object[] { "type", type });
                    list.Add(new object[] { "entries", entries });
                    list.Add(new object[] { "rewardType", rewardType });
                    list.Add(new object[] { "rewards", rewards });
                    list.Add(new object[] { "copySort", copySort });
                    list.Add(new object[] { "limitUser", limitUser });
                    Sqlite.NonQuery(sqlTxt, list);

                    // 写入地图
                    string map = arr[(int)ExcelType_v1.Copy.地图];
                    string pos_1 = arr[(int)ExcelType_v1.Copy.前];
                    string pos_2 = arr[(int)ExcelType_v1.Copy.后];
                    string pos_3 = arr[(int)ExcelType_v1.Copy.左];
                    string pos_4 = arr[(int)ExcelType_v1.Copy.右];
                    string pos_5 = arr[(int)ExcelType_v1.Copy.上];
                    string pos_6 = arr[(int)ExcelType_v1.Copy.下];
                    string monster = arr[(int)ExcelType_v1.Copy.怪物];
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.数量], out int mNum)) mNum = 0;
                    if (mNum < 0) mNum = 0;
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.强制攻击], out int force)) force = 0;
                    if (force != 0 || force != 1) force = 0;
                    string mDrop = arr[(int)ExcelType_v1.Copy.掉落];

                    // 通关地图
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.通关地图], out int entriesMap)) entriesMap = 0;
                    if (entriesMap != 0 && entriesMap != 1) entriesMap = 0;
                    // 初始地图
                    if (!int.TryParse(arr[(int)ExcelType_v1.Copy.初始地图], out int initialMap)) initialMap = 0;
                    if (initialMap != 0 && initialMap != 1) initialMap = 0;

                    list = new()
                    {
                        new object[] { "@copy", copy }
                    };
                    sqlTxt = "INSERT INTO `副本地图` (`copy`,`位置`,`前`,`后`,`左`,`右`,`上`,`下`,`怪物`,`数量`,`强制`,`打怪奖励`,`通关地图`,`初始地图`) VALUES (@copy,@map,@pos_1,@pos_2,@pos_3,@pos_4,@pos_5,@pos_6,@monser,@mNum,@force,@mDrop,0,0);";
                    list.Add(new object[] { "map", map == "" ? null : map });
                    list.Add(new object[] { "pos_1", pos_1 == "" ? null : pos_1 });
                    list.Add(new object[] { "pos_2", pos_2 == "" ? null : pos_2 });
                    list.Add(new object[] { "pos_3", pos_3 == "" ? null : pos_3 });
                    list.Add(new object[] { "pos_4", pos_4 == "" ? null : pos_4 });
                    list.Add(new object[] { "pos_5", pos_5 == "" ? null : pos_5 });
                    list.Add(new object[] { "pos_6", pos_6 == "" ? null : pos_6 });
                    list.Add(new object[] { "monser", monster == "" ? null : monster });
                    list.Add(new object[] { "mNum", monster == "" ? null : mNum });
                    list.Add(new object[] { "force", monster == "" ? null : force });
                    list.Add(new object[] { "mDrop", monster == "" ? null : mDrop });

                    Sqlite.NonQuery(sqlTxt, list);

                    list = new()
                    {
                        new object[] { "@copy", copy },
                        new object[] { "@map", map }
                    };
                    if (entriesMap == 1)
                    {
                        sqlTxt = "UPDATE `副本地图` SET `通关地图`=1 WHERE `copy`=@copy AND `位置`=@map;";
                        sqlTxt += "UPDATE `副本地图` SET `通关地图`=0 WHERE `copy`=@copy AND `位置`!=@map;";
                        Sqlite.NonQuery(sqlTxt, list);
                    }
                    if (initialMap == 1)
                    {
                        sqlTxt = "UPDATE `副本地图` SET `初始地图`=1 WHERE `copy`=@copy AND `位置`=@map;";
                        sqlTxt += "UPDATE `副本地图` SET `初始地图`=0 WHERE `copy`=@copy AND `位置`!=@map;";
                        Sqlite.NonQuery(sqlTxt, list);
                    }
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
         V1Write.   state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + V1Write.Numwrite.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }

    }
}

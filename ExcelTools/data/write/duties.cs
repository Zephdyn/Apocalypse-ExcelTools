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
    internal class Duties
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

            Int64 startTime = Stopwatch.GetTimestamp();

            string duties, show, aim, typeString, drop, dutiesSeries, dutiesNext, dutiesNpc, sqlTxt;
            int errorCheck = 0, tNum = 0;
            try
            {
                V1Write.state = 2;
                // 先进行处理, 将任务目标清理

                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    duties = arr[(int)ExcelType_v1.Duties.任务].Trim();
                    if (duties == "") continue;
                    // 每3000提交一次事务


                    V1Write.Numwrite++; // 执行的次数
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }
                    List<object[]> list = new()
                    {
                        new object[] { "@任务", duties }
                    };

                    if (Sqlite.Select("SELECT `任务` FROM `任务配置2` WHERE `任务` = @任务", list))
                        Sqlite.NonQuery("DELETE FROM `任务配置2` WHERE `任务` = @任务", list);
                }

                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("任务配置", "SELECT `任务` FROM `任务配置`", null, "任务");

                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    duties = arr[(int)ExcelType_v1.Duties.任务].Trim();
                    if (duties == "") continue;

                    V1Write.Numwrite++; // 执行的次数
                                        // 每1000提交一次事务
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }
                    if (!int.TryParse(arr[(int)ExcelType_v1.Duties.类型], out int type)) type = 2;
                    if (type < 0 || type > 3) type = 2;
                    // 任务配置
                    typeString = type switch
                    {
                        0 => "主线",
                        1 => "支线",
                        2 => "日常",
                        3 => "循环",
                        _ => "异常"
                    };
                    drop = arr[(int)ExcelType_v1.Duties.掉落];

                    List<object[]> sList = new()
                    {
                        new object[] { "@任务", duties }
                    };

                    if (selectArr.Exists(t => t == duties))
                    {
                        sqlTxt = "UPDATE `任务配置` SET `类型`=@类型 , `掉落` = @掉落 WHERE `任务` = @任务;";
                    }
                    else
                    {
                        sqlTxt = "INSERT INTO `任务配置` (`任务`,`类型`,`掉落`) VALUES (@任务,@类型,@掉落);";
                        selectArr.Add(duties);
                    }
                    sList.Add(new object[] { "@类型", typeString });
                    sList.Add(new object[] { "@掉落", drop });
                    Sqlite.NonQuery(sqlTxt, sList);

                    /*任务配置2*/

                    aim = arr[(int)ExcelType_v1.Duties.任务目标];
                    string[] tarr = aim.Split(new char[] { '\n' });
                    List<object[]> aimList = new();
                    foreach (var item in tarr)
                    {
                        string[] tarr2 = item.Split(new string[] { "[too]" }, StringSplitOptions.None);
                        if (tarr2.Length != 3) continue;
                        tarr2[0] = tarr2[0] switch
                        {
                            "怪物" => "1",
                            "物品" => "2",
                            "货币" => "3",
                            "NPC" => "4",
                            _ => ""
                        };
                        if (tarr2[0] == "") continue;
                        if (tarr2[1] == "") continue;
                        if (tarr2[0] == "货币")
                        {
                            if (!int.TryParse(tarr2[1], out tNum)) tNum = 1;
                            if (tNum < 1 || tNum > 3) tNum = 1;
                            tarr2[1] = tNum.ToString();
                        }
                        if (!int.TryParse(tarr2[2], out tNum)) tNum = 1;
                        if (tNum < 1) tNum = 1;
                        // 添加处理完毕的数组
                        aimList.Add(tarr2);
                    }

                    foreach (var item in aimList)
                    {
                        sList = new()
                        {
                            new object[] { "@任务", duties },
                            new object[] { "@类型", item[0] },
                            new object[] { "@目标", item[1] }
                        };
                        if (Sqlite.Select("SELECT `任务` FROM `任务配置2`  WHERE `任务`=@任务 AND `类型`=@类型 AND `目标`=@目标;", sList)) sqlTxt = "UPDATE `任务配置2` SET `数量` = @数量 WHERE `任务` = @任务 AND `类型`=@类型 AND `目标`=@目标;;";
                        else sqlTxt = "INSERT INTO `任务配置2` (`任务`,`类型`,`目标`,`数量`) VALUES (@任务,@类型,@目标,@数量);";
                        sList.Add(new object[] { "@数量", item[2] });
                        Sqlite.NonQuery(sqlTxt, sList);
                    }

                    /*任务配置3*/
                    dutiesSeries = arr[(int)ExcelType_v1.Duties.任务系列];
                    dutiesNext = arr[(int)ExcelType_v1.Duties.下个任务];
                    dutiesNpc = arr[(int)ExcelType_v1.Duties.交付对象];
                    sList = new()
                    {
                        new object[] { "@任务", duties }
                    };
                    if (Sqlite.Select("SELECT `任务` FROM `任务配置3`  WHERE `任务`=@任务", sList)) sqlTxt = "UPDATE `任务配置3` SET `系列`=@系列 , `下个` = @下个, `交付` = @交付, `完成` = @完成 WHERE `任务` = @任务";
                    else sqlTxt = "INSERT INTO `任务配置3` (`任务`,`系列`,`下个`,`交付`,`完成`) VALUES (@任务,@系列,@下个,@交付,@完成)";
                    sList.Add(new object[] { "@系列", dutiesSeries });
                    sList.Add(new object[] { "@下个", dutiesNext });
                    sList.Add(new object[] { "@交付", dutiesNpc });
                    sList.Add(new object[] { "@完成", "0" });
                    Sqlite.NonQuery(sqlTxt, sList);

                    if (!int.TryParse(arr[(int)ExcelType_v1.Duties.放弃], out int giveUp)) giveUp = 0;
                    if (giveUp < 0 || giveUp > 1) giveUp = 0;
                    Update.DbConfig("任务配置", duties, "放弃", giveUp.ToString(), false);
                    show = arr[(int)ExcelType_v1.Duties.描述];
                    Update.Show(Update.EShow.任务, duties, show, show == "");


                    if (arr[(int)ExcelType_v1.Duties.限制消耗] != "")
                    {
                        string ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Duties.限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.任务, arr[(int)ExcelType_v1.Duties.任务], arr[(int)ExcelType_v1.Duties.限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.任务, arr[(int)ExcelType_v1.Duties.任务], arr[(int)ExcelType_v1.Duties.限制消耗], true);
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
            V1Write.state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + V1Write.Numwrite.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");

            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }

    }
}

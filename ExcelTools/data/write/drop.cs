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
    internal class Drop
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

            try
            {
                // 删除所有掉落记录
                V1Write.state = 2;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Drop.掉落].Trim() == "") continue;

                    V1Write.Numwrite++;
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }


                    List<object[]> list = new()
                    {
                        new object[2] { "@id", arr[(int)ExcelType_v1.Drop.掉落] }
                    };

                    if (Sqlite.Select("SELECT id FROM `掉落配置` WHERE `掉落` = @id", list)) // 查询是否已有掉落
                        Sqlite.NonQuery("DELETE  FROM `掉落配置` WHERE `掉落` = @id", list);

                }

                V1Write.Numwrite = 0;
                V1Write.state = 1;
                int por = 0, min = 0, max = 0, del = 0;
                string temp = "", goods = null;
                string[] tarr, yarr, narr;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Drop.掉落].Trim() == "") continue;
                    // 类型必须为1-3
                    if (!int.TryParse(arr[(int)ExcelType_v1.Drop.类型], out int dropType)) continue;
                    if (dropType < 1 || dropType > 3) continue;

                    V1Write.Numwrite++;
                    // 每1000提交一次事务
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    List<object[]> list = new()
                    {
                        new object[2] { "@id", arr[(int)ExcelType_v1.Drop.掉落] }
                    };

                    if (!int.TryParse(arr[(int)ExcelType_v1.Drop.掉落地图], out int dropMap)) dropMap = 0;
                    /*格式:
                     *物品名称[too]最小数量[too]最多数量[too]掉落几率[next]限制职业[too]限制职业2
                     *无职业限制则为 格式2: 物品名称[too]最小数量[too]最多数量[too]掉落几率
                     *格式3: 物品名称[too]掉落数量[too]掉落几率
                     */

                    for (var j = 1; j < 4; j++)
                    {
                        switch (j)
                        {
                            case 3:
                                temp = arr[(int)ExcelType_v1.Drop.物品];
                                del = 0;
                                break;
                            case 1:
                                temp = arr[(int)ExcelType_v1.Drop.经验];
                                del = 1;// 每个索引-1
                                break;
                            case 2:
                                temp = arr[(int)ExcelType_v1.Drop.货币];
                                del = 0;
                                break;
                        }
                        tarr = temp.Split(new char[] { '\n' });
                        foreach (string item in tarr)
                        {
                            narr = item.Split(new string[] { "[next]" }, StringSplitOptions.None);
                            if (narr.Length < 1 || narr.Length > 2) continue;
                            yarr = narr[0].Split(new string[] { "[too]" }, StringSplitOptions.None);
                            if (yarr.Length != 4 && yarr.Length != 3)
                            {
                                if (j != 1 || yarr.Length != 2) continue;
                            }

                            // 物品
                            if (j == 3)
                            {
                                if (yarr[0] == "") continue;
                                else goods = yarr[0];
                            }


                            // 货币
                            if (j == 2)
                            {
                                if (!int.TryParse(yarr[0], out int num) && num > 3) num = 1;
                                goods = num.ToString();
                            };
                            // 经验
                            if (j == 1) goods = null;

                            // 几率
                            if ((j != 1 && yarr.Length == 3) || (j == 1 && yarr.Length == 2)) { if (!int.TryParse(yarr[2 - del], out por)) por = 0; }
                            else { if (!int.TryParse(yarr[3 - del], out por)) por = 0; }

                            // 最小数量
                            if ((j != 1 && yarr.Length == 3) || (j == 1 && yarr.Length == 2)) { if (!int.TryParse(yarr[1 - del], out min)) min = 0; }
                            else { if (!int.TryParse(yarr[1 - del], out min)) min = 0; }

                            // 最大数量
                            if ((j != 1 && yarr.Length == 3) || (j == 1 && yarr.Length == 2)) max = min;
                            else { if (!int.TryParse(yarr[2 - del], out max)) max = 0; }
                            list = new()
                            {
                                new object[2] { "@id", arr[(int)ExcelType_v1.Drop.掉落] }, // 掉落ID
                                new object[2] { "@dropMap", dropMap }, // 掉落地图
                                new object[2] { "@type", j },// 类型
                                new object[2] { "@dropType", dropType },// 掉落类型
                                new object[2] { "@goods", goods },// 物品
                                new object[2] { "@por", por }, // 几率
                                new object[2] { "@min", min }, // 最小几率
                                new object[2] { "@max", max }, // 最大几率
                                new object[2] { "@occ", narr.Length == 1 ? null : narr[1] }// 职业
                            };
                            Sqlite.NonQuery("INSERT INTO `掉落配置` (`掉落`,`类型`,`物品`,`几率`,`最小数量`,`最大数量`,`掉落类型`,`职业`,`地图掉落`) VALUES (@id, @type, @goods, @por, @min, @max, @dropType, @occ, @dropMap)", list);
                        }
                    }
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
            return 0;
        }

    }
}

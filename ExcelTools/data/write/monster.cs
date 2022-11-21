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
    internal class Monster
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
                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("config_monster", "SELECT `id`,`monster` FROM `config_monster`", null, "monster");
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Monster.怪物].Trim() == "") continue;
                    V1Write.Numwrite++;
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    List<object[]> list = new()
                    {
                        new object[2] { "@monster", arr[(int)ExcelType_v1.Monster.怪物] }
                    };

                    if (selectArr.Exists(t => t == arr[(int)ExcelType_v1.Monster.怪物]))
                    {
                        write.CommandText = @"
                        UPDATE `config_monster` SET
                        `lv` = @lv, `life` =@life, `mana` =@mana, `attack` =@attack, `def` =@def, `crit_prob` =@prob,
                        `crit_hurt` =@hurt, `drop` = @drop, `show` =@show, `dodge`=@dodge, `hit` = @hit, `show_map`=@show_map
                        WHERE `monster` = @monster";

                    }
                    else
                    {
                        write.CommandText = @"
                        INSERT INTO `config_monster`
                        (`monster`,`lv`,`life`,`mana`,`attack`,`def`,`crit_prob`,`crit_hurt`,`drop`,`show`,`dodge`,`hit`,`show_map`) VALUES
                        (@monster,@lv,@life,@mana,@attack,@def,@prob,@hurt,@drop,@show,@dodge,@hit,@show_map);";
                        selectArr.Add(arr[(int)ExcelType_v1.Monster.怪物]);
                    }
                    // 对数据进行处理
                    write.Parameters.AddWithValue("@monster", arr[(int)ExcelType_v1.Monster.怪物]);
                    if (!Int64.TryParse(arr[(int)ExcelType_v1.Monster.等级], out Int64 lnum)) lnum = 1;
                    if (lnum < 1) lnum = 1;
                    write.Parameters.AddWithValue("@lv", lnum);
                    if (!Int64.TryParse(arr[(int)ExcelType_v1.Monster.生命], out lnum)) lnum = 1;
                    if (lnum < 1) lnum = 1;
                    write.Parameters.AddWithValue("@life", lnum);
                    if (!Int64.TryParse(arr[(int)ExcelType_v1.Monster.法力], out lnum)) lnum = 1;
                    if (lnum < 1) lnum = 1;
                    write.Parameters.AddWithValue("@mana", lnum);
                    if (!Int64.TryParse(arr[(int)ExcelType_v1.Monster.攻击], out lnum)) lnum = 1;
                    if (lnum < 1) lnum = 1;
                    write.Parameters.AddWithValue("@attack", lnum);
                    if (!Int64.TryParse(arr[(int)ExcelType_v1.Monster.防御], out lnum)) lnum = 1;
                    if (lnum < 1) lnum = 1;
                    write.Parameters.AddWithValue("@def", lnum);
                    if (!double.TryParse(arr[(int)ExcelType_v1.Monster.暴击概率], out double dnum)) dnum = 0;
                    if (dnum < 0) dnum = 0;
                    write.Parameters.AddWithValue("@prob", dnum);
                    if (!double.TryParse(arr[(int)ExcelType_v1.Monster.暴击伤害], out dnum)) dnum = 0;
                    if (dnum < 0) dnum = 0;
                    write.Parameters.AddWithValue("@hurt", dnum);
                    if (!double.TryParse(arr[(int)ExcelType_v1.Monster.闪避], out dnum)) dnum = 0;
                    if (dnum < 0) dnum = 0;
                    write.Parameters.AddWithValue("@dodge", dnum);
                    if (!double.TryParse(arr[(int)ExcelType_v1.Monster.命中], out dnum)) dnum = 0;
                    if (dnum < 0) dnum = 0;
                    write.Parameters.AddWithValue("@hit", dnum);

                    write.Parameters.AddWithValue("@drop", TString.HandingBank(arr[(int)ExcelType_v1.Monster.掉落]));
                    write.Parameters.AddWithValue("@show", arr[(int)ExcelType_v1.Monster.对外显示]);
                    write.Parameters.AddWithValue("@show_map", arr[(int)ExcelType_v1.Monster.地图显示]);

                    write.ExecuteNonQuery();
                    write.Parameters.Clear();
                    // 描述
                    Update.Show(Update.EShow.怪物, arr[(int)ExcelType_v1.Monster.怪物], arr[(int)ExcelType_v1.Monster.描述], false);

                    // 怪物技能
                    string skill = arr[(int)ExcelType_v1.Monster.技能];
                    if (skill == "") continue;
                    string[] tarr = skill.Split(new char[] { '\n' });
                    Sqlite.NonQuery("DELETE FROM `怪物技能` WHERE `怪物` = @monster", list);


                    foreach (string item in tarr)
                    {
                        list = new()
                        {
                            new object[] { "@monster", arr[(int)ExcelType_v1.Monster.怪物] },
                            new object[] { "@skill", item }
                        };
                        Sqlite.NonQuery("INSERT INTO `怪物技能` (`怪物`,`技能`) VALUES (@monster,@skill)", list);
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

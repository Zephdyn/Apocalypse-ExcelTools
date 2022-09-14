using System;
using System.Collections.Generic;
using System.Linq;

using excelTools.excel;
using excelTools.assist;
using Microsoft.Data.Sqlite;
using System.Windows;
using System.Data;
using System.Diagnostics;

namespace excelTools.data
{
    internal class V1Write
    {
        // 0未开启, 1 v1插入
        public static int state;
        // 写入值
        public static int Numwrite;

        // 0正确 -1当前无数据 1写入失败

        /// <summary>
        /// 装备
        /// </summary>
        /// <returns></returns>
        public static int Equip()
        {
            List<string[]> arrWrite = Program.arrWrite;
            if (arrWrite.Count == 0)
            {
                return -1;
            }
            // 获取连接
            SqliteConnection con = Config.connection;
            // 开启事务
            SqliteTransaction transaction = con.BeginTransaction();
            // sql操作
            SqliteCommand write = con.CreateCommand();

            long startTime = Stopwatch.GetTimestamp();

            int testing, numCount = 0;

            try
            {
                state = 1;
                for (int i = 0; i < arrWrite.Count; i++)
                {
                    numCount++;
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    Numwrite = i + 1;
                    string[] arr = arrWrite[i].ToArray();
                    List<object[]> list = new()
                    {
                        new object[2] { "@装备", arr[(int)ExcelType_v1.Equip.名称] }
                    };

                    if (Sqlite.Select("SELECT `装备` FROM `装备配置` WHERE `装备` = @装备", list))
                    {
                        write.CommandText = @"
                        UPDATE `装备配置` SET
                        `部位` = @部位, `等级` = @等级, `生命` = @生命, `法力` = @法力, `攻击` = @攻击, `防御` = @防御, `闪避`=@闪避, `命中`=@命中,
                        `暴击概率` = @暴击概率, `暴击伤害` = @暴击伤害, `技能` = @技能, `饥饿` = @饥饿, `bind` = @bind
                        WHERE `装备` = @装备";

                    }
                    else
                    {
                        write.CommandText = @"
                        INSERT INTO `装备配置`
                        (`装备`,`部位`,`等级`,`生命`,`法力`,`攻击`,`防御`,`闪避`,`命中`,`暴击概率`,`暴击伤害`,`技能`,`饥饿`,`bind`) VALUES
                        (@装备,@部位,@等级,@生命,@法力,@攻击,@防御,@闪避,@命中,@暴击概率,@暴击伤害,@技能,@饥饿,@bind);";
                    }
                    // 对数据进行处理
                    //"名称",
                    write.Parameters.AddWithValue("@装备", arr[(int)ExcelType_v1.Equip.名称]);
                    //"职业",
                    //"绑定",
                    arr[3] = arr[3] == "1" ? "1" : "0";
                    write.Parameters.AddWithValue("@bind", arr[(int)ExcelType_v1.Equip.绑定]);
                    //"部位",
                    write.Parameters.AddWithValue("@部位", arr[(int)ExcelType_v1.Equip.部位]);
                    //"等级",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.等级], out int num)) num = 1;
                    write.Parameters.AddWithValue("@等级", num);
                    //"生命",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.生命], out num)) num = 0;
                    write.Parameters.AddWithValue("@生命", num);
                    //"法力",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.法力], out num)) num = 0;
                    write.Parameters.AddWithValue("@法力", num);
                    //"攻击",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.攻击], out num)) num = 0;
                    write.Parameters.AddWithValue("@攻击", num);
                    //"防御",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.防御], out num)) num = 0;
                    write.Parameters.AddWithValue("@防御", num);
                    //"命中",
                    if (!double.TryParse(arr[(int)ExcelType_v1.Equip.命中], out double dnum)) dnum = 0;
                    write.Parameters.AddWithValue("@命中", dnum.ToString("0.00"));
                    //"闪避",
                    if (!double.TryParse(arr[(int)ExcelType_v1.Equip.闪避], out dnum)) dnum = 0;
                    write.Parameters.AddWithValue("@闪避", dnum.ToString("0.00"));
                    //"暴击概率",
                    if (!double.TryParse(arr[(int)ExcelType_v1.Equip.暴击概率], out dnum)) dnum = 0;
                    write.Parameters.AddWithValue("@暴击概率", dnum.ToString("0.00"));
                    //"暴击伤害",
                    if (!double.TryParse(arr[(int)ExcelType_v1.Equip.暴击伤害], out dnum)) dnum = 0;
                    write.Parameters.AddWithValue("@暴击伤害", dnum.ToString("0.00"));
                    //"饥饿",
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.饥饿], out num)) dnum = 0;
                    write.Parameters.AddWithValue("@饥饿", num);
                    //"卸下",
                    //"技能",
                    if (arr[(int)ExcelType_v1.Equip.技能] == null) arr[(int)ExcelType_v1.Equip.技能] = "";
                    write.Parameters.AddWithValue("@技能", arr[(int)ExcelType_v1.Equip.技能]);
                    write.ExecuteNonQuery();
                    write.Parameters.Clear();
                    //"描述"
                    Update.Show(Update.EShow.装备, arr[(int)ExcelType_v1.Equip.名称], arr[(int)ExcelType_v1.Equip.描述], false);

                    list = new()
                    {
                        new object[2] { "@装备", arr[(int)ExcelType_v1.Equip.名称] },
                        new object[2] { "@类型", "1" }
                    };
                    // 职业 [装备配置2]
                    // arr[2]
                    // 对arr[2]进行处理
                    string temp = TString.HandingBank(arr[(int)ExcelType_v1.Equip.职业]);
                    if (Sqlite.Select("SELECT `装备` FROM `装备配置2` WHERE `装备` = @装备 AND `类型` = 1", list))
                    {
                        write.CommandText = @"
                        UPDATE `装备配置2` SET
                        `类型` = 1, `目标` = @目标
                        WHERE `装备` = @装备";
                        testing = 0;
                    }
                    else
                    {
                        write.CommandText = @"
                        INSERT INTO `装备配置2`
                        (`装备`,`类型`,`目标`) VALUES
                        (@装备,@类型,@目标);";
                        testing = 1;
                    }
                    write.Parameters.AddWithValue("@装备", arr[(int)ExcelType_v1.Equip.名称]);
                    write.Parameters.AddWithValue("@类型", 1);
                    write.Parameters.AddWithValue("@目标", temp);
                    if (testing == 0 || (testing == 1 && temp != ""))
                        write.ExecuteNonQuery();
                    write.Parameters.Clear();

                    // 卸下 [装备特性配置]
                    if (Sqlite.Select("SELECT `ID` FROM `装备特性配置` WHERE `equip` = @装备 AND `type` = @类型", list))
                    {
                        write.CommandText = @"
                        UPDATE `装备特性配置` SET
                        `type` = @类型, `state` = @状态
                        WHERE `equip` = @装备";
                        testing = 0;
                    }
                    else
                    {
                        write.CommandText = @"
                        INSERT INTO `装备特性配置`
                        (`equip`,`state`,`type`) VALUES
                        (@装备,@状态,@类型);";
                        testing = 1;
                    }
                    write.Parameters.AddWithValue("@装备", arr[(int)ExcelType_v1.Equip.名称]);
                    if (!int.TryParse(arr[(int)ExcelType_v1.Equip.卸下], out num)) num = 0;
                    if (num < 0 || num > 1) num = 0;
                    write.Parameters.AddWithValue("@状态", num);
                    write.Parameters.AddWithValue("@类型", 1);
                    if (testing == 0 || (testing == 1 && num == 1))
                        write.ExecuteNonQuery();
                    write.Parameters.Clear();
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            return 0;
        }
        /// <summary>
        /// 地图
        /// </summary>
        /// <returns></returns>
        public static int Map()
        {
            var arrWrite = excelTools.excel.Program.arrWrite;
            if (arrWrite.Count == 0) return -1;
            // 获取连接
            var con = Config.connection;
            // 开启事务
            var transaction = con.BeginTransaction();
            // sql操作
            SqliteCommand write = con.CreateCommand();

            long startTime = Stopwatch.GetTimestamp();

            int numCount = 0, errorCheck = 0;
            string temp, ret;
            try
            {
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {

                    string[] arr = arrWrite[i].ToArray();
                    if (arr[(int)ExcelType_v1.Map.位置].Trim() == "") continue;
                    numCount++;
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    Numwrite = i + 1;
                    List<object[]> list = new()
                    {
                        new object[2] { "@map", arr[(int)ExcelType_v1.Map.位置] }
                    };

                    if (Sqlite.Select("SELECT `id` FROM `config_map` WHERE `map` = @map", list))
                    {
                        write.CommandText = @"
                        UPDATE `config_map` SET
                        `pos_1`= @pos1 , `pos_2` =@pos2, `pos_3`= @pos3, `pos_4` = @pos4, `pos_5` = @pos5, `pos_6` = @pos6,
                        `status`= @status
                        WHERE `map` = @map";

                    }
                    else
                    {
                        write.CommandText = @"
                        INSERT INTO `config_map`
                        (`map`,`pos_1`, `pos_2`,`pos_3`,`pos_4` ,`pos_5`, `pos_6`,`status`) VALUES
                        (@map, @pos1, @pos2, @pos3, @pos4, @pos5, @pos6, @status);";
                    }
                    // 对数据进行处理
                    write.Parameters.AddWithValue("@map", arr[(int)ExcelType_v1.Map.位置]);
                    write.Parameters.AddWithValue("@pos1", arr[(int)ExcelType_v1.Map.前]);
                    write.Parameters.AddWithValue("@pos2", arr[(int)ExcelType_v1.Map.后]);
                    write.Parameters.AddWithValue("@pos3", arr[(int)ExcelType_v1.Map.左]);
                    write.Parameters.AddWithValue("@pos4", arr[(int)ExcelType_v1.Map.右]);
                    write.Parameters.AddWithValue("@pos5", arr[(int)ExcelType_v1.Map.上]);
                    write.Parameters.AddWithValue("@pos6", arr[(int)ExcelType_v1.Map.下]);
                    if (!int.TryParse(arr[(int)ExcelType_v1.Map.安全区], out int num)) num = 0;
                    else if (num < 0 || num > 1) num = 0;
                    write.Parameters.AddWithValue("@status", num);
                    write.ExecuteNonQuery();
                    write.Parameters.Clear();
                    // 描述
                    Update.Show(Update.EShow.地图, arr[(int)ExcelType_v1.Map.位置], arr[(int)ExcelType_v1.Map.描述], false);
                    // 限制消耗
                    if (arr[(int)ExcelType_v1.Map.限制消耗] != "")
                    {
                        ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Map.限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.地图, arr[(int)ExcelType_v1.Map.位置], arr[(int)ExcelType_v1.Map.限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.地图, arr[(int)ExcelType_v1.Map.位置], arr[(int)ExcelType_v1.Map.限制消耗], true);

                    // 饥饿消耗
                    if (!double.TryParse(arr[(int)ExcelType_v1.Map.饥饿减少], out double dnum)) dnum = 1;
                    Update.PluginConfig(Update.EPluginConfig.地图, arr[(int)ExcelType_v1.Map.位置], "1", dnum.ToString("0.00"), false);

                    // 怪物刷新
                    temp = arr[(int)ExcelType_v1.Map.怪物];
                    // 删除已有数量以及已有配置
                    write.CommandText = "DELETE FROM `怪物刷新配置` WHERE 位置 = @map";

                    write.Parameters.AddWithValue("@map", arr[(int)ExcelType_v1.Map.位置]);
                    write.ExecuteNonQuery();
                    write.Parameters.Clear();
                    if (temp == "") continue;

                    string[] tarr = temp.Split(new char[] { '\n' });
                    foreach (string item in tarr)
                    {
                        string[] yarr = item.Split(new string[] { "[too]" }, StringSplitOptions.None);
                        if (yarr.Length != 4) continue;
                        if (yarr[0].Trim() == "") continue;
                        write.CommandText = @"
                        INSERT INTO `怪物刷新配置`
                        (`位置`,`怪物`,`刷新数量`,`刷新间隔`,`攻击几率`) VALUES
                        (@位置, @怪物, @刷新数量, @刷新间隔, @攻击几率);";
                        write.Parameters.AddWithValue("@位置", arr[(int)ExcelType_v1.Map.位置]);
                        write.Parameters.AddWithValue("@怪物", yarr[0]);
                        if (!int.TryParse(yarr[1], out num)) num = 0;
                        else if (num < 0) num = 0;
                        write.Parameters.AddWithValue("@刷新数量", num);
                        if (!int.TryParse(yarr[2], out num)) num = 0;
                        else if (num < 0) num = 0;
                        write.Parameters.AddWithValue("@刷新间隔", num);
                        if (!int.TryParse(yarr[3], out num)) num = 0;
                        else if (num < 0) num = 0;
                        write.Parameters.AddWithValue("@攻击几率", num);
                        write.ExecuteNonQuery();
                        write.Parameters.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }
        /// <summary>
        /// 怪物
        /// </summary>
        /// <returns></returns>
        public static int Monster()
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

            int numCount = 0;
            try
            {
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Monster.怪物].Trim() == "") continue;
                    numCount++;
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    Numwrite = i + 1;
                    List<object[]> list = new()
                    {
                        new object[2] { "@monster", arr[(int)ExcelType_v1.Monster.怪物] }
                    };

                    if (Sqlite.Select("SELECT `id` FROM `config_monster` WHERE `monster` = @monster", list))
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
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            return 0;
        }
        /// <summary>
        /// 掉落
        /// </summary>
        /// <returns></returns>
        public static int Drop()
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

            int numCount = 0;
            try
            {
                // 删除所有掉落记录
                state = 2;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Drop.掉落].Trim() == "") continue;
                    numCount++;
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }

                    Numwrite = i + 1;
                    List<object[]> list = new()
                    {
                        new object[2] { "@id", arr[(int)ExcelType_v1.Drop.掉落] }
                    };

                    if (Sqlite.Select("SELECT id FROM `掉落配置` WHERE `掉落` = @id", list)) // 查询是否已有掉落
                        Sqlite.NonQuery("DELETE  FROM `掉落配置` WHERE `掉落` = @id", list);

                }

                numCount = 0;
                state = 1;
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
                    numCount++;
                    // 每1000提交一次事务
                    if (i > 0 && i % 1000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    Numwrite = i + 1;
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
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            return 0;
        }
        /// <summary>
        /// 物品
        /// </summary>
        /// <returns></returns>
        public static int Goods()
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

            string ret, sqlTxt;
            int errorCheck = 0;
            int numCount = 0;
            try
            {
                // 0(null) 1(num) 2(double) 3(string)
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Goods.物品].Trim() == "") continue;


                    /*格式:
                     *1(普通消耗) 2(礼包) 3(VIP等级) 4(VIP积分) 5(技能学习) 6(坐骑获取) 7(坐骑喂养) 8(亲密度物品) 9(数据清除) 10(鱼饵)
                     */
                    // 类型必须为1-10
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.类型], out int goodsType)) continue;
                    if (goodsType < 1 || goodsType > 10) continue;
                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }

                    // 写入名称
                    object goodsName;
                    if (goodsType == 3 || goodsType == 4 || goodsType == 7 || goodsType == 8)
                    {
                        // 内容为整数
                        if (!int.TryParse(arr[(int)ExcelType_v1.Goods.名称], out int num)) num = 0;
                        goodsName = num;
                    }
                    else if (goodsType == 9)
                    {
                        // 内容为空
                        goodsName = null;
                    }
                    else if (goodsType == 10)
                    {
                        // 内容为双精度小数
                        if (!double.TryParse(arr[(int)ExcelType_v1.Goods.名称], out double dnum)) dnum = 0;
                        dnum = double.Parse(dnum.ToString("0.00")); // 保证两位小数
                        goodsName = dnum;
                    }
                    else
                    {
                        // 内容为字符串
                        goodsName = arr[(int)ExcelType_v1.Goods.名称];
                    }
                    // 绑定
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.绑定], out int goodsBind)) goodsBind = 0;
                    if (goodsBind < 0 || goodsBind > 1) goodsBind = 0;
                    // 批量
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.批量], out int goodsBatch)) goodsBatch = 0;
                    if (goodsBatch < 0 || goodsBatch > 1) goodsBatch = 0;
                    // 战斗
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.战斗], out int goodsFight)) goodsFight = 1;
                    if (goodsFight < 0 || goodsFight > 1) goodsFight = 0;
                    List<object[]> list = new()
                    {
                        new object[2] { "@goods", arr[(int)ExcelType_v1.Goods.物品] },

                    };

                    if (Sqlite.Select("SELECT `道具` FROM `道具配置` WHERE `道具`= @goods", list)) sqlTxt = "UPDATE `道具配置` SET `类型` = @type ,`名称` = @value,`绑定` = @bind,'批量' = @batch,`战斗` = @fight WHERE `道具` = @goods";
                    else sqlTxt = "INSERT INTO `道具配置` (`道具`,`类型`,`名称`,`绑定`,'批量',`战斗`) VALUES (@goods, @type, @value, @bind, @batch, @fight)";

                    list = new()
                    {
                        new object[2] { "@goods", arr[(int)ExcelType_v1.Goods.物品] },
                        new object[2] { "@type", goodsType },
                        new object[2] { "@value", goodsName },
                        new object[2] { "@bind", goodsBind },
                        new object[2] { "@batch", goodsBatch },
                        new object[2] { "@fight", goodsFight }

                    };
                    Sqlite.NonQuery(sqlTxt, list);

                    // 限制消耗
                    if (arr[(int)ExcelType_v1.Goods.限制消耗] != "")
                    {
                        ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Goods.限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.物品, arr[(int)ExcelType_v1.Goods.物品], arr[(int)ExcelType_v1.Goods.限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.物品, arr[(int)ExcelType_v1.Goods.物品], arr[(int)ExcelType_v1.Goods.限制消耗], true);

                    // 冷却
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.冷却], out int goodsWait)) goodsWait = 0;
                    if (goodsWait < 0) goodsWait = 0;
                    Update.PluginConfig(Update.EPluginConfig.物品, arr[(int)ExcelType_v1.Goods.物品], "2", goodsWait.ToString(), false);

                    // 使用提示
                    Update.Show(Update.EShow.物品提示, arr[(int)ExcelType_v1.Goods.物品], arr[(int)ExcelType_v1.Goods.提示], arr[(int)ExcelType_v1.Goods.提示] == "");
                    // 描述
                    Update.Show(Update.EShow.物品, arr[(int)ExcelType_v1.Goods.物品], arr[(int)ExcelType_v1.Goods.描述], arr[(int)ExcelType_v1.Goods.描述] == "");
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");

            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }
        /// <summary>
        /// 商店
        /// </summary>
        /// <returns></returns>
        public static int Shop()
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

            string map, priceName, show, stock, intro, sqlTxt, shop;
            int numCount = 0, nNum, oNum;
            try
            {
                // 0(null) 1(num) 2(double) 3(string)
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    shop = arr[(int)ExcelType_v1.Shop.商店].Trim();
                    if (shop == "") continue;
                    if (arr[(int)ExcelType_v1.Shop.名称].Trim() == "") continue;
                    // 商店类型 0为出售 1为收购
                    if (!int.TryParse(arr[(int)ExcelType_v1.Shop.类型], out int type)) type = 0;
                    if (type < 0 || type > 1) type = 0;

                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    map = arr[(int)ExcelType_v1.Shop.位置];
                    int typeMap = map == "" ? 1 : 2; // 商店位置类型,1为全图, 2为单地图

                    // 价格类型 -1,1,2,3
                    if (!int.TryParse(arr[(int)ExcelType_v1.Shop.价格类型], out int priceType)) priceType = 1;
                    if (priceType < 1 || priceType > 3) if (priceType != -1) priceType = 1;
                    // 物品名称 货币时无需名称
                    if (priceType == -1) priceName = arr[(int)ExcelType_v1.Shop.物品名称];
                    else priceName = "";

                    show = arr[(int)ExcelType_v1.Shop.描述];
                    stock = arr[(int)ExcelType_v1.Shop.库存];
                    string[] arrS = stock.Split(new string[] { "[too]" }, StringSplitOptions.None);

                    int tNum = arrS.Length;
                    // 对arrS中的内容进行处理,将其修改为整数
                    for (var j = 0; j < tNum; j++)
                    {
                        if (!int.TryParse(arrS[j], out int tNum2)) tNum2 = 0;
                        if (tNum2 < 0) tNum2 = 0;
                        arrS[j] = tNum2.ToString();
                    }
                    switch (tNum)
                    {
                        case 0:
                            stock = "0[分割]0";
                            nNum = 0;
                            break;
                        case 1:
                            stock = arrS[0] + "[分割]" + arrS[0];
                            nNum = int.Parse(arrS[0]);
                            break;
                        case 2:
                            stock = arrS[0] + "[分割]" + arrS[1];
                            Random rd = new();
                            nNum = rd.Next(int.Parse(arrS[0]), int.Parse(arrS[1]));
                            break;
                        default:
                            stock = "0[分割]0";
                            nNum = 0;
                            break;
                    }

                    // 商品单价
                    string priceOne = arr[(int)ExcelType_v1.Shop.单价];
                    arrS = priceOne.Split(new string[] { "[too]" }, StringSplitOptions.None);

                    tNum = arrS.Length;
                    // 对arrS中的内容进行处理,将其修改为整数
                    for (var j = 0; j < tNum; j++)
                    {
                        if (!int.TryParse(arrS[j], out int tNum2)) tNum2 = 0;
                        if (tNum2 < 0) tNum2 = 0;
                        arrS[j] = tNum2.ToString();
                    }
                    switch (tNum)
                    {
                        case 0:
                            priceOne = "0[分割]0";
                            oNum = 0;
                            break;
                        case 1:
                            priceOne = arrS[0] + "[分割]" + arrS[0];
                            oNum = int.Parse(arrS[0]);
                            break;
                        case 2:
                            priceOne = arrS[0] + "[分割]" + arrS[1];
                            Random rd = new();
                            oNum = rd.Next(int.Parse(arrS[0]), int.Parse(arrS[1]));
                            break;
                        default:
                            priceOne = "0[分割]0";
                            oNum = 0;
                            break;
                    }
                    // 刷新类型
                    if (!int.TryParse(arr[(int)ExcelType_v1.Shop.刷新类型], out int restTyep)) restTyep = 1;
                    if (restTyep < 1 || restTyep > 2) restTyep = 1;

                    // 刷新几率
                    if (!int.TryParse(arr[(int)ExcelType_v1.Shop.刷新几率], out int restPro)) restPro = 0;
                    if (restPro < 0 || restPro > 100) restPro = 0;

                    intro = arr[(int)ExcelType_v1.Shop.介绍];




                    List<object[]> list = new()
                    {
                        new object[2] { "@商店", shop },
                        new object[2] { "@物品", arr[(int)ExcelType_v1.Shop.名称].Trim() }
                    };

                    // 商店收购, 所有内容加密
                    if (type == 1)
                    {
                        for (var j = 0; j < list.Count; j++)
                        {
                            object value = null;
                            if (list[j].GetValue(1) != null)
                            {
                                value = (string)list[j].GetValue(1).ToString();
                                value = Code.Asm_TextTo16((string)value);
                            }
                            list[j].SetValue(value, 1);
                        }
                    }
                    if (Sqlite.Select($"SELECT `商店` FROM `{(type == 0 ? "商店配置2" : "商店配置3")}` WHERE `商店` = @商店 AND `物品` = @物品", list))
                    {

                        sqlTxt = $"UPDATE `{(type == 0 ? "商店配置2" : "商店配置3")}`  SET `状态`=@状态,`物品`=@物品,`数量`=@数量,`数量2`=@数量2,`数量3`=@数量3,`价格类型`=@价格类型,`价格`=@价格,`价格2`=@价格2,`刷新几率`=@刷新几率,`刷新类型`=@刷新类型,`刷新时间`=@刷新时间,`描述`=@描述{(type == 0 ? ",`价格物品`=@价格物品" : "")}";
                        sqlTxt += " WHERE `商店` =@商店 AND `物品` = @物品";
                    }
                    else
                    {
                        sqlTxt = $"INSERT INTO `{(type == 0 ? "商店配置2" : "商店配置3")}` (`商店`,`状态`,`物品`,`数量`,`数量2`,`数量3`,`价格类型`,`价格`,`价格2`,`刷新几率`,`刷新类型`,`刷新时间`,`描述`{(type == 0 ? ",`价格物品`" : "")})";
                        sqlTxt += $" VALUES (@商店,@状态,@物品,@数量,@数量2,@数量3,@价格类型,@价格,@价格2,@刷新几率,@刷新类型,@刷新时间,@描述{(type == 0 ? ",@价格物品" : "")})";
                    }

                    list = new()
                    {
                        new object[2] { "@商店", shop },
                        new object[2] { "@物品", arr[(int)ExcelType_v1.Shop.名称].Trim() },
                        new object[2] { "@状态", nNum == 0 ? 0 : 1 }, // 有数量为1. 无数量状态为0
                        new object[2] { "@数量", nNum },
                        new object[2] { "@数量2", nNum },
                        new object[2] { "@数量3", stock },
                        new object[2] { "@价格类型", priceType },
                        new object[2] { "@价格", oNum },
                        new object[2] { "@价格2", priceOne },
                        new object[2] { "@刷新几率", restPro },
                        new object[2] { "@刷新类型", restTyep },
                        new object[2] { "@刷新时间", TString.GetTime(1) },
                        new object[2] { "@描述", intro }
                    };
                    if (type == 0)
                        list.Add(new object[2] { "@价格物品", priceName });

                    // 商店收购, 所有内容加密
                    if (type == 1)
                    {
                        for (var j = 0; j < list.Count; j++)
                        {
                            object value = null;
                            if (list[j].GetValue(1) != null)
                            {
                                value = (string)list[j].GetValue(1).ToString();
                                value = Code.Asm_TextTo16((string)value);
                            }
                            list[j].SetValue(value, 1);
                        }
                    }

                    Sqlite.NonQuery(sqlTxt, list);

                    list = new()
                    {
                        new object[2] { "@商店", shop }
                    };
                    if (Sqlite.Select($"SELECT `商店` FROM `商店配置` WHERE `商店` = @商店", list))
                    {

                        sqlTxt = $"UPDATE `商店配置`  SET `类型`=@类型, `位置`=@位置";
                        sqlTxt += " WHERE `商店` =@商店";
                    }
                    else
                    {
                        sqlTxt = $"INSERT INTO `商店配置` (`商店`,`类型`,`位置`) VALUES (@商店,@类型,@位置)";
                    }
                    list.Add(new object[2] { "@类型", typeMap });
                    list.Add(new object[2] { "@位置", map });
                    Sqlite.NonQuery(sqlTxt, list);

                    Update.Show(Update.EShow.商店, shop, show, show == "");
                }
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                MessageBox.Show(ex.ToString());
                return 1;
            }
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            return 0;
        }
        /// <summary>
        /// 任务
        /// </summary>
        /// <returns></returns>
        public static int Duties()
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
            int errorCheck = 0, numCount = 0, tNum = 0;
            try
            {
                state = 2;
                // 先进行处理, 将任务目标清理

                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    duties = arr[(int)ExcelType_v1.Duties.任务].Trim();
                    if (duties == "") continue;
                    // 每3000提交一次事务


                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
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

                numCount = 0;
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    duties = arr[(int)ExcelType_v1.Duties.任务].Trim();
                    if (duties == "") continue;

                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
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

                    if (Sqlite.Select("SELECT `任务` FROM `任务配置`  WHERE `任务`=@任务;", sList)) sqlTxt = "UPDATE `任务配置` SET `类型`=@类型 , `掉落` = @掉落 WHERE `任务` = @任务;";
                    else sqlTxt = "INSERT INTO `任务配置` (`任务`,`类型`,`掉落`) VALUES (@任务,@类型,@掉落);";
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
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");

            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }
        /// <summary>
        /// 副本
        /// </summary>
        /// <returns></returns>
        public static int Copy()
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
            int numCount = 0, errorCheck = 0;
            try
            {
                state = 2;
                // 先进行处理, 将副本地图清理
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    copy = arr[(int)ExcelType_v1.Copy.副本].Trim();
                    if (copy == "") continue;
                    // 每3000提交一次事务


                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
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
                numCount = 0;
                state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    copy = arr[(int)ExcelType_v1.Copy.副本].Trim();
                    if (copy == "") continue;

                    numCount++; // 执行的次数
                    Numwrite = i + 1; // 外部显示当前处理进度
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

                    if (Sqlite.Select("SELECT `ID` FROM `副本配置` WHERE `copy` = @copy", list)) sqlTxt = "UPDATE `副本配置` SET `type`=@type, `t_number`=@entries, `reward_1`=@rewardType, `c_drop`=@rewards, `sor`=@copySort,`limit_user`=@limitUser WHERE `copy` = @copy;";
                    else sqlTxt = "INSERT INTO `副本配置` (`copy`,`type`,`t_number`,`reward_1`,`c_drop`,`sor`,`limit_user`) VALUES (@copy,@type,@entries,@rewardType,@rewards,@copySort,@limitUser);";
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
            state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + numCount.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            if (errorCheck > 0)
                MessageBox.Show($"本次限制消耗共有：{errorCheck}处错误未写入", "提示");
            return 0;
        }

    }
}

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
    internal class Map
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

            long startTime = Stopwatch.GetTimestamp();

            int errorCheck = 0;
            string temp, ret;
            try
            {
                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("SELECT `id`,`map` FROM `config_map`", null, "map");

                for (var i = 0; i < arrWrite.Count; i++)
                {

                    string[] arr = [.. arrWrite[i]];
                    if (arr[(int)ExcelType_v1.Map.位置].Trim() == "") continue;
                    V1Write.Numwrite++;
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }

                    List<object[]> list =
                    [
                        ["@map", arr[(int)ExcelType_v1.Map.位置]]
                    ];

                    if (selectArr.Exists(t => t == arr[(int)ExcelType_v1.Map.位置]))
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
                        selectArr.Add(arr[(int)ExcelType_v1.Map.位置]);
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

                    string[] tarr = temp.Split(['\n']);
                    foreach (string item in tarr)
                    {
                        string[] yarr = item.Split(["[too]"], StringSplitOptions.None);
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

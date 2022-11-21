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
    internal class Equip
    {
        public static int V1()
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

            int testing,errorCheck = 0; ;
            string ret;
            try
            {
                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("装备配置", "SELECT `装备` FROM `装备配置`", null, "装备");
                for (int i = 0; i < arrWrite.Count; i++)
                {
                    V1Write.Numwrite++;
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }


                    string[] arr = arrWrite[i].ToArray();
                    List<object[]> list = new()
                    {
                        new object[2] { "@装备", arr[(int)ExcelType_v1.Equip.名称] }
                    };

                    if (selectArr.Exists(t => t == arr[(int)ExcelType_v1.Equip.名称]))
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
                        selectArr.Add(arr[(int)ExcelType_v1.Equip.名称]); // 已添加,后续有重复的视为修改
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

                    // 限制消耗
                    if (arr[(int)ExcelType_v1.Equip.装备限制消耗] != "")
                    {
                        ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Equip.装备限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.装备, arr[(int)ExcelType_v1.Equip.名称], arr[(int)ExcelType_v1.Equip.装备限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.装备, arr[(int)ExcelType_v1.Equip.名称], arr[(int)ExcelType_v1.Equip.装备限制消耗], true);
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

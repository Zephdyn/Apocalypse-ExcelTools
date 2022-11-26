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
using static excelTools.excel.ExcelType_v1;

namespace ExcelTools.data.write
{
    class Skill
    {
        public static int V1()
        {
            var arrWrite = Program.arrWrite;
            if (arrWrite.Count == 0) return -1;
            // 获取连接
            var con = Config.connection;
            // 开启事务
            var transaction = con.BeginTransaction();
            // sql操作
            SqliteCommand write = con.CreateCommand();
            // 开始时间
            long startTime = Stopwatch.GetTimestamp();

            string skill, ret;
            int errorCheck = 0;
            try
            {
                // 删除所有技能
                V1Write.state = 2;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    skill = arr[(int)ExcelType_v1.Skill.名称].Trim();
                    if (skill == "") continue;
                    // 技能类型 0回复 1伤害 2辅助 3BUFF 4被动
                    if (!int.TryParse(arr[(int)ExcelType_v1.Skill.类型], out int type)) type = 0;
                    if (type < 0 || type > 4) continue;

                    V1Write.Numwrite++;
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }


                    List<object[]> list = new()
                    {
                        new object[2] { "@skill", Code.Asm_TextTo16( skill) }
                    };

                    if (Sqlite.Select("SELECT 技能 FROM `技能配置` WHERE `技能` = @skill", list)) // 查询是否已有技能
                        Sqlite.NonQuery("DELETE  FROM `技能配置` WHERE `技能` = @skill", list);

                }

                V1Write.Numwrite = 0;
                V1Write.state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = arrWrite[i].ToArray();
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    skill = arr[(int)ExcelType_v1.Skill.名称].Trim();
                    if (skill == "") continue;
                    // 技能类型 0回复 1伤害 2辅助 3BUFF 4被动
                    if (!int.TryParse(arr[(int)ExcelType_v1.Skill.类型], out int type)) type = 0;
                    if (type < 0 || type > 4) continue;

                    V1Write.Numwrite++; // 执行的次数
                                        // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }
                    // 插入通用

                    // 描述
                    Update.Show(Update.EShow.技能, skill, arr[(int)ExcelType_v1.Skill.描述], false);
                    // 限制消耗
                    if (arr[(int)ExcelType_v1.Skill.限制消耗] != "")
                    {
                        ret = TString.LimitConfigCheck(arr[(int)ExcelType_v1.Skill.限制消耗]);
                        if (ret != "") errorCheck++;
                        else Update.Check(Update.ECheck.技能, skill, arr[(int)ExcelType_v1.Skill.限制消耗], false);
                    }
                    else Update.Check(Update.ECheck.技能, skill, arr[(int)ExcelType_v1.Skill.限制消耗], true);
                    // 查询
                    Update.SkillConfigInsert(skill, "5", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.查询], 1, 0, 0));
                    // 暴击
                    Update.SkillConfigInsert(skill, "6", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.暴击], 1, 0, 0));
                    // 冷却
                    Update.SkillConfigInsert(skill, "2", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.冷却时间], 1, 0, 0));
                    switch (type)
                    {
                        case 0:
                            Reply(skill, arr);
                            break;
                        case 1:
                            Hurt(skill, arr);
                            break;
                        case 2:
                            Assist(skill, arr);
                            break;
                        case 3:
                            BUFF(skill, arr);
                            break;
                        case 4:
                            Passive(skill, arr);
                            break;
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



        private static void Reply(string skill, string[] arr)// 回复
        {
            Update.SkillConfigInsert(skill, "0003", arr[(int)ExcelType_v1.Skill.生命回复]);
            Update.SkillConfigInsert(skill, "0004", arr[(int)ExcelType_v1.Skill.法力回复]);
            Update.SkillConfigInsert(skill, "0005", arr[(int)ExcelType_v1.Skill.饥饿回复]);

            Update.SkillConfigInsert(skill, "0006", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.复活技能], 1, 0, 0));
            Update.SkillConfigInsert(skill, "0007", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.队友释放], 1, 0, 0));
            // 持续时间>0 ,则视为持续技能
            Update.SkillConfigInsert(skill, "0008", TString.GetInt(arr[(int)ExcelType_v1.Skill.持续回复], 0) > 0 ? 1 : 0);
            Update.SkillConfigInsert(skill, "0009", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.持续回复], int.MaxValue, 0, 0));
        }
        private static void Hurt(string skill, string[] arr)// 伤害
        {


        }
        private static void Assist(string skill, string[] arr)// 辅助
        {


        }
        private static void BUFF(string skill, string[] arr)// 增减益
        {


        }
        private static void Passive(string skill, string[] arr)// 被动
        {

        }
    }
}

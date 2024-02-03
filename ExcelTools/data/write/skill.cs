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
                    string[] arr = [.. arrWrite[i]];
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    skill = arr[(int)ExcelType_v1.Skill.名称].Trim();
                    if (skill == "") continue;
                    // 技能类型 0回复 1伤害 2辅助 3BUFF 4被动
                    if (!int.TryParse(arr[(int)ExcelType_v1.Skill.类型], out int type)) continue;
                    if (type < 0 || type > 4) continue;

                    V1Write.Numwrite++;
                    // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();
                    }


                    List<object[]> list =
                    [
                        ["@skill", Code.Asm_TextTo16( skill)]
                    ];

                    if (Sqlite.Select("SELECT 技能 FROM `技能配置` WHERE `技能` = @skill", list)) // 查询是否已有技能
                        Sqlite.NonQuery("DELETE  FROM `技能配置` WHERE `技能` = @skill", list);

                }

                V1Write.Numwrite = 0;
                V1Write.state = 1;
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = [.. arrWrite[i]];
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    skill = arr[(int)ExcelType_v1.Skill.名称].Trim();
                    if (skill == "") continue;
                    // 技能类型 0回复 1伤害 2辅助 3BUFF 4被动
                    if (!int.TryParse(arr[(int)ExcelType_v1.Skill.类型], out int type)) continue;
                    if (type < 0 || type > 4) continue;

                    V1Write.Numwrite++; // 执行的次数
                                        // 每3000提交一次事务
                    if (i > 0 && i % 3000 == 0)
                    {
                        transaction.Commit();

                        transaction = con.BeginTransaction();

                        write = con.CreateCommand();
                    }
                    // ===========插入通用===========
                    // 类型
                    Update.SkillConfigInsert(skill, "1", type);
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
            #region 配置
            //' (0003) 内容:生命回复
            //' (0004) 内容:法力回复
            //' (0005) 内容:饥饿回复
            //' (0006) 内容:0非 1是 复活技能
            //' (0007) 内容:0非 1是 对队友释放
            //' (0008) 内容:0非 1是 持续恢复
            //' (0009) 内容:持续时间
            #endregion
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
            #region 配置
            //' (1000) 内容:技能伤害
            //' (1010) 内容:固定吸血 0关闭 1开启
            //' (1011) 内容:固定吸血数值
            //' (1020) 内容:百分比吸血 0关闭 1开启
            //' (1021) 内容:百分比吸血值
            //' (1022) 内容:吸血概率 百分比
            #endregion
            Update.SkillConfigInsert(skill, "1000", arr[(int)ExcelType_v1.Skill.伤害]);
            Update.SkillConfigInsert(skill, "1010", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.是否固定吸血], 1, 0, 0));
            Update.SkillConfigInsert(skill, "1011", arr[(int)ExcelType_v1.Skill.固定吸血]);
            Update.SkillConfigInsert(skill, "1020", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.是否百分比吸血], 1, 0, 0));
            Update.SkillConfigInsert(skill, "1021", TString.GetDoubleTidy(arr[(int)ExcelType_v1.Skill.百分比吸血], 100, 0, 0));
            Update.SkillConfigInsert(skill, "1022", arr[(int)ExcelType_v1.Skill.吸血概率]);
        }
        private static void Assist(string skill, string[] arr)// 辅助
        {
            #region 配置
            //' (2000) 内容:(0固定传送) (1指定随机传送) (2全图随机传送) (3增加经验) (4增加等级)
            //' (2001) 内容:技能效果
            #endregion
            // 效果类型 0-4
            if (!int.TryParse(arr[(int)ExcelType_v1.Skill.效果类型], out int type)) return;
            if (type < 1 || type > 5) return;
            Update.SkillConfigInsert(skill, "2000", type);
            Update.SkillConfigInsert(skill, "2001", arr[(int)ExcelType_v1.Skill.效果]);

        }
        private static void BUFF(string skill, string[] arr)// 增减益
        {
            #region 配置
            //' (3000) 内容:持续类型(0,临时秒数)
            //' (3001) 内容:增幅时间
            //' (3002) 内容:状态名称
            //' (3003) 内容：空或0,BUFF技能  1为DEBUFF技能
            //' (3007) 内容:0非 1是 对队友释放
            //' (3008) 内容:0非 1是 对自己释放 默认1
            //' (3009) 内容:0非 1是 对怪物释放
            //' (3010) 内容:0非 1是 嘲讽技能
            //' (3050) 内容:增幅攻击
            //' (3051) 内容:增幅防御
            //' (3052) 内容:增幅生命
            //' (3053) 内容:增幅法力
            //' (3054) 内容:增幅命中
            //' (3055) 内容:增幅闪避
            //' (3056) 内容:增幅暴击概率
            //' (3057) 内容:增幅暴击伤害
            //' (3100) 内容:掉率加成
            //' (3101) 内容:固定经验
            //' (3102) 内容:百分比经验
            //' (3500) 内容:货币固定加成
            //' (3600) 内容:货币百分比加成
            #endregion
            Update.SkillConfigInsert(skill, "3000", 0);
            Update.SkillConfigInsert(skill, "3001", arr[(int)ExcelType_v1.Skill.增幅时间]);
            Update.SkillConfigInsert(skill, "3002", arr[(int)ExcelType_v1.Skill.状态名称]);
            Update.SkillConfigInsert(skill, "3003", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.减益], 1, 0, 0));
            Update.SkillConfigInsert(skill, "3007", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.队友释放], 1, 0, 0));
            Update.SkillConfigInsert(skill, "3008", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.自己释放], 1, 0, 1));
            Update.SkillConfigInsert(skill, "3009", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.怪物释放], 1, 0, 0));
            Update.SkillConfigInsert(skill, "3010", TString.GetIntTidy(arr[(int)ExcelType_v1.Skill.嘲讽技能], 1, 0, 0));

            Update.SkillConfigInsert(skill, "3050", arr[(int)ExcelType_v1.Skill.攻击]);
            Update.SkillConfigInsert(skill, "3051", arr[(int)ExcelType_v1.Skill.防御]);
            Update.SkillConfigInsert(skill, "3052", arr[(int)ExcelType_v1.Skill.生命]);
            Update.SkillConfigInsert(skill, "3053", arr[(int)ExcelType_v1.Skill.法力]);
            Update.SkillConfigInsert(skill, "3054", arr[(int)ExcelType_v1.Skill.命中]);
            Update.SkillConfigInsert(skill, "3055", arr[(int)ExcelType_v1.Skill.闪避]);
            Update.SkillConfigInsert(skill, "3056", arr[(int)ExcelType_v1.Skill.暴击概率]);
            Update.SkillConfigInsert(skill, "3057", arr[(int)ExcelType_v1.Skill.暴击伤害]);

            Update.SkillConfigInsert(skill, "3100", arr[(int)ExcelType_v1.Skill.掉率]);
            Update.SkillConfigInsert(skill, "3101", arr[(int)ExcelType_v1.Skill.经验固定]);
            Update.SkillConfigInsert(skill, "3102", arr[(int)ExcelType_v1.Skill.经验百分比]);
            Update.SkillConfigInsert(skill, "3500", arr[(int)ExcelType_v1.Skill.货币固定]);
            Update.SkillConfigInsert(skill, "3600", arr[(int)ExcelType_v1.Skill.货币百分比]);
        }
        private static void Passive(string skill, string[] arr)// 被动
        {
            #region 配置
            //' (4000) 内容:1:被动类型(增加属性)
            //' (4002) 内容:状态名称
            //' (4003) 内容:可用上限,-1则为无限制
            //' (4050) 内容:增幅攻击
            //' (4051) 内容:增幅防御
            //' (4052) 内容:增幅生命
            //' (4053) 内容:增幅法力
            //' (4054) 内容:增幅命中
            //' (4055) 内容:增幅闪避
            //' (4056) 内容:增幅暴击概率
            //' (4057) 内容:增幅暴击伤害
            #endregion
            Update.SkillConfigInsert(skill, "4000", 1);
            Update.SkillConfigInsert(skill, "4002", arr[(int)ExcelType_v1.Skill.状态名称]);
            Update.SkillConfigInsert(skill, "4003", TString.GetInt(arr[(int)ExcelType_v1.Skill.叠加上限], 0));

            Update.SkillConfigInsert(skill, "4050", arr[(int)ExcelType_v1.Skill.攻击]);
            Update.SkillConfigInsert(skill, "4051", arr[(int)ExcelType_v1.Skill.防御]);
            Update.SkillConfigInsert(skill, "4052", arr[(int)ExcelType_v1.Skill.生命]);
            Update.SkillConfigInsert(skill, "4053", arr[(int)ExcelType_v1.Skill.法力]);
            Update.SkillConfigInsert(skill, "4054", arr[(int)ExcelType_v1.Skill.命中]);
            Update.SkillConfigInsert(skill, "4055", arr[(int)ExcelType_v1.Skill.闪避]);
            Update.SkillConfigInsert(skill, "4056", arr[(int)ExcelType_v1.Skill.暴击概率]);
            Update.SkillConfigInsert(skill, "4057", arr[(int)ExcelType_v1.Skill.暴击伤害]);

        }
    }
}

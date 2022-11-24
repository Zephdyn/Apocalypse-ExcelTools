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

            string skill;
            int errorCheck = 0;
            try
            {
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

                    switch (type)
                    {
                        case 0:
                            errorCheck += Reply(arr);
                            break;
                        case 1:
                            errorCheck += Hurt(arr);
                            break;
                        case 2:
                            errorCheck += Assist(arr);
                            break;
                        case 3:
                            errorCheck += BUFF(arr);
                            break;
                        case 4:
                            errorCheck += Passive(arr);
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


        // 返回值>0 则限制消耗填入异常
        private static int Reply(string[] arr)// 回复
        {

            return 0;
        }
        private static int Hurt(string[] arr)// 伤害
        {

            return 0;
        }
        private static int Assist(string[] arr)// 辅助
        {

            return 0;
        }
        private static int BUFF(string[] arr)// 增减益
        {

            return 0;
        }
        private static int Passive(string[] arr)// 被动
        {
            return 0;
        }
    }
}

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
    internal class Goods
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
            // 当前数量置0
            V1Write.Numwrite = 0;
            string ret, sqlTxt;
            int errorCheck = 0;
            try
            {
                // 0(null) 1(num) 2(double) 3(string)
                V1Write.state = 1;
                // 读取所有数据
                List<string> selectArr = Sqlite.SelectAll("SELECT `道具` FROM `道具配置`", null, "道具");
                for (var i = 0; i < arrWrite.Count; i++)
                {
                    // 删除第一个
                    string[] arr = [.. arrWrite[i]];
                    arr = arr.Where((val, idx) => idx > 0).ToArray();
                    if (arr[(int)ExcelType_v1.Goods.物品].Trim() == "") continue;


                    /*格式:
                     *1消耗类道具 2礼包类道具 3vip类道具(vip卡) 4vip类道具(vip积分卡) 5技能书 6坐骑蛋 7坐骑喂养 8亲密道具 9数据清空 10鱼饵 11任务 12免死亡惩罚 13材料
                     */
                    // 类型必须为1-13
                    if (!int.TryParse(arr[(int)ExcelType_v1.Goods.类型], out int goodsType)) continue;
                    if (goodsType < 1 || goodsType > 13) continue;
                    V1Write.Numwrite++; // 当前处理进度
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
                    else if (goodsType == 9 || goodsType == 12 || goodsType==13)
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
                    List<object[]> list =
                    [
                        ["@goods", arr[(int)ExcelType_v1.Goods.物品]],

                    ];

                    if (selectArr.Exists(t => t == arr[(int)ExcelType_v1.Goods.物品]))
                    {
                        sqlTxt = "UPDATE `道具配置` SET `类型` = @type ,`名称` = @value,`绑定` = @bind,'批量' = @batch,`战斗` = @fight WHERE `道具` = @goods";
                    }
                    else
                    {
                        sqlTxt = "INSERT INTO `道具配置` (`道具`,`类型`,`名称`,`绑定`,'批量',`战斗`) VALUES (@goods, @type, @value, @bind, @batch, @fight)";
                        selectArr.Add(arr[(int)ExcelType_v1.Goods.物品]);
                    }

                    list =
                    [
                        ["@goods", arr[(int)ExcelType_v1.Goods.物品]],
                        ["@type", goodsType],
                        ["@value", goodsName],
                        ["@bind", goodsBind],
                        ["@batch", goodsBatch],
                        ["@fight", goodsFight]

                    ];
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

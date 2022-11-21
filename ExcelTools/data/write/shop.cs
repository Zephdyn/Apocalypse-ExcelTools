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
    internal class Shop
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

            string map, priceName, show, stock, intro, sqlTxt, shop;
            int nNum, oNum;
            try
            {
                // 0(null) 1(num) 2(double) 3(string)
                V1Write.state = 1;
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

                    V1Write.Numwrite++; // 执行的次数
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
            V1Write.state = 0;
            transaction.Commit();

            string stime = ((Stopwatch.GetTimestamp() - startTime) / (double)Stopwatch.Frequency).ToString("0.00");
            MessageBox.Show("执行数据量：" + V1Write.Numwrite.ToString() + "\r\n耗时：" + stime + " s", "写入完毕");
            return 0;
        }

    }
}

using excelTools.excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelTools.assist
{

    class Code
    {
        /// <summary>
        /// 将16进制转为字符串 [编码为gb2312]
        /// </summary>
        /// <param name="str">要转换的文本</param>
        /// <returns></returns>
        public static string Asm_16ToText(string str)
        {
            if (str == null)
                return "";
            str = str.Replace(",", "");
            str = str.Replace("\n", "");
            str = str.Replace("\\", "");
            str = str.Replace(" ", "");
            if (str.Length % 2 != 0)
            {
                str += "20";//空格
            }
            // 需要将 hex 转换成 byte 数组。
            byte[] bytes = new byte[str.Length / 2];

            for (int i = 0; i < bytes.Length; i++)
            {
                try
                {
                    // 每两个字符是一个 byte。
                    bytes[i] = byte.Parse(str.Substring(i * 2, 2),
                    System.Globalization.NumberStyles.HexNumber);
                }
                catch
                {
                    // Rethrow an exception with custom message.
                    // throw new ArgumentException("hex is not a valid hex number!", "hex");
                }
            }
            System.Text.Encoding chs = System.Text.Encoding.GetEncoding("gb2312");
            return chs.GetString(bytes);
        }


        /// <summary>
        /// <函数：Decode>
        /// 将字符串转为16进制 [编码为gb2312]
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string Asm_TextTo16(string str)
        {
            /*if ((strDecode.Length % 2) != 0)
            {
                strDecode += " ";//空格
                                 //throw new ArgumentException("s is not valid chinese string!");
            }*/
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            System.Text.Encoding chs = System.Text.Encoding.GetEncoding("GB2312");
            byte[] bytes = chs.GetBytes(str);
            string ret = "";
            for (int i = 0; i < bytes.Length; i++)
            {
                // 处理 \r\n换行
                if (i + 1 <= bytes.Length)
                {
                    if (bytes[i] == 13 && bytes[i + 1] == 10)
                    {

                        ret += "0D0A";
                        i++;
                        continue;
                    }
                }
                // 处理 \n换行
                if (bytes[i] == 10)
                {
                    ret += "0D0A";
                    continue;
                }
                // 处理 \r换行
                if (bytes[i] == 13)
                {
                    ret += "0D0A";
                    continue;
                }
                ret += string.Format("{0:X000}", bytes[i]);
            }
            ret = ret.ToLower();
            return ret.ToUpper();
        }



    }

    /// <summary>
    /// 对字符串进行处理
    /// </summary>
    class TString
    {
        /// <summary>
        /// 转整数
        /// </summary>
        /// <param name="text">原文本</param>
        /// <param name="def">转换失败,返回默认值</param>
        /// <returns></returns>
        public static int GetInt(string text, int def)
        {
            if (!int.TryParse(text, out int num)) num = def;
            return num;
        }
        /// <summary>
        /// 转整数,并整理
        /// </summary>
        /// <param name="text"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <param name="def"></param>
        /// <returns></returns>
        public static int GetIntTidy(string text, int max, int min, int def)
        {
            if (!int.TryParse(text, out int num)) num = def;
            return TNumber.IntTidy(num, max, min, def);
        }
        /// <summary>
        /// 转双精度小数,并整理
        /// </summary>
        /// <param name="text"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <param name="def"></param>
        /// <returns></returns>
        public static double GetDoubleTidy(string text, double max, double min, double def)
        {
            if (!double.TryParse(text, out double num)) num = def;
            return TNumber.DoubleTidy(num, max, min, def);
        }
        /// <summary>
        /// 获取时间
        /// </summary>
        /// <param name="type">1年到秒,</param>
        /// <returns>返回文本时间</returns>
        public static string GetTime(int type)
        {
            return type switch
            {
                1 => DateTime.Now.ToString("yyyy年M月d日h小时m分s秒"),
                _ => "",
            };
        }
        /// <summary>
        /// 判断字符串是否为正整数
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool CheckStringJustInt(string str)
        {
            if (!int.TryParse(str, out int num)) num = 0;
            return num > 0;
        }
        /// <summary>
        /// 删除空行,变成[too]分割
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string HandingBank(string str)
        {
            string[] sArr = str.Split(['\n']);
            string temp = "";
            foreach (var sTemp in sArr)
            {
                if (sTemp == "") continue;
                temp += sTemp + "[too]";
            }
            if (temp != "") temp = temp[0..^5];
            return temp;
        }
        /// <summary>
        /// 限制消耗检测, 返回错误信息
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string LimitConfigCheck(string str)
        {
            if (str == "") return "";
            int state;
            string[] item = new string[2];
            string[] sArr = str.Split(['\n']);
            foreach (var temp in sArr)
            {
                state = 0;
                string[] vs = temp.Split(["[too]" ], StringSplitOptions.None);
                if (vs.Length != 3) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 格式不符合";
                if (vs[0].Length == 0) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 功能类型不能为空";
                if (vs[1] != "0" && vs[1] != "1" && vs[1] != "2") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为0/1/2";
                switch (vs[0])
                {
                    case "性别":
                        if (vs[1] != "0") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为0";
                        if (vs[2] != "男" && vs[2] != "女") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为男/女";
                        state = 1;
                        break;
                    case "任务":
                        if (vs[1] == "2") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型不能为0";
                        state = 1;
                        break;
                    case "NPC":
                        if (vs[1] != "0") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为0";
                        state = 1;
                        break;
                    case "地图":
                        if (vs[1] != "0") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为0";
                        state = 1;
                        break;
                    case "职业":
                        if (vs[1] != "0") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型应为0";
                        state = 1;
                        break;
                    case "成就":
                        if (vs[1] == "1") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型不能为1";
                        state = 1;
                        break;
                    case "技能":
                        if (vs[1] == "1") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型不能为1";
                        state = 1;
                        break;
                    case "装备":
                        if (vs[1] != "0") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型必须为0";
                        if (vs[2].Trim() == "") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 装备名称不能为空";
                        state = 1;
                        break;
                    case "VIP等级":
                        if (vs[1] == "1") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型不能为1";
                        if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: VIP等级必须为正整数";
                        state = 1;
                        break;
                    case "等级":
                        if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 等级必须为正整数:";
                        state = 1;
                        break;
                    case "经验":
                        if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 经验必须为正整数";
                        state = 1;
                        break;
                }
                // 防止substring异常
                if (vs[0].Length > 2)
                {
                    item[0] = vs[0][..2];
                    item[1] = vs[0][2..];
                    switch (item[0])
                    {
                        case "属性":
                            if (item[1] == "") break;
                            if (vs[1] == "1")
                                if (item[1] != "生命" && item[1] != "法力" && item[1] != "饥饿") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型为1时,必须为生命/法力/饥饿";
                            if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 属性数值必须为正整数";
                            state = 1;
                            break;
                        case "钱庄":
                            if (item[1] == "") break;
                            if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 属性数值必须为正整数";
                            state = 1;
                            break;
                        case "坐骑":
                            if (item[1] == "") break;
                            if (vs[1] == "2") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 操作类型不能为2";
                            if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 坐骑等级必须为正整数";
                            state = 1;
                            break;
                        case "部位":
                            if (item[1] == "") break;
                            if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 强化等级必须为正整数";
                            state = 1;
                            break;
                        case "货币":
                            if (item[1] != "" && item[1] != "2" && item[1] != "3") return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 货币必须为(货币,货币2,货币3)";
                            if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 货币数量必须为正整数";
                            state = 1;
                            break;
                    }
                }
                if (state == 0) // 说明为物品
                {
                    if (!CheckStringJustInt(vs[2])) return $"错误内容: {temp} {System.Environment.NewLine}错误信息: 物品数量必须为正整数";
                }

            }

            return "";
        }
    }

    /// <summary>
    /// 对数值进行处理
    /// </summary>
    class TNumber
    {
        /// <summary>
        /// 整理整数
        /// </summary>
        /// <param name="num">当前</param>
        /// <param name="max">上限</param>
        /// <param name="main">下限</param>
        /// <param name="def">超过上限,低于下限则为默认</param>
        /// <returns></returns>
        public static int IntTidy(int num, int max, int min, int def)
        {
            if (num > max) return def;
            if (num < min) return def;
            return num;
        }
        /// <summary>
        /// 整理双精度小数
        /// </summary>
        /// <param name="num">当前</param>
        /// <param name="max">上限</param>
        /// <param name="main">下限</param>
        /// <param name="def">超过上限,低于下限则为默认</param>
        /// <returns></returns>
        public static double DoubleTidy(double num, double max, double min, double def)
        {
            if (num > max) return def;
            if (num < min) return def;
            return num;
        }
        /// <summary>
        /// 字符串转整数
        /// </summary>
        /// <param name="str">字符串</param>
        /// <param name="def">默认</param>
        /// <returns></returns>
        public static int IntTryParse(string str,int def)
        {
            if (!int.TryParse(str, out int num)) num = def;
            return num;
        }
        /// <summary>
        /// 字符串转长整数
        /// </summary>
        /// <param name="str">字符串</param>
        /// <param name="def">默认</param>
        /// <returns></returns>
        public static Int64 Int64TryParse(string str, Int64 def)
        {
            if (!Int64.TryParse(str, out Int64 num)) num = def;
            return num;
        }
        /// <summary>
        /// 字符串转双精度小数
        /// </summary>
        /// <param name="str">字符串</param>
        /// <param name="def">默认</param>
        /// <returns></returns>
        public static double DoubleTryParse(string str, double def)
        {
            if (!double.TryParse(str, out double num)) num = def;
            return num;
        }
    }
}

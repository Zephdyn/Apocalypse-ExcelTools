using System;
using System.Collections.Generic;

using Microsoft.Data.Sqlite;
using excelTools.assist;
using System.Windows;

namespace excelTools.data
{
    class Config
    {
        public static SqliteConnection connection;
        /// <summary>
        /// 打开数据库
        /// </summary>
        /// <param name="filepath">数据库路径</param>
        /// <returns></returns>
        public static bool Open(string filepath)
        {
            try
            {
                // 只读写已经存在的库, 避免路径写错时静默创建出一个空库
                var builder = new SqliteConnectionStringBuilder
                {
                    DataSource = filepath,
                    Mode = SqliteOpenMode.ReadWrite
                };
                // 尝试打开数据库
                connection = new SqliteConnection(builder.ToString());
                connection.Open();
                DbSchema.Reset();
                // 读取数据库版本
                return true;
            }
            catch
            {
                Close();
                return false;
            }
        }
        /// <summary>
        /// 关闭数据库
        /// </summary>
        public static void Close()
        {
            DbSchema.Reset();
            if (connection == null) return;
            try
            {
                connection.Close();
                connection.Dispose();
            }
            finally
            {
                connection = null;
            }
        }
        /// <summary>
        /// 显示数据库连接状态
        /// </summary>
        /// <returns></returns>
        public static int Status()
        {
            if (connection == null) return -1;
            return (int)connection.State;
        }
        /// <summary>
        /// 获取数据库版本, 返回一个数组, [0]为版本 [1]为版本号 [2]出错信息,正常时返回空
        /// </summary>
        /// <returns></returns>
        public static string[] Ver()
        {
            string[] ret = new string[3];

            var command = connection.CreateCommand();
            command.CommandText = @"SELECT id,c_key,value FROM `record_ini` WHERE file = $file AND sec =$sec AND (c_key = $c_key OR c_key = $c_key2)";

            command.Parameters.AddWithValue("$file", Code.Asm_TextTo16("系统配置"));
            command.Parameters.AddWithValue("$sec", Code.Asm_TextTo16("数据库"));
            command.Parameters.AddWithValue("$c_key", Code.Asm_TextTo16("版本"));
            command.Parameters.AddWithValue("$c_key2", Code.Asm_TextTo16("版本-New"));
            try
            {
                var reader = command.ExecuteReader();
                var write = connection.CreateCommand();
                while (reader.Read())
                {
                    var c_key = reader.GetString(1);
                    if (c_key == Code.Asm_TextTo16("版本")) ret[1] = Code.Asm_16ToText(reader.GetString(2));
                    else ret[0] = Code.Asm_16ToText(reader.GetString(2));
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                // 读取失败;
                ret[2] = ex.ToString();

            }

            return ret;

        }

    }

    class Update
    {
        /// <summary>
        /// 描述
        /// </summary>
        public enum EShow
        {
            //类型: 1装备 2道具 3技能 4地图 5怪物 6任务 7NPC 8职业 9成就  10事件 11商店  12道具提示 13坐骑描述 14合成 15婚宴描述
            装备 = 1,
            物品 = 2,
            技能 = 3,
            地图 = 4,
            怪物 = 5,
            任务 = 6,
            NPC = 7,
            职业 = 8,
            成就 = 9,
            事件 = 10,
            商店 = 11,
            物品提示 = 12,
            坐骑描述 = 13,
            合成 = 14,
            婚宴描述 = 15
        }
        /// <summary>
        /// 限制消耗
        /// </summary>
        public enum ECheck
        {
            地图 = 1,
            任务 = 2,
            物品 = 3,
            单人副本 = 15,
            技能 = 17,
            装备 = 19,
            /*
            ' 1. 地图
            ' 2. 任务
            ' 3. 物品
            ' 4. 职业
            ' 5. 坐骑(调用自己的,暂时不用此处)
            ' 6. 合成
            ' 7. 钱庄(调用自己的,暂时不用此处)
            ' 8. 打造
            ' 9. 婚宴(调用自己的,暂时不用此处)
            ' 10.邀请(调用自己的,暂时不用此处)
            ' 11.强化(调用自己的,暂时不用此处)
            ' 12.NPC对话(调用自己的,暂时不用此处)
            ' 13.职业遗忘(调用自己的,暂时不用此处)
            ' 14.无尽塔(暂无)
            ' 15.副本地图
            ' 16.鱼竿修复
            ' 17.技能
            ' 18.NPC 技能学习( 格式 NPC名称|技能名称
            ' 19.装备
            ' 100.属性修改
            */
        }
        /// <summary>
        /// 拓展配置
        /// </summary>
        public enum EPluginConfig
        {
            地图 = 6,
            物品 = 7
        }
        /// <summary>
        /// 自定义属性配置
        /// </summary>
        public enum EExtendStats
        {
            等级 = 1,
            装备 = 2,
            坐骑 = 3,
            怪物 = 4
        }
        /// <summary>
        /// 描述
        /// </summary>
        /// <param name="type">类型</param>
        /// <param name="name">名称</param>
        /// <param name="value">内容</param>
        /// <param name="delete">是否要删除</param>
        /// <param name="insert">是否强制插入</param>
        /// <returns></returns>
        public static bool Show(EShow type, string name, string value, bool delete)
        {
            //SQlite.查找(数据库_句柄_系统, “描述配置”, “类型 = '” ＋ Asm_TextTo16 (到文本 (type)) ＋ “' and 名称 = '” ＋ Asm_TextTo16 (name) ＋ “'”, “描述”
            //类型: 1装备 2道具 3技能 4地图 5怪物 6任务 7NPC 8职业 9成就  10事件 11商店  12道具提示 13坐骑描述 14合成 15婚宴描述
            var con = Config.connection;
            // sql操作
            SqliteCommand command = con.CreateCommand();

            List<object[]> list =
            [
                ["@类型", Code.Asm_TextTo16(((int)type).ToString())],
                ["@名称", Code.Asm_TextTo16(name)]
            ];
            if (Sqlite.Select("SELECT 描述 FROM 描述配置 WHERE 类型 = @类型 AND 名称 = @名称;", list))
            {
                if (delete)
                {
                    command.CommandText = @"
                    DELETE FROM 描述配置
                    WHERE 类型 = @类型 AND 名称 = @名称;
                    ";
                }
                else
                {
                    command.CommandText = @"
                    UPDATE 描述配置
                    SET 描述 = @描述
                    WHERE 类型 = @类型 AND 名称 = @名称;
                ";
                }
            }
            else
            {
                if (delete) return true;
                command.CommandText = @"
                    INSERT INTO 描述配置
                    (类型,名称,描述) VALUES
                    (@类型,@名称,@描述);
                    ";
            }

            command.Parameters.AddWithValue("@类型", Code.Asm_TextTo16(((int)type).ToString()));
            command.Parameters.AddWithValue("@名称", Code.Asm_TextTo16(name));
            if (delete == false) command.Parameters.AddWithValue("@描述", Code.Asm_TextTo16(value));
            command.ExecuteNonQuery();
            command.Parameters.Clear();
            return true;
        }
        /// <summary>
        /// 限制消耗
        /// </summary>
        /// <param name="type"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="delete"></param>
        /// <returns></returns>
        public static bool Check(ECheck type, string name, string value, bool delete)
        {
            var con = Config.connection;
            // sql操作
            SqliteCommand command = con.CreateCommand();
            if (value == "") delete = true;
            List<object[]> list = [
                ["@ID", ((int)type).ToString()],
                ["@name", Code.Asm_TextTo16(name)]
            ];
            if (Sqlite.Select("SELECT ID FROM `消耗限制` WHERE `ID` = @ID AND `name` = @name;", list))
            {
                if (delete)
                {
                    command.CommandText = @"
                    DELETE FROM `消耗限制`
                    WHERE `ID` = @ID AND `name` = @name;
                    ";
                }
                else
                {
                    command.CommandText = @"
                    UPDATE `消耗限制`
                    SET `msg` = @msg
                    WHERE `ID` = @ID AND `name` = @name;
                ";
                }
            }
            else
            {
                if (delete) return true;
                command.CommandText = @"
                    INSERT INTO `消耗限制`
                    (ID,name,msg) VALUES
                    (@ID,@name,@msg);
                    ";
            }

            command.Parameters.AddWithValue("@ID", ((int)type).ToString());
            command.Parameters.AddWithValue("@name", Code.Asm_TextTo16(name));
            if (delete == false) command.Parameters.AddWithValue("@msg", Code.Asm_TextTo16(value));
            command.ExecuteNonQuery();
            command.Parameters.Clear();
            return true;
        }
        /// <summary>
        /// 拓展配置
        /// </summary>
        /// <param name="id">ID</param>
        /// <param name="sec">名称</param>
        /// <param name="type">类型</param>
        /// <param name="msg">内容</param>
        /// <param name="delete">是否删除</param>
        /// <returns></returns>
        public static bool PluginConfig(EPluginConfig id, string sec, string type, string msg, bool delete)
        {

            var con = Config.connection;
            // sql操作
            SqliteCommand command = con.CreateCommand();
            if (msg == "") delete = true;
            List<object[]> list = [
               ["@ID", Code.Asm_TextTo16(((int)id).ToString())],
                ["@sec", Code.Asm_TextTo16(sec)],
                ["@type", Code.Asm_TextTo16(type)]
            ];
            if (Sqlite.Select("SELECT ID FROM `拓展配置` WHERE `ID` = @ID AND `sec` = @sec AND `type` = @type;", list))
            {
                if (delete)
                {
                    command.CommandText = @"
                    DELETE FROM `拓展配置`
                    WHERE `ID` = @ID AND `sec` = @sec AND `type` = @type;
                    ";
                }
                else
                {
                    command.CommandText = @"
                    UPDATE `拓展配置`
                    SET `msg` = @msg
                    WHERE `ID` = @ID AND `sec` = @sec AND `type` = @type;
                    ";
                }
            }
            else
            {
                if (delete) return true;
                command.CommandText = @"
                    INSERT INTO `拓展配置`
                    (ID,sec,type,msg) VALUES
                    (@ID,@sec,@type,@msg);
                    ";
            }

            command.Parameters.AddWithValue("@ID", Code.Asm_TextTo16(((int)id).ToString()));
            command.Parameters.AddWithValue("@sec", Code.Asm_TextTo16(sec));
            command.Parameters.AddWithValue("@type", Code.Asm_TextTo16(type));
            if (delete == false) command.Parameters.AddWithValue("@msg", Code.Asm_TextTo16(msg));
            command.ExecuteNonQuery();
            command.Parameters.Clear();
            return true;
        }
        /// <summary>
        /// 数据配置
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sec"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="delete"></param>
        /// <returns></returns>
        public static bool DbConfig(string file, string sec, string key, string value, bool delete)
        {
            string sqlTxt;
            List<object[]> list =
            [
                ["@file", Code.Asm_TextTo16(file)],
                ["@sec", Code.Asm_TextTo16(sec)],
                ["@key", Code.Asm_TextTo16(key)]
            ];
            if (Sqlite.Select("SELECT id FROM `record_ini` WHERE `file` = @file AND `sec` = @sec AND `c_key` = @key;", list))
            {
                if (delete)
                {
                    sqlTxt = "DELETE FROM `record_ini` WHERE  `file` = @file AND `sec` = @sec AND `c_key` = @key;";
                }
                else
                {
                    sqlTxt = "UPDATE `record_ini` SET `value` = @value WHERE  `file` = @file AND `sec` = @sec AND `c_key` = @key;";
                }
            }
            else
            {
                if (delete) return true;
                sqlTxt = "INSERT INTO `record_ini` (`file`,`sec`,`c_key`,`value`) VALUES (@file,@sec,@key,@value);";
            }

            list.Add(["@value", Code.Asm_TextTo16(value)]);
            return Sqlite.NonQuery(sqlTxt, list) > 0;
        }
        /// <summary>
        /// 技能配置插入
        /// </summary>
        /// <param name="skill">技能名称</param>
        /// <param name="type">类型</param>
        /// <param name="value">值</param>
        /// <returns></returns>
        public static bool SkillConfigInsert(string skill, string type, object value)
        {
            // 为空不插入
            if (value.ToString() == "") return false;
            //CREATE TABLE 技能配置 (技能 TEXT, 类型 TEXT, 内容 TEXT);
            var con = Config.connection;
            // sql操作
            SqliteCommand command = con.CreateCommand();


            command.CommandText = "INSERT INTO 技能配置 (技能,类型,内容) VALUES (@skill,@type,@value);";


            command.Parameters.AddWithValue("@skill", Code.Asm_TextTo16(skill));
            command.Parameters.AddWithValue("@type", Code.Asm_TextTo16(type));
            command.Parameters.AddWithValue("@value", Code.Asm_TextTo16(value.ToString()));
            command.ExecuteNonQuery();
            command.Parameters.Clear();
            return true;
        }

        public static void ExtendStatsInsert(SqliteCommand write, EExtendStats type, string name, string stats, Int64 number)
        {
            DbSchema.Ensure();
            write.CommandText = "INSERT INTO 自定义属性_数据 (类型,名称,属性,数值) VALUES (@类型,@名称,@属性,@数值);";

            write.Parameters.AddWithValue("@类型", (int)type);
            write.Parameters.AddWithValue("@名称", name);
            write.Parameters.AddWithValue("@属性", stats);
            write.Parameters.AddWithValue("@数值", number);
            write.ExecuteNonQuery();
            write.Parameters.Clear();
        }

        public static void ExtendStatsClear(SqliteCommand write, EExtendStats type, string name)
        {
            DbSchema.Ensure();
            write.CommandText = "DELETE FROM 自定义属性_数据 WHERE 类型=@类型 AND 名称=@名称;";

            write.Parameters.AddWithValue("@类型", (int)type);
            write.Parameters.AddWithValue("@名称", name);
            write.ExecuteNonQuery();
            write.Parameters.Clear();
        }
    }
}

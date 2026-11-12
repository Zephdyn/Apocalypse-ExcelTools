using System;
using System.Collections.Generic;

namespace excelTools.data
{
    /// <summary>
    /// 数据库结构检查
    /// 游戏 4.0.0 的数据库更新脚本(数据库更新_400_0_系统_0.txt)新增了下面这些表,
    /// 缺失时装备 / 怪物的自定义属性无法写入(报 no such table)
    /// </summary>
    internal static class DbSchema
    {
        /// <summary>
        /// 4.0.0 新增的表, 建表语句与游戏更新脚本保持一致
        /// </summary>
        private static readonly (string Name, string Sql)[] V400Tables =
        [
            ("自定义属性_数据", @"CREATE TABLE IF NOT EXISTS 自定义属性_数据 (
    id INTEGER      PRIMARY KEY AUTOINCREMENT
                    UNIQUE,
    类型 INTEGER,
    名称 TEXT,
    属性 VARCHAR (32),
    数值 BIGINT
);"),
            ("自定义属性_名", @"CREATE TABLE IF NOT EXISTS 自定义属性_名 (
    id   INTEGER      PRIMARY KEY ASC AUTOINCREMENT,
    属性 VARCHAR (32) UNIQUE
);"),
            ("自定义公式_伤害", @"CREATE TABLE IF NOT EXISTS 自定义公式_伤害 (
    id   INTEGER        PRIMARY KEY AUTOINCREMENT,
    方式   VARCHAR (32),
    权重   INTEGER,
    对象类型 INTEGER,
    对象   VARCHAR (128),
    公式   VARCHAR (1024) 
);"),
        ];

        // 结构是否已通过检查
        private static bool _verified;

        /// <summary>
        /// 获取缺失的表名, 全部存在时返回空集合
        /// </summary>
        public static List<string> GetMissingTables()
        {
            List<string> missing = [];
            foreach (var table in V400Tables)
            {
                if (!Sqlite.TableExists(table.Name)) missing.Add(table.Name);
            }
            return missing;
        }

        /// <summary>
        /// 重新检查并记录结构状态
        /// </summary>
        /// <returns>返回缺失的表名, 全部存在时返回空集合</returns>
        public static List<string> Check()
        {
            List<string> missing = GetMissingTables();
            _verified = missing.Count == 0;
            return missing;
        }

        /// <summary>
        /// 创建缺失的表
        /// </summary>
        /// <param name="error">失败原因</param>
        /// <returns>是否全部创建成功</returns>
        public static bool CreateMissingTables(out string error)
        {
            error = "";
            List<string> missing = GetMissingTables();
            if (missing.Count == 0)
            {
                _verified = true;
                return true;
            }
            try
            {
                foreach (var table in V400Tables)
                {
                    if (!missing.Contains(table.Name)) continue;
                    Sqlite.NonQuery(table.Sql, null);
                }
                // 建表后复查, 确认没有静默失败
                missing = GetMissingTables();
                if (missing.Count > 0)
                {
                    error = "创建后仍然缺少：" + string.Join("、", missing);
                    _verified = false;
                    return false;
                }
                _verified = true;
                return true;
            }
            catch (Exception ex)
            {
                _verified = false;
                error = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 检查结构是否就绪, 未就绪时抛出可读的异常(用于写入前的兜底)
        /// </summary>
        public static void Ensure()
        {
            if (_verified) return;
            List<string> missing = Check();
            if (missing.Count == 0) return;
            throw new InvalidOperationException(
                "当前数据库缺少 4.0.0 新增的表：" + string.Join("、", missing) +
                "。请重新连接数据库选择自动建表, 或把服务端数据库升级到 4.0.0 后再使用。");
        }

        /// <summary>
        /// 断开数据库时重置结构状态
        /// </summary>
        public static void Reset() => _verified = false;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.Sqlite;

namespace excelTools.data
{
    class Sqlite
    {
        /// <summary>
        /// 查询所有记录
        /// </summary>
        /// <param name="sql">语句</param>
        /// <param name="value">参数绑定</param>
        /// <param name="get">获取字段</param>
        /// <returns></returns>
        public static List<string> SelectAll( string sql, List<object[]> value, string get)
        {
            List<string> arr = [];
            var con = Config.connection;
            using SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            if (value != null)
            {
                for (var i = 0; i < value.Count; i++)
                {
                    command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
                }
            }
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                arr.Add(reader[get].ToString());
            }
            return arr;
        }
        /// <summary>
        /// 查询, 返回是否有此记录
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool Select(string sql, List<object[]> value)
        {
            var con = Config.connection;
            using SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            if (value != null)
            {
                for (var i = 0; i < value.Count; i++)
                {
                    command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
                }
            }
            using var reader = command.ExecuteReader();
            return reader.Read();
        }
        /// <summary>
        /// 判断表是否存在
        /// </summary>
        /// <param name="table">表名</param>
        /// <returns></returns>
        public static bool TableExists(string table)
        {
            var con = Config.connection;
            if (con == null) return false;
            using SqliteCommand command = con.CreateCommand();
            command.CommandText = "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = @name LIMIT 1;";
            command.Parameters.AddWithValue("@name", table);
            using var reader = command.ExecuteReader();
            return reader.Read();
        }
        /// <summary>
        /// 执行操作, 返回操作行数
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static int NonQuery(string sql, List<object[]> value)
        {

            var con = Config.connection;
            using SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            if (value != null)
            {
                for (var i = 0; i < value.Count; i++)
                {
                    command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
                }
            }
            return command.ExecuteNonQuery();

        }


        public static Int64 GetCount(string table, string where, List<object[]> value)// 获取表内记录
        {

            string sql = $"SELECT COUNT(*) AS NumberOfOrders FROM {table}";
            if (where != null && where != "")
            {
                sql += " WHERE " + where;
            }
            var con = Config.connection;
            using SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            if (value != null)
            {
                for (var i = 0; i < value.Count; i++)
                {
                    command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
                }
            }
            using var reader = command.ExecuteReader();
            reader.Read();
            Int64 count = (Int64)reader[0];
            return count;
        }
    }
}

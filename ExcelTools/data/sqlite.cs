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
        /// 查询, 返回是否有此记录
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool Select(string sql, List<object[]> value)
        {

            var con = Config.connection;
            SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            for (var i = 0; i < value.Count; i++)
            {
                command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
            }
            var reader = command.ExecuteReader();
            bool state = reader.Read();
            reader.Close();
            command.Dispose();
            return state;
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
            SqliteCommand command = con.CreateCommand();
            command.CommandText = sql;
            for (var i = 0; i < value.Count; i++)
            {
                command.Parameters.AddWithValue((string)value[i].GetValue(0), value[i].GetValue(1) ?? DBNull.Value);
            }
            return command.ExecuteNonQuery();

        }
    }
}

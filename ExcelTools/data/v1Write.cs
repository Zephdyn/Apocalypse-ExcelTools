using System;
using System.Collections.Generic;
using System.Linq;

using excelTools.excel;
using excelTools.assist;
using Microsoft.Data.Sqlite;
using System.Windows;
using System.Data;
using System.Diagnostics;

namespace excelTools.data
{
    internal class V1Write
    {
        /// <summary>
        /// 0未开启, 1 v1插入  2 清理
        /// </summary>
        public static int state;
        // 写入值
        public static int Numwrite;

        // 0正确 -1当前无数据 1写入失败

        // 装备
        public static int Equip()
            => ExcelTools.data.write.Equip.V1();
        // 地图
        public static int Map()
            => ExcelTools.data.write.Map.V1();
        // 怪物
        public static int Monster()
            => ExcelTools.data.write.Monster.V1();
        // 掉落
        public static int Drop()
            => ExcelTools.data.write.Drop.V1();
        // 物品
        public static int Goods()
            => ExcelTools.data.write.Goods.V1();
        // 商店
        public static int Shop()
             => ExcelTools.data.write.Shop.V1();
        // 任务
        public static int Duties()
             => ExcelTools.data.write.Duties.V1();
        // 副本
        public static int Copy()
            => ExcelTools.data.write.Copy.V1();
        // 技能
        public static int Skill()
            => ExcelTools.data.write.Skill.V1();

    }
}


namespace excelTools.excel
{
    /// <summary>
    /// excel读取类型
    /// </summary>
    class ExcelType
    {
        /// <summary>
        /// excel读取类型枚举
        /// </summary>
        public enum Read
        {
            equip,
            map,
            monster,
            drop,
            goods,
            shop,
            duties,
            copyOne
        }
        /// <summary>
        /// 读取可用列缓存
        /// </summary>
        public struct RCache
        {
            public int _rowId;  // 要读取列的ID
            public int _classID;// 要读取类的ID
        }
    }
    /// <summary>
    /// V1 的类型读取
    /// </summary>
    class ExcelTypeRead_v1
    {
        /// <summary>
        /// 装备
        /// </summary>
        public static readonly string[] equip = new string[]
        {
            "名称",
            "职业",
            "绑定",
            "部位",
            "等级",
            "生命",
            "法力",
            "攻击",
            "防御",
            "命中",
            "闪避",
            "暴击概率",
            "暴击伤害",
            "饥饿",
            "卸下",
            "技能",
            "描述",
            "装备限制消耗"
        };
        /// <summary>
        /// 地图
        /// </summary>
        public static readonly string[] map = new string[]
        {
            "位置",
            "前",
            "后",
            "左",
            "右",
            "上",
            "下",
            "描述",
            "安全区",
            "怪物",
            "饥饿减少",
            "限制消耗"
        };
        /// <summary>
        /// 怪物
        /// </summary>
        public static readonly string[] monster = new string[]
        {
            "怪物",
            "等级",
            "生命",
            "法力",
            "攻击",
            "防御",
            "闪避",
            "命中",
            "暴击概率",
            "暴击伤害",
            "技能",
            "掉落",
            "描述",
            "对外显示",
            "地图显示"
        };
        /// <summary>
        /// 掉落
        /// </summary>
        public static readonly string[] drop = new string[] {
            "掉落",
            "类型",
            "经验",
            "货币",
            "物品",
            "掉落地图"
        };
        /// <summary>
        /// 物品
        /// </summary>
        public static readonly string[] goods = new string[] {
            "物品",
            "类型",
            "名称",
            "绑定",
            "批量",
            "战斗",
            "冷却",
            "提示",
            "描述",
            "限制消耗"
        };
        /// <summary>
        /// 商店
        /// </summary>
        public static readonly string[] shop = new string[] {
            "商店",
            "位置",
            "类型",
            "描述",
            "名称",
            "库存",
            "价格类型",
            "物品名称",
            "单价",
            "刷新类型",
            "刷新几率",
            "介绍"
        };
        /// <summary>
        /// 任务
        /// </summary>
        public static readonly string[] duties = new string[] {
            "任务",
            "类型",
            "掉落",
            "描述",
            "交付对象",
            "任务系列",
            "下个任务",
            "任务目标",
            "放弃",
            "限制消耗"
        };
        /// <summary>
        /// 单人副本
        /// </summary>
        public static readonly string[] copy = new string[] {
            "副本",
            "副本类型",
            "进入次数",
            "奖励类型",
            "通过奖励",
            "副本排序",
            "限制人数",
            "限制消耗",
            "地图",
            "初始地图",
            "通关地图",
            "前",
            "后",
            "左",
            "右",
            "上",
            "下",
            "怪物",
            "数量",
            "掉落",
            "强制攻击"
        };

    }

    class ExcelType_v1
    {
        /// <summary>
        /// 装备
        /// </summary>
        public enum Equip
        {
            名称 = 1,
            职业 = 2,
            绑定 = 3,
            部位 = 4,
            等级 = 5,
            生命 = 6,
            法力 = 7,
            攻击 = 8,
            防御 = 9,
            命中 = 10,
            闪避 = 11,
            暴击概率 = 12,
            暴击伤害 = 13,
            饥饿 = 14,
            卸下 = 15,
            技能 = 16,
            描述 = 17,
            装备限制消耗=18,
        };
        /// <summary>
        /// 地图
        /// </summary>
        public enum Map
        {
            位置 = 1,
            前 = 2,
            后 = 3,
            左 = 4,
            右 = 5,
            上 = 6,
            下 = 7,
            描述 = 8,
            安全区 = 9,
            怪物 = 10,
            饥饿减少 = 11,
            限制消耗 = 12
        };
        /// <summary>
        /// 怪物
        /// </summary>
        public enum Monster
        {
            怪物,
            等级,
            生命,
            法力,
            攻击,
            防御,
            闪避,
            命中,
            暴击概率,
            暴击伤害,
            技能,
            掉落,
            描述,
            对外显示,
            地图显示
        };
        /// <summary>
        /// 掉落
        /// </summary>
        public enum Drop
        {
            掉落,
            类型,
            经验,
            货币,
            物品,
            掉落地图
        };
        /// <summary>
        /// 物品
        /// </summary>
        public enum Goods
        {
            物品,
            类型,
            名称,
            绑定,
            批量,
            战斗,
            冷却,
            提示,
            描述,
            限制消耗
        };
        /// <summary>
        /// 商店
        /// </summary>
        public enum Shop
        {
            商店,
            位置,
            类型,
            描述,
            名称,
            库存,
            价格类型,
            物品名称,
            单价,
            刷新类型,
            刷新几率,
            介绍
        };
        /// <summary>
        /// 任务
        /// </summary>
        public enum Duties
        {
            任务,
            类型,
            掉落,
            描述,
            交付对象,
            任务系列,
            下个任务,
            任务目标,
            放弃,
            限制消耗
        };
        /// <summary>
        /// 单人副本
        /// </summary>
        public enum Copy
        {        
            副本,
            副本类型,
            进入次数,
            奖励类型,
            通过奖励,
            副本排序,
            限制人数,
            限制消耗,
            地图,
            初始地图,
            通关地图,
            前,
            后,
            左,
            右,
            上,
            下,
            怪物,
            数量,
            掉落,
            强制攻击
        };
    }

}

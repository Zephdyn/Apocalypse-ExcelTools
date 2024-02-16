
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
        public static readonly string[] equip =
        [
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
            "装备限制消耗",
            "自定义属性"
        ];
        /// <summary>
        /// 地图
        /// </summary>
        public static readonly string[] map =
        [
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
        ];
        /// <summary>
        /// 怪物
        /// </summary>
        public static readonly string[] monster =
        [
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
            "地图显示",
            "自定义属性"
        ];
        /// <summary>
        /// 掉落
        /// </summary>
        public static readonly string[] drop = [
            "掉落",
            "类型",
            "经验",
            "货币",
            "物品",
            "掉落地图"
        ];
        /// <summary>
        /// 物品
        /// </summary>
        public static readonly string[] goods = [
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
        ];
        /// <summary>
        /// 商店
        /// </summary>
        public static readonly string[] shop = [
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
        ];
        /// <summary>
        /// 任务
        /// </summary>
        public static readonly string[] duties = [
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
        ];
        /// <summary>
        /// 副本
        /// </summary>
        public static readonly string[] copy = [
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
        ];
        /// <summary>
        /// 技能
        /// </summary>
        public static readonly string[] skill = [
            "名称",
            "类型",
            "冷却时间",
            "查询",
            "暴击",
            "限制消耗",
            "描述",
            "生命回复",
            "法力回复",
            "饥饿回复",
            "复活技能",
            "队友释放",
            "持续回复",
            "伤害",
            "是否固定吸血",
            "是否百分比吸血",
            "固定吸血",
            "百分比吸血",
            "吸血概率",
            "效果类型",
            "效果",
            "增幅类型",
            "增幅时间",
            "状态名称",
            "减益",
            "自己释放",
            "怪物释放",
            "嘲讽技能",
            "生命",
            "法力",
            "攻击",
            "防御",
            "命中",
            "闪避",
            "暴击概率",
            "暴击伤害",
            "掉率",
            "经验固定",
            "经验百分比",
            "货币固定",
            "货币百分比",
            "叠加上限"
        ];

    }

    class ExcelType_v1
    {
        /// <summary>
        /// 装备
        /// </summary>
        public enum Equip
        {
            名称 ,
            职业 ,
            绑定 ,
            部位 ,
            等级 ,
            生命 ,
            法力 ,
            攻击 ,
            防御 ,
            命中 ,
            闪避 ,
            暴击概率 ,
            暴击伤害 ,
            饥饿 ,
            卸下 ,
            技能 ,
            描述 ,
            装备限制消耗 ,
            自定义属性 ,
        };
        /// <summary>
        /// 地图
        /// </summary>
        public enum Map
        {
            位置 ,
            前 ,
            后 ,
            左 ,
            右 ,
            上 ,
            下 ,
            描述 ,
            安全区 ,
            怪物 ,
            饥饿减少 ,
            限制消耗 
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
            地图显示,
            自定义属性 ,
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
        /// 副本
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

        /// <summary>
        /// 技能
        /// </summary>
        public enum Skill
        {
            名称,
            类型,
            冷却时间,
            查询,
            暴击,
            限制消耗,
            描述,
            生命回复,
            法力回复,
            饥饿回复,
            复活技能,
            队友释放,
            持续回复,
            伤害,
            是否固定吸血,
            是否百分比吸血,
            固定吸血,
            百分比吸血,
            吸血概率,
            效果类型,
            效果,
            增幅类型,
            增幅时间,
            状态名称,
            减益,
            自己释放,
            怪物释放,
            嘲讽技能,
            生命,
            法力,
            攻击,
            防御,
            命中,
            闪避,
            暴击概率,
            暴击伤害,
            掉率,
            经验固定,
            经验百分比,
            货币固定,
            货币百分比,
            叠加上限
        };
    }

}

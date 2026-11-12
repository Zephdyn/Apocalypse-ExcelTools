# 开发环境

	推荐:
	IDE: visual studio 2026
	NET版本: 10.0

# 运行环境

	Windows 10 1607 及以上
	.NET 10 Desktop Runtime（或使用自包含方式发布）

# 发布

	单文件 exe（自包含, 客户机不用装 .NET, 约 71 MB）:
	dotnet publish ExcelTools\ExcelTools.csproj -c Release -r win-x64 -p:PublishSingleFile=true
	输出: ExcelTools\bin\Release\net10.0-windows\win-x64\publish\ExcelTools.exe

	常用变体:
	. 32 位: -r win-x86
	. 单文件 + 框架依赖(约 22 MB, 客户机需装 .NET 10 Desktop Runtime): 再加 --no-self-contained
	. 目录形式(启动最快, 40+ 个文件, 约 150 MB): 去掉 -p:PublishSingleFile=true
	. 单文件不压缩(约 150 MB, 启动略快): 再加 -p:EnableCompressionInSingleFile=false
	. 指定输出目录: 加 -o ..\dist\win-x64

	说明:
	. 体积构成: 应用与三方库约 22 MB, .NET 10 运行时(WPF)约 131 MB, 压缩后单文件约 71 MB
	. WPF 不支持裁剪, 加 -p:PublishTrimmed=true 会报 NETSDK1168
	. 单文件首次运行会把原生库解压到 %TEMP%\.net\, 启动约 1.5 秒(普通版本约 1 秒)
	. Visual Studio: 右键项目 → 发布 → 文件夹 → 目标运行时 win-x64 → 勾选"生成单个文件"

# 天启幻境II_ExcelTools

## 预览
![image](windowImport.png)

## 介绍
	可以将xls/xlsx中的数据快速导入天启幻境II的数据库


## 支持导入
	. 地图
	. 掉落
	. 副本
	. 怪物
	. 任务
	. 商店
	. 物品
	. 装备
	. 技能
MATLAB_STK_TLE_Importer

在轨航天器 TLE 查询、下载及 STK 三维可视化导入工具​

📌 简介

本软件基于 MATLAB App（使用 uifigureGUI）开发，集成了以下功能：

    1、CelesTrak SATCAT 卫星搜索（支持名称模糊搜索与 NORAD 编号精确查询，JSON API）

    2、表格多选批量下载 TLE（自动生成 .tle文件）

    3、连接 AGI STK​ 并导入卫星轨道，建立场景、添加地面站、计算访问数据

    4、中文界面优化（微软雅黑字体、颜色区分状态）

    5、实时进度显示（下载进度条 + 主进度对话框）

    6、适用于航天任务分析、轨道仿真、教学演示等场景。

✨ 主要功能

1、卫星搜索​

    按名称模糊匹配或 NORAD 编号精确查找

    支持筛选：有效载荷 / 活跃 / 在轨
 
    搜索结果以表格展示，带“选择”复选框

2、批量 TLE 下载​

    从表格选择或手动输入 NORAD 编号（逗号分隔）

    调用 CelesTrak gp.php接口下载最新 TLE

    支持部署环境（Java 网络库 / curl）与 MATLAB 环境（webread）

3、STK 场景构建与导入​

    自动连接本地 STK（支持多版本 COM ProgID）

    新建场景并设置时间范围

    解析 TLE 文件并批量创建卫星对象

    可选创建星座、地面站（预设中国主要站点）、省会城市

4、计算卫星-地面站访问链并导出报告（TXT / CSV）


🛠️ 运行环境

    操作系统：Windows（推荐 Win10/Win11）

    MATLAB：R2016b 及以上（使用 uifigure，建议 R2019b+ 以获得最佳兼容性）

    STK：已安装 AGI STK（支持 STK 11~13，自动探测 COM 接口）

    网络：可访问 https://celestrak.org（用于 SATCAT 查询与 TLE 下载）

    编译器（可选）：若需生成独立可执行文件，需 MATLAB Compiler

🚀 使用方法

启动程序​

    matlab
    复制
    Search_Download_TLE_ImportSTK

    或在 MATLAB 编辑器直接运行脚本。

搜索卫星​

    在左侧输入名称或 NORAD 编号，设置筛选条件与最大返回数

    点击【搜索卫星】，结果将显示在表格中

选择并下载 TLE​

    在表格勾选需要的卫星，或手动在右侧文本框输入 NORAD 编号（逗号分隔）

    点击【下载并保存 TLE】，进度条与状态栏会显示过程

下载的 TLE 保存在程序目录下的 downloaded_satellites.tle

导入 STK​

    设置场景时间范围（天）、选择地面站（可多选）

    勾选【中国省会城市】可自动添加全国省会设施

    点击【连接 STK 软件并导入 TLE 中所有卫星】

程序会自动创建场景、导入卫星、建立星座与地面站、计算访问数据并导出报告至 CZML_Results目录

查看结果​

STK 场景中可查看卫星轨道、访问时段

报告文件：Access_Report.txt、Access_Data.csv

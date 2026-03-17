# 出勤业绩整合系统

这是一个可直接在浏览器打开的本地小系统，用来把两类表格整合到一起：

- 出勤打卡表：按小时记录员工当天是否在线/出勤
- 业绩统计表：记录 N5 派单、自拓、今日量、日目标、月目标和完成率

系统会按 `日期 + 姓名` 自动合并数据，并生成整合看板。

## 功能

- 员工档案管理：姓名、团队、个人日目标、月目标
- 团队档案管理：团队名称、团队长、团队每日设备铺设目标、团队每月设备铺设目标
- 出勤打卡：9 点到 18 点的时段勾选，自动计算出勤次数
- 业绩录入：录入 N5 派单和自拓，自动计算今日量和完成率
- 月度汇总：自动汇总当月派单、自拓、达成、月度完成率
- 整合看板：在一个表里查看出勤、业绩、团队长和团队设备铺设目标
- Excel 导出：支持导出为 `.xlsx`，包含员工档案、出勤打卡、业绩统计、整合看板四个工作表
- 数据导入导出：支持导出为 JSON，也支持再次导入
- 本地存储：默认保存在浏览器 `localStorage`

## 使用方式

直接双击打开 [index.html](/Users/Administrator/Documents/Playground/attendance_performance_system/index.html) 即可。

如果浏览器限制本地脚本，也可以在该目录启动一个本地静态服务：

```powershell
cd C:\Users\Administrator\Documents\Playground\attendance_performance_system
python -m http.server 8080
```

然后访问：

[http://localhost:8080](http://localhost:8080)

## GitHub Pages 部署

当前这套前端是纯静态页面，可以直接部署到 GitHub Pages。

仓库里已经补好了工作流：

- [.github/workflows/deploy-pages.yml](/Users/Administrator/Documents/Playground/.github/workflows/deploy-pages.yml)

部署方式：

1. 把当前项目推到 GitHub 仓库
2. 默认分支使用 `main`
3. 在 GitHub 仓库里打开 `Settings -> Pages`
4. `Source` 选择 `GitHub Actions`
5. 推送代码后，工作流会自动把 `attendance_performance_system` 目录发布到 GitHub Pages

## 云端存储说明

如果你只是把当前版本部署到 GitHub Pages：

- 页面可以在线访问
- 但数据仍然保存在浏览器 `localStorage`
- 换电脑、换浏览器、换账号后，数据不会自动同步

如果你要真正“云端同步”，建议用：

- `GitHub Pages`：负责部署前端页面
- `Supabase / Firebase`：负责数据库和账号

这样才能做到多人共用、跨设备同步、数据不丢。

页面顶部现在提供两个导出按钮：

- `导出Excel`：导出当前日期和当前筛选条件下的多工作表 Excel 文件
- `导出JSON`：导出完整原始数据，便于备份和再次导入

## 数据说明

- 出勤次数：当天勾选的时段数
- 今日量：`N5派单 + N5自拓`
- 今日完成率：`今日量 / 个人日目标`
- 当月派单总量：当前月份内该员工每天的 `N5派单` 累加
- 当月自拓总量：当前月份内该员工每天的 `N5自拓` 累加
- 当月达成：`当月派单总量 + 当月自拓总量`
- 月度完成率：`当月达成 / 月目标`

## 说明

- 页面中已预置了一份根据截图整理的示例数据
- 为了兼容网页端显示，大表格已改为卡片纵向排版，横向滚动只保留在表格容器内部
- 如果你后面愿意，我可以继续帮你加：
  - Excel 导入导出
  - 团队汇总报表
  - 登录权限
  - 后端数据库版本

# 出勤业绩整合系统

这是一个可直接部署到静态站点的前端系统，用来把以下内容整合到一个页面里：

- 出勤打卡
- BD 个人业绩
- 团队长与团队设备铺设目标
- 团队汇总与整合看板
- Supabase 云端同步

## 当前能力

- 团队档案：团队名称、团队长、日设备铺设目标、月设备铺设目标
- 员工档案：所属团队、个人月目标
- 出勤打卡：按小时勾选并自动计算出勤次数
- 业绩录入：录入 `N5派单`、`N5自拓`，并按日期填写当日日目标
- 整合看板：按 `日期 + 姓名` 自动汇总
- 导出：支持 `Excel` 和 `JSON`
- 自动保存：本地自动保存 + Supabase 云端同步
- 多设备同步：同账号登录后自动读取云端最新数据

## 本地运行

直接打开 [index.html](/Users/Administrator/Documents/Playground/attendance_performance_system/index.html) 即可。

如果浏览器限制本地脚本，建议用静态服务运行：

```powershell
cd C:\Users\Administrator\Documents\Playground\attendance_performance_system
python -m http.server 8080
```

打开：

[http://localhost:8080](http://localhost:8080)

## Supabase 接入

### 1. 创建 Supabase 项目

在 Supabase 控制台创建一个项目。

### 2. 建表

把 [supabase.schema.sql](/Users/Administrator/Documents/Playground/attendance_performance_system/supabase.schema.sql) 里的 SQL 复制到 Supabase SQL Editor 里执行。

### 3. 配置前端

编辑 [supabase.config.js](/Users/Administrator/Documents/Playground/attendance_performance_system/supabase.config.js)：

```js
window.SUPABASE_CONFIG = {
  url: "https://你的项目id.supabase.co",
  anonKey: "你的anon key"
};
```

### 4. 登录使用

页面顶部有云端同步区域：

- `登录云端`：已有账号直接登录
- `注册账号`：新账号注册
- `退出登录`：退出当前账号

如果你在 Supabase 项目里开启了邮箱验证，注册后需要先去邮箱确认，再回来登录。

## GitHub Pages 部署

仓库已经包含 GitHub Pages 工作流：

- [.github/workflows/deploy-pages.yml](/Users/Administrator/Documents/Playground/.github/workflows/deploy-pages.yml)

你只需要：

1. 把仓库推到 GitHub
2. 仓库默认分支设为 `main`
3. 在 `Settings -> Pages` 里把 `Source` 设为 `GitHub Actions`
4. 在 `Settings -> Secrets and variables -> Actions` 新建这两个仓库密钥：
   - `SUPABASE_URL`
   - `SUPABASE_ANON_KEY`

之后每次推送，工作流会自动部署 `attendance_performance_system` 到 GitHub Pages，并在部署时生成 `supabase.config.js`。

## 数据说明

- 出勤次数：当天勾选的时段数
- 今日量：`N5派单 + N5自拓`
- 今日完成率：`今日量 / 当日日目标`
- 月度达成：当前月份内每天业绩累计
- 月度完成率：`月度达成 / 个人月目标`
- 团队日完成率：`团队今日量 / 团队日设备铺设目标`
- 团队月完成率：`团队月度达成 / 团队月设备铺设目标`

## 文件说明

- [index.html](/Users/Administrator/Documents/Playground/attendance_performance_system/index.html)：页面结构
- [styles.css](/Users/Administrator/Documents/Playground/attendance_performance_system/styles.css)：界面样式
- [app.js](/Users/Administrator/Documents/Playground/attendance_performance_system/app.js)：业务逻辑、自动保存、Supabase 同步
- [supabase.config.js](/Users/Administrator/Documents/Playground/attendance_performance_system/supabase.config.js)：前端 Supabase 配置
- [supabase.schema.sql](/Users/Administrator/Documents/Playground/attendance_performance_system/supabase.schema.sql)：建表和策略 SQL

## 下一步建议

如果你要继续往正式系统走，我建议下一步做：

1. 账号角色：管理员、团队长、普通 BD
2. 按团队长只看自己团队的数据
3. 操作日志
4. Excel 导入原始表格

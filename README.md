# 微信群上新转发自动化

这个目录里放的是一个 Windows 本地脚本骨架，用来把 `来源微信群` 的 `最新一条消息` 原样转发到一个或多个 `目标微信群`。

联调建议：

- 先保持 `dry_run: true`
- 先用你自己的测试群验证，不要直接对真实客户群试发
- 如果你要联调转发窗口本身，可以把 `preview_forward_dialog` 设为 `true`

它适合你的场景：

- 来源和目标都在微信 PC 版
- 内容可能是文字、图片、视频
- 希望每天定时执行

当前实现走的是 `微信自身的转发流程`，不是重新拼消息，所以比复制粘贴更适合图片和视频。

## 现阶段能力

- 打开微信主窗口
- 搜索并进入来源群
- 尝试定位当前聊天里的最新一条消息
- 如果拿不到消息控件，就按窗口相对坐标点击最新可见消息
- 调起右键菜单并点击 `转发`
- 在转发窗口里搜索目标群并发送

## 已知限制

- 微信 UI 会升级，控件名称可能变化，第一次通常需要调试
- “最新一条消息” 目前是按界面底部控件做启发式定位，不保证 100% 命中
- 如果回退到坐标模式，微信窗口大小变化会影响命中位置
- 如果当天来源群里发了多条上新，当前脚本只转发一条
- 如果微信弹出升级提示、登录失效、群名重名，都会影响自动化

## 安装

```powershell
python -m pip install -r requirements.txt
Copy-Item config.example.json config.json
```

然后编辑 `config.json`，至少填这几项：

- `source_chat`: 来源群名称
- `target_chats`: 目标群名称列表
- `dry_run`: 先保持 `true`，确认定位没问题后再改成 `false`
- `preview_forward_dialog`: 打开转发流程但不执行最终发送，适合联调
- `search_box_x_ratio` / `search_box_y_ratio`: 左上角搜索框的大致相对位置
- `text_message_click_x_ratio` / `text_message_click_y_ratio`: 上一条文字消息的大致位置
- `message_click_x_ratio` / `message_click_y_ratio`: 坐标回退模式下，点击聊天区消息的大致相对位置
- `grouped_video_click_x_ratio` / `grouped_video_click_y_ratio`: 多选模式下紧挨着的视频消息位置
- `multi_select_checkbox_x_ratio`: 多选模式左侧勾选圈的大致横向位置
- `multi_select_menu_y_offset`: 文字消息右键菜单中“多选”的纵向偏移
- `multi_select_menu_row_index`: 右键菜单中“多选”所在的行序号
- `multi_select_menu_click_ratio`: 右键菜单里“多选”所在的大致纵向比例
- `multi_forward_button_x_ratio` / `multi_forward_button_y_ratio`: 多选模式底部“逐条转发/转发”按钮位置
- `forward_dialog_*`: 转发弹窗中搜索框、结果项、发送按钮的大致位置
- `forward_dialog_cancel_x_ratio` / `forward_dialog_cancel_y_ratio`: 转发弹窗里取消按钮位置
- `forward_text_then_video`: 按“文字消息 + 紧跟视频消息”逐条转发
- `forward_as_grouped_pair`: 优先走“多选组合后逐条转发”的流程
- `batch_mode_enabled`: 开启从底部往上批量扫描
- `max_pairs_per_run`: 单次最多处理多少组上新
- `max_scroll_pages`: 单次最多往上翻多少次
- `stop_after_seen_pairs`: 连续遇到多少个已发送组合后停止
- `stop_after_empty_pages`: 连续几页都没识别到组合后停止
- `scroll_wheel_delta`: 每次往上翻的滚轮力度
- `state_file`: 已发送组合的本地去重状态文件

## 先做一次本地验证

1. 打开并登录微信 PC 版
2. 确保来源群和目标群都能在微信搜索框里搜到
3. 在当前目录运行：

```powershell
python .\wechat_forwarder.py --config .\config.json
```

如果输出 `Dry run 模式`，说明脚本至少已经找到微信窗口、来源群和一条候选消息。

如果定位失败，可以导出 UI 树：

```powershell
python .\wechat_forwarder.py --config .\config.json --dump-ui
```

会生成 `wechat_ui_dump.txt`，后续可以根据这个文件继续针对你的微信版本做适配。

## 接到每天定时任务

先把 `config.json` 里的 `dry_run` 改成 `false`，确认手动执行能真的转发，再配置 Windows 任务计划程序。

示例命令：

```powershell
schtasks /Create /SC DAILY /TN "WeChatForwarder" /TR "powershell -NoProfile -ExecutionPolicy Bypass -Command cd 'C:\Users\Administrator\Documents\Playground'; python .\wechat_forwarder.py --config .\config.json" /ST 09:30
```

这会每天 `09:30` 执行一次。你也可以把时间改成自己需要的时间。

## 更稳的下一步

如果你要真正长期使用，我建议下一步做这 3 个增强：

1. 增加“只转发今天还没转过的消息”，避免重复发
2. 增加“只处理包含上新关键词或图片/视频的消息”
3. 根据你电脑上的微信 UI 实际结构，定制最新消息定位逻辑

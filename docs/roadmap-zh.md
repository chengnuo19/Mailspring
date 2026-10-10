# Mailspring 二次开发改进清单（对标新版 Windows Outlook）

> 基线：Mailspring 1.26.0（`master@cfb96a4`），开发分支 `dev/outlook-ui`。
> 对照物：用户提供的新版 Outlook（深色、背景图主题）截图，以及在本机开发版中添加 Gmail 账户后的主界面截图（2026-10-10）。
> 工作量：S ≤ 1 天，M 2–4 天，L ≥ 1 周。优先级：P0 先做，P1 主体，P2 之后。

## 一、现状快照

| 区域 | 新版 Outlook | Mailspring 现状 |
|---|---|---|
| 标题栏 | 应用名 + 居中搜索框 + 通知铃铛 / 「…」 | 标题显示「Mailspring · 当前邮件主题」，搜索框在邮件列表上方 |
| 命令区 | 「文件 / 主页 / 查看 / 帮助」标签 + 带文字的简化命令栏，「新邮件」为蓝色主按钮 | 一排纯图标按钮，写信按钮是左上角小图标 |
| 左侧应用栏 | 邮件、日历、联系人、待办等彩色图标 | 已有导航栏，只有邮件、日历两项 |
| 文件夹面板 | 「收藏夹」+ 按账户折叠；中文文件夹名；未读数为蓝色数字 | 无收藏夹；**Inbox / Sent Mail / All Mail / Spam / Trash / Important 显示英文**；未读数为色块徽标 |
| 邮件列表 | 「重点 / 其他」标签；彩色首字母头像；按日期分组（更早…）；未读左侧蓝条 + 主题蓝色；日期「周五 9/25」 | 无重点收件箱、无头像、无分组；选中项整行青色实底；日期「周五下午10:44」 |
| 阅读窗格 | 发件人头像 + 操作按钮组；未选中时显示背景图 | 有「Inbox」米色标签块，回复入口为「写回复......」 |
| 右侧栏 | 默认无 | 默认显示联系人栏，含 Mailspring Pro 推广 |
| 主题 | 背景图 + 浮在上面的圆角面板 | 纯色深色主题 |
| 发信 | 走 HTTPS，不受 SMTP 端口限制 | 走 SMTP；本机出口节点（DigitalOcean）封了 465/587，**目前无法发信** |

## 二、改进项

### P0：见效快、影响面大

| # | 改进项 | 说明 | 涉及文件 | 工作量 |
|---|---|---|---|---|
| 1 | 标准文件夹显示中文 | 有 `role` 的文件夹改用已有的 `LocalizedStringForRole`，目前只有设置页在用 | `app/src/flux/models/category.ts`，`app/internal_packages/account-sidebar/` | S |
| 2 | 中文用语按 Outlook 统一 | Archive→存档、Snooze→推迟（现为「已设置为Snoozed」）、Trash→已删除邮件、Thread→对话、Reply All→全部回复、Compose New Message→新邮件、「写回复......」→「回复」；修正「Send and Archive→发送到收件箱并存档」错译 | `app/lang/zh-CN.json` | S |
| 3 | 补全缺失翻译与占位符 | 缺 15 条（含导航栏的 Mail / Calendar / Sections）；2 条译文丢了 `%@`，导致 Yahoo 应用密码说明里的链接不显示 | `app/lang/zh-CN.json` | S |
| 4 | 修复拼接造成的多余空格 | 「选择所有 已加星标」 | 生成该菜单项的代码（待定位） | S |
| 5 | 日期格式对齐 Outlook | 今天显示时间；一周内显示「周五 9/25」；更早显示「2026/7/26」；英文格式 `MMMM Do, h:mmA` 本地化 | `thread-list/lib/thread-list-columns.tsx`，`DateUtils` | S |
| 6 | 默认值对齐 Outlook | 默认三栏布局（`core.workspace.mode` 改为 `split`）；默认快捷键方案改为已内置的 `Outlook.json` | `app/src/config-schema.ts` | S |
| 7 | 右侧联系人栏默认关闭、去掉 Pro 推广 | Outlook 没有这一栏 | `participant-profile` 插件 | S |
| 8 | HTTPS 发信（Graph / Gmail API） | **已技术验证**：Mailspring 自带的微软、Google 应用都能用；Graph `sendMail` 返回 202，Gmail API 发送成功。微软约 5–7 天，Gmail 1–2 天，大附件、追踪、定时发送 2–3 天 | `app/src/flux/tasks/send-draft-task.ts`、`mailsync-bridge`、`onboarding-constants.ts`（加 `Mail.Send`）、新建 API 发信模块 | L（约 2 周） |

### P1：Outlook 外观与结构（主体工作）

| # | 改进项 | 说明 | 涉及文件 | 工作量 |
|---|---|---|---|---|
| 9 | `ui-outlook` / `ui-outlook-dark` 主题 | Fluent 配色（主色 `#0F6CBD`）、Segoe UI + 微软雅黑 UI、圆角与阴影、选中态改为浅色高亮 + 左侧蓝条、未读主题蓝色 | 新建 `app/internal_packages/ui-outlook*/`，参考 `ui-dark/styles/ui-variables.less`、`app/static/style/base/ui-variables.less` | M |
| 10 | 邮件列表改造 | 彩色首字母头像；按日期分组标题（今天 / 昨天 / 本周 / 上周 / 更早）；未读左侧蓝条；悬停快捷按钮（删除 / 标记 / 已读）换成 Outlook 样式和顺序 | `thread-list/lib/thread-list.tsx`、`thread-list-columns.tsx`、`thread-list-quick-actions.tsx` | M–L |
| 11 | 文件夹面板改造 | 顶部「收藏夹」（跨账户固定常用文件夹）；账户可折叠；文件夹图标统一；未读数改为蓝色数字，垃圾邮件显示 [总数] | `account-sidebar/` | M |
| 12 | 阅读窗格改造 | 发件人头像 + 回复 / 全部回复 / 转发 / 「…」按钮组；去掉或重做「Inbox」标签块；空状态显示主题背景图 | `message-list/` | M |
| 13 | 标题栏 | 应用名 + 居中搜索框 + 通知铃铛 + 「…」，搜索框从列表上方移到标题栏 | `app/src/sheet-toolbar.tsx`、`thread-search` | L |
| 14 | 功能区（Ribbon） | 「文件 / 主页 / 查看 / 帮助」标签；「主页」为简化命令栏：新邮件（主按钮 + 下拉）、删除▾、存档、标记/取消标记▾、推迟▾、「…」，可折叠；无选中时按钮置灰 | 新插件 + `thread-list/lib/thread-toolbar-buttons.tsx` | L |
| 15 | 背景图主题 | 窗口背景图 + 半透明圆角面板（类似 Outlook 的「山湖」主题） | 主题插件 + `workspace.less` | M |
| 16 | 左侧应用栏补全 | 增加联系人、待办等入口，图标改为彩色 | `app/src/components/nav-rail.tsx` 及各插件的 `NavRail:Item` | S–M |
| 17 | 「重点 / 其他」收件箱 | Gmail 可用其分类标签，微软可用 Graph 的 `inferenceClassification`，其他 IMAP 账户需要启发式规则；先调研 | `thread-list`、新插件 | L（含调研） |

### P2：性能与稳定性

| # | 改进项 | 说明 | 涉及文件 | 工作量 |
|---|---|---|---|---|
| 18 | 启动时文件夹 / 标签查询慢 | 首次启动 `SELECT Folder` / `SELECT Label` 耗时 2–5 秒，之后约 0.5 秒（日志实测） | `app/src/flux/stores/database-store.ts` 及调用方 | M |
| 19 | 合并上游有用的 PR | 例如：Windows 托盘图标更新后保持固定、通知音量与自定义提示音、账户分组（带颜色）、自定义排序 | 视 PR 而定 | 每个 S–M |
| 20 | 已完成：Windows 开发环境修复 | ① `compile-cache-ts.js`：多窗口同时写缓存时 `EPERM` 导致模块加载失败，已改为忽略缓存写入失败；② `mailsync-process.ts`：新增 `MAILSYNC_BINARY_PATH`，解决 mailsync 要求路径含「mailspring」；③ 仅开发模式的 `MAILSPRING_DEV_SKIP_SMTP_CHECK`，允许只收信添加账户 | 已改，未提交 | — |

### 待补充核对

- 写信窗口、打开单封邮件的细节还没截图，核对后补进第 12 项和写信相关条目。
- 浅色主题下的效果，在第 9 项完成后再对照。

## 三、建议的第二阶段顺序

1. **第 1–7 项（约 2–3 天）**：都是小改动，做完后中文界面、日期和默认布局会立刻明显更像 Outlook。
2. **第 9、10、12 项（约 1–1.5 周）**：主题 + 邮件列表 + 阅读窗格，完成后「看起来像 Outlook」。
3. **第 8 项 HTTPS 发信（约 2 周）**：可以和第 2 步并行，完成后本机能正常收发。
4. **第 11、13、14、15 项**：文件夹面板、标题栏、功能区、背景图，完成后结构上接近 Outlook。
5. 第 16–19 项按需穿插。

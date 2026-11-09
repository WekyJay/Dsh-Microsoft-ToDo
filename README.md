# Dsh-Microsoft-ToDo

Microsoft To Do 面板插件，为 DeepSeek Harness 的 Web 界面加一个「微软待办」面板：左侧清单（支持本地分组），右侧任务（新增 / 完成 / 重开 / 标星 / 删除），顶部今日统计，外加「我的一天」速选清单。

A Microsoft To Do panel for DeepSeek Harness: lists on the left (with local groups), tasks on the right (add / complete / reopen / star / delete), today's counters on top, and a My Day short list.

---

## 功能 / Features

| 中文 | English |
| --- | --- |
| 清单浏览、任务增删改、完成与重开 | Browse lists; create, update, complete, reopen, delete tasks |
| 今日 / 逾期 / 重要统计 | Today / overdue / starred counters |
| 分组：侧栏按组收纳清单，本机保存 | Groups: collect lists in the sidebar, stored locally |
| 我的一天：每天挑一批任务先做，次日自动清空 | My Day: pick a few tasks for today, resets each day |
| 设备码登录，令牌由上游 MCP 服务缓存 | Device-code sign-in; the token cache lives in the upstream MCP server |
| 中英双语界面 | Chinese and English UI |

## 安装 / Install

插件是一个标准的 DSH bundle（`package.json` 里的 `dsh.bundle` + `dsh.client`）。两种装法：

1. **从本地目录安装**（开发用）：

   ```text
   plugin_manager: action=install_bundle, target=<本仓库绝对路径>
   ```

   已装过时先 `remove_bundle target=dsh-microsoft-todo`，DSH 不会热重载宿主模块。

2. **从插件市场安装**：在 DSH 的设置 → 插件市场里添加本仓库的索引地址 `https://<owner>.github.io/Dsh-Microsoft-ToDo/market/index.json`，然后在 Discover 里搜索「微软待办」安装。

## 首次使用 / First run

面板第一次打开会显示登录页：点「开始登录」，复制设备代码，在浏览器打开 `https://microsoft.com/devicelogin` 输入并同意授权。之后令牌由 `microsoft-todo-mcp` 子进程缓存在 `~/.config/microsoft-todo-mcp/token-cache.json`（权限 0600），访问令牌过期时自动用刷新令牌续期，无需重复登录。

## 数据与本机文件 / Local state

| 文件 | 内容 | 是否入库提交 |
| --- | --- | --- |
| `groups.json` | 清单分组（微软 Graph 没有分组资源，只能本机保存） | 否 |
| `myday.json` | 「我的一天」当天选中的任务 id | 否 |
| `~/.config/microsoft-todo-mcp/token-cache.json` | MSAL 令牌缓存（上游服务写入） | 否，且不要提交 |

`.gitignore` 已排除以上个人数据；分组和「我的一天」都不会同步到其他设备或微软 To Do 客户端。

## 已知限制 / Known limitations

- 微软 Graph 的 To Do API **没有分组**，也没有 **My Day** 资源（Planner 有 `myDayTasks`，To Do 没有）。因此本插件的分组与「我的一天」是本机实现。
- 每个清单最多读取 200 条任务（`list_tasks` 的 `maxResults` 上限内），超出的任务不会出现在面板里。
- 任务排序沿用 Graph 返回的顺序（To Do 客户端的自定义排序不在 API 里）。

## 依赖 / Dependencies

- 上游 MCP 服务 [`microsoft-todo-mcp`](https://www.npmjs.com/package/microsoft-todo-mcp)（`npx -y microsoft-todo-mcp`），由本插件的宿主半边常驻拉起，走 MCP stdio。
- 该服务默认使用微软公共客户端 id（Microsoft Graph Command Line Tools），只需 `Tasks.ReadWrite` 委托权限，无需自建 Entra 应用。

## 卸载 / Uninstall

```text
plugin_manager: action=remove_bundle, target=dsh-microsoft-todo
```

删除后 `groups.json` / `myday.json` 仍留在插件目录；令牌缓存需手动删除 `~/.config/microsoft-todo-mcp/`。

## License

MIT

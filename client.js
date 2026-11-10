/**
 * Microsoft To Do bridge, browser half.
 *
 * Registers one global panel: a sidebar entry (`sidebar.panellist`) and the
 * page it opens (`main`), both under the same panel id. The page reads the
 * {@link ROUTE_PREFIX} JSON API served by the host half.
 */
window.__ModuleLoader__.load({
  id: 'dsh-microsoft-todo',
  factory(require) {
    const React = require('react')
    const h = React.createElement
    const { useState, useEffect, useCallback, useRef } = React

    const NS = 'microsoftTodo'
    const PANEL_ID = 'microsoft-todo'
    const API = 'microsoft-todo/api/'
    const LS_LIST = 'microsoft-todo:list'

    const DICT = {
      zh: { panel: '微软待办', title: '微软待办', addPlaceholder: '添加任务，回车保存…', add: '添加', refresh: '刷新', today: '今日', overdue: '逾期', starred: '重要', lists: '清单', open: '未完成', done: '已完成', all: '全部', empty: '这个清单还没有任务', emptyOpen: '没有未完成的任务，清爽！', loading: '正在连接 Microsoft To Do…', retry: '重试', signInTitle: '连接你的 Microsoft To Do', signInBody: '需要在微软登录一次，之后由本机保存的令牌自动续期。', signInButton: '开始登录', signInHint: '打开下面的链接，输入代码完成授权：', codeCopied: '代码已复制', waiting: '等待你在浏览器里完成授权…', deleteConfirm: '删除这条任务？', failed: '操作失败', account: '已连接', star: '标为重要', unstar: '取消重要', markDone: '标记完成', reopen: '恢复未完成', recommended: '今日优先', groups: '分组', ungrouped: '未分组', groupSettings: '分组设置', newGroup: '新建分组', groupPlaceholder: '分组名称', save: '保存', cancel: '取消', deleteGroup: '删除分组', noGroup: '不分组', groupHint: '分组只保存在本机：微软 Graph 不提供 To Do 清单的分组信息，因此这里的分组不会同步到其他设备。', myday: '我的一天', mydayEmpty: '今天还没挑任务', mydayHint: '点任务右侧的太阳图标，把任务放进「我的一天」；每天自动清空。', addMyDay: '加入我的一天', removeMyDay: '移出我的一天', flaggedEmails: '标记的电子邮件', detail: '详情', notes: '备注', notesPlaceholder: '写点备注…', dueDate: '截止日期', reminder: '提醒我', noReminder: '无提醒', recurrence: '重复', repeatNone: '不重复', repeatDaily: '每天', repeatWeekly: '每周', repeatMonthly: '每月', repeatInterval: '间隔', repeatDays: '重复于', repeatEnd: '结束', repeatNoEnd: '永不结束', repeatEndDate: '在日期结束', steps: '下一步', stepPlaceholder: '添加一个步骤…', promote: '提升为任务', attachments: '添加文件', attach: '上传文件', noAttachments: '还没有文件', remove: '移除', saving: '保存中…', saved: '已保存', uploading: '上传中…', fileTooLarge: '文件超过 25 MB', repeatHint: '服务端可能把结束方式保存为「永不结束」，这是 Graph 的已知问题。', doneLabel: '完成', version: '版本', compatOk: 'DSH 兼容', compatBad: 'DSH 版本过低', compatUnknown: 'DSH 版本未知', compatWarn: '当前 DSH {dsh} 低于本插件要求的 {required}，部分功能可能异常。请升级 DSH（dsh update）后刷新页面。', compatUnknownHint: '未能识别当前 DSH 版本，无法确认兼容性。', deleteTask: '删除任务' },
      en: { panel: 'To Do', title: 'Microsoft To Do', addPlaceholder: 'Add a task, press Enter…', add: 'Add', refresh: 'Refresh', today: 'Today', overdue: 'Overdue', starred: 'Starred', lists: 'Lists', open: 'Open', done: 'Completed', all: 'All', empty: 'No tasks in this list yet', emptyOpen: 'Nothing open here — all clear.', loading: 'Connecting to Microsoft To Do…', retry: 'Retry', signInTitle: 'Connect your Microsoft To Do', signInBody: 'Sign in once; the token cached on this machine refreshes itself afterwards.', signInButton: 'Sign in', signInHint: 'Open the link below and enter the code:', codeCopied: 'Code copied', waiting: 'Waiting for you to finish signing in…', deleteConfirm: 'Delete this task?', failed: 'Action failed', account: 'Connected', star: 'Mark as important', unstar: 'Remove importance', markDone: 'Mark complete', reopen: 'Reopen', recommended: 'Today first', groups: 'Groups', ungrouped: 'Ungrouped', groupSettings: 'Manage groups', newGroup: 'New group', groupPlaceholder: 'Group name', save: 'Save', cancel: 'Cancel', deleteGroup: 'Delete group', noGroup: 'No group', groupHint: 'Groups are stored on this machine only: the Microsoft Graph To Do API does not expose list groups, so they do not sync to other devices.', myday: 'My Day', mydayEmpty: 'Nothing picked for today yet', mydayHint: 'Use the sun button on a task to add it to My Day; the list resets every day.', addMyDay: 'Add to My Day', removeMyDay: 'Remove from My Day', flaggedEmails: 'Flagged emails', detail: 'Details', notes: 'Notes', notesPlaceholder: 'Write a note…', dueDate: 'Due date', reminder: 'Remind me', noReminder: 'No reminder', recurrence: 'Repeat', repeatNone: 'Never', repeatDaily: 'Daily', repeatWeekly: 'Weekly', repeatMonthly: 'Monthly', repeatInterval: 'Every', repeatDays: 'Repeat on', repeatEnd: 'Ends', repeatNoEnd: 'Never', repeatEndDate: 'On date', steps: 'Steps', stepPlaceholder: 'Add a step…', promote: 'Promote to task', attachments: 'Files', attach: 'Upload file', noAttachments: 'No files yet', remove: 'Remove', saving: 'Saving…', saved: 'Saved', uploading: 'Uploading…', fileTooLarge: 'File is larger than 25 MB', repeatHint: 'Graph may persist the range as never-ending; that is a known service issue.', doneLabel: 'Complete', version: 'Version', compatOk: 'DSH compatible', compatBad: 'DSH too old', compatUnknown: 'DSH version unknown', compatWarn: 'This DSH is {dsh}, below the {required} this plugin needs. Some features may misbehave. Upgrade DSH (dsh update), then reload this page.', compatUnknownHint: 'The running DSH version could not be detected, so compatibility is unverified.', deleteTask: 'Delete task' },
    }

    const CSS = `
/* shadcn/ui design tokens, bridged onto the harness theme variables. */
.mtx-page{
  --mtx-bg:var(--dsw-alias-bg-layer-1);
  --mtx-surface:var(--dsw-alias-bg-layer-2);
  --mtx-fg:var(--dsw-alias-label-primary);
  --mtx-muted:var(--dsw-alias-label-secondary);
  --mtx-border:var(--dsw-alias-border-l2);
  --mtx-border-strong:var(--dsw-alias-border-l4);
  --mtx-accent:var(--dsw-alias-interactive-bg-hover);
  --mtx-primary:#2564CF;
  --mtx-primary-hover:#1B4EB0;
  --mtx-primary-fg:#FFFFFF;
  --mtx-destructive:#C4314B;
  --mtx-ring:color-mix(in srgb,var(--mtx-primary) 42%,transparent);
  --mtx-r:10px;
  --mtx-r-sm:8px;
  display:flex;flex-direction:column;height:100%;min-height:0;
  background:var(--mtx-bg);color:var(--mtx-fg);
  font-size:13px;line-height:1.5;letter-spacing:-.006em;
  -webkit-font-smoothing:antialiased;
}
.mtx-page *{box-sizing:border-box}
.mtx-page ::-webkit-scrollbar{width:10px;height:10px}
.mtx-page ::-webkit-scrollbar-thumb{background:color-mix(in srgb,var(--mtx-muted) 26%,transparent);border-radius:999px;border:3px solid transparent;background-clip:content-box}
.mtx-page ::-webkit-scrollbar-track{background:transparent}
.mtx-page :focus-visible{outline:none;box-shadow:0 0 0 2px var(--mtx-bg),0 0 0 4px var(--mtx-ring)}

/* ---------- header ---------- */
.mtx-railhead{display:flex;align-items:center;gap:8px;padding:2px 4px 12px}
.mtx-railhead .mtx-iconbtn{width:30px;height:30px;border:1px solid var(--mtx-border);border-radius:var(--mtx-r-sm);background:var(--mtx-bg)}
.mtx-railhead .mtx-iconbtn:hover{background:color-mix(in srgb,var(--mtx-accent) 72%,var(--mtx-bg))}
.mtx-brand{display:flex;align-items:center;gap:10px;min-width:0;flex:1}
.mtx-logo{
  width:34px;height:34px;border-radius:var(--mtx-r);flex:none;
  background:linear-gradient(180deg,#2E6FE0,#1D4FA6);color:#fff;
  display:flex;align-items:center;justify-content:center;
  font-size:16px;font-weight:700;
  box-shadow:0 1px 2px rgba(15,23,42,.18),inset 0 1px 0 rgba(255,255,255,.18);
}
.mtx-title{margin:0;font-size:15.5px;font-weight:600;letter-spacing:-.02em;line-height:1.3}
.mtx-sub{margin:2px 0 0;font-size:11.5px;color:var(--mtx-muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mtx-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;padding:0 2px 12px;margin-bottom:8px;border-bottom:1px solid var(--mtx-border)}
.mtx-stat{
  min-width:0;padding:6px 9px;border-radius:var(--mtx-r-sm);
  border:1px solid var(--mtx-border);background:var(--mtx-surface);
  box-shadow:0 1px 1px rgba(15,23,42,.04);
}
.mtx-stat-label{font-size:10px;font-weight:500;color:var(--mtx-muted);letter-spacing:.02em}
.mtx-stat-value{margin-top:1px;font-size:17px;font-weight:600;line-height:1.25;font-variant-numeric:tabular-nums;letter-spacing:-.02em}
.mtx-stat.today .mtx-stat-value{color:var(--mtx-primary)}
.mtx-stat.overdue .mtx-stat-value{color:var(--mtx-destructive)}
.mtx-stat.starred .mtx-stat-value{color:#B0790F}

/* ---------- buttons (shadcn variants) ---------- */
.mtx-btn{
  display:inline-flex;align-items:center;justify-content:center;gap:6px;
  height:32px;padding:0 12px;border-radius:var(--mtx-r-sm);
  border:1px solid var(--mtx-border);background:var(--mtx-bg);color:var(--mtx-fg);
  font:inherit;font-size:12.5px;font-weight:500;white-space:nowrap;cursor:pointer;
  transition:background-color .15s ease,border-color .15s ease,color .15s ease,box-shadow .15s ease;
}
.mtx-btn:hover{background:color-mix(in srgb,var(--mtx-accent) 72%,var(--mtx-bg));border-color:var(--mtx-border-strong)}
.mtx-btn:active{transform:translateY(.5px)}
.mtx-btn:disabled{opacity:.5;pointer-events:none}
.mtx-btn.primary{background:var(--mtx-primary);border-color:var(--mtx-primary);color:var(--mtx-primary-fg);box-shadow:0 1px 2px rgba(37,100,207,.28)}
.mtx-btn.primary:hover{background:var(--mtx-primary-hover);border-color:var(--mtx-primary-hover)}

/* ---------- shell ---------- */
.mtx-body{flex:1;display:flex;min-height:0}
.mtx-aside{
  width:276px;flex:none;display:flex;flex-direction:column;min-height:0;
  overflow-y:auto;padding:14px 10px 12px;border-right:1px solid var(--mtx-border);
  background:color-mix(in srgb,var(--mtx-surface) 42%,var(--mtx-bg));
}
.mtx-aside-head{
  display:flex;justify-content:space-between;align-items:center;
  padding:2px 8px 8px;font-size:10.5px;font-weight:600;letter-spacing:.06em;
  text-transform:uppercase;color:var(--mtx-muted);
}
.mtx-listrow{
  display:flex;align-items:center;gap:8px;width:100%;height:32px;padding:0 8px;
  border:0;border-radius:var(--mtx-r-sm);background:transparent;color:inherit;
  font:inherit;font-size:12.5px;font-weight:450;text-align:left;cursor:pointer;
  transition:background-color .14s ease,color .14s ease;
}
.mtx-listrow:hover{background:color-mix(in srgb,var(--mtx-accent) 66%,transparent)}
.mtx-listrow.on{background:color-mix(in srgb,var(--mtx-primary) 12%,transparent);color:var(--mtx-fg);font-weight:600;box-shadow:inset 2px 0 0 var(--mtx-primary)}
.mtx-listrow.on:hover{background:color-mix(in srgb,var(--mtx-primary) 16%,transparent)}
.mtx-listname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-listcount{flex:none;font-size:11px;color:var(--mtx-muted);font-variant-numeric:tabular-nums}
.mtx-listdot{width:6px;height:6px;border-radius:999px;flex:none;background:color-mix(in srgb,var(--mtx-primary) 55%,transparent)}
.mtx-listdot.late{background:var(--mtx-destructive)}
.mtx-listrow.sub{padding-left:22px}
.mtx-myday{height:34px;font-weight:600}
.mtx-aside-sub{padding:10px 8px 4px;font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--mtx-muted)}
.mtx-group{
  display:flex;align-items:center;gap:6px;width:100%;margin-top:6px;padding:6px 8px;
  border:0;border-radius:var(--mtx-r-sm);background:transparent;color:inherit;
  font:inherit;cursor:pointer;
}
.mtx-group:hover{background:color-mix(in srgb,var(--mtx-accent) 66%,transparent)}
.mtx-gchev{width:11px;height:11px;flex:none;opacity:.6;transition:transform .16s ease}
.mtx-gchev.open{transform:rotate(90deg)}
.mtx-gname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11.5px;font-weight:600;color:var(--mtx-muted);text-align:left}
.mtx-gcount{font-size:10.5px;color:var(--mtx-muted);font-variant-numeric:tabular-nums}

/* ---------- sidebar extras ---------- */
.mtx-todaycard{margin-top:14px;padding-top:10px;border-top:1px solid var(--mtx-border)}
.mtx-todaycard h4{margin:0 0 4px;padding:0 8px;font-size:10.5px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--mtx-muted)}
.mtx-mini{
  display:flex;align-items:flex-start;gap:7px;width:100%;padding:5px 8px;
  border:0;border-radius:var(--mtx-r-sm);background:transparent;color:inherit;
  font:inherit;font-size:12px;line-height:1.4;text-align:left;cursor:pointer;
}
.mtx-mini:hover{background:color-mix(in srgb,var(--mtx-accent) 66%,transparent)}
.mtx-mini-mark{width:5px;height:5px;margin-top:6px;flex:none;border-radius:999px;background:var(--mtx-primary)}
.mtx-mini-mark.late{background:var(--mtx-destructive)}
.mtx-mini-text{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-mini-list{flex:none;max-width:84px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:10.5px;color:var(--mtx-muted)}
.mtx-aside-foot{margin-top:auto;padding-top:8px;display:flex;flex-direction:column;gap:2px}
.mtx-footbtn{
  display:flex;align-items:center;gap:7px;width:100%;height:30px;padding:0 8px;
  border:0;border-radius:var(--mtx-r-sm);background:transparent;color:var(--mtx-muted);
  font:inherit;font-size:12px;cursor:pointer;transition:background-color .14s ease,color .14s ease;
}
.mtx-footbtn:hover{background:color-mix(in srgb,var(--mtx-accent) 66%,transparent);color:var(--mtx-fg)}
.mtx-verline{display:flex;align-items:center;justify-content:space-between;gap:6px;padding:4px 8px 0;font-size:10.5px;color:var(--mtx-muted)}
.mtx-verchip{border:1px solid transparent;border-radius:999px;padding:1px 7px;font-size:10px;font-weight:500;line-height:1.65;white-space:nowrap}
.mtx-verchip.ok{color:#1B7A45;background:rgba(27,122,69,.1);border-color:rgba(27,122,69,.24)}
.mtx-verchip.bad{color:#B42318;background:rgba(180,35,24,.1);border-color:rgba(180,35,24,.28)}
.mtx-verchip.unknown{color:var(--mtx-muted);background:var(--mtx-surface);border-color:var(--mtx-border)}

/* ---------- main ---------- */
.mtx-main{flex:1;min-width:0;display:flex;flex-direction:column;min-height:0;padding-top:14px}
.mtx-warn{
  margin:0 20px 12px;padding:9px 12px;border-radius:var(--mtx-r-sm);
  font-size:12px;line-height:1.55;color:#8A1C13;
  background:rgba(180,35,24,.07);border:1px solid rgba(180,35,24,.22);
}
.mtx-toolbar{display:flex;align-items:center;gap:12px;padding:14px 20px 8px;flex:none}
.mtx-tabs{
  display:inline-flex;gap:2px;padding:3px;border-radius:10px;
  background:color-mix(in srgb,var(--mtx-surface) 86%,var(--mtx-bg));
  border:1px solid var(--mtx-border);
}
.mtx-tab{
  height:26px;padding:0 11px;border:0;border-radius:7px;background:transparent;
  color:var(--mtx-muted);font:inherit;font-size:12px;font-weight:500;cursor:pointer;
  transition:background-color .14s ease,color .14s ease,box-shadow .14s ease;
}
.mtx-tab:hover{color:var(--mtx-fg)}
.mtx-tab.on{background:var(--mtx-bg);color:var(--mtx-fg);box-shadow:0 1px 2px rgba(15,23,42,.14),0 0 0 1px var(--mtx-border)}
.mtx-hint{margin-left:auto;font-size:11.5px;color:var(--mtx-muted);font-variant-numeric:tabular-nums}
.mtx-add{display:flex;gap:8px;align-items:center;padding:0 20px 10px;flex:none}
.mtx-input{
  flex:1;min-width:0;height:34px;padding:0 11px;border-radius:var(--mtx-r-sm);
  border:1px solid var(--mtx-border);background:var(--mtx-bg);color:var(--mtx-fg);
  font:inherit;font-size:13px;outline:none;
  transition:border-color .15s ease,box-shadow .15s ease;
}
.mtx-input::placeholder{color:var(--mtx-muted)}
.mtx-input:focus{border-color:var(--mtx-primary);box-shadow:0 0 0 3px var(--mtx-ring)}
.mtx-date{
  height:34px;padding:0 8px;border-radius:var(--mtx-r-sm);
  border:1px solid var(--mtx-border);background:var(--mtx-bg);color:var(--mtx-muted);
  font:inherit;font-size:12px;outline:none;
}
.mtx-date:focus{border-color:var(--mtx-primary);box-shadow:0 0 0 3px var(--mtx-ring)}
.mtx-listhead{display:flex;align-items:center;gap:8px;min-height:28px;padding:0 20px 8px;flex:none}
.mtx-listhead-title{font-size:13.5px;font-weight:600;letter-spacing:-.015em;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-listhead-count{font-size:11.5px;color:var(--mtx-muted);font-variant-numeric:tabular-nums}
.mtx-list{flex:1;min-height:0;overflow-y:auto;padding:0 12px 24px;display:flex;flex-direction:column;gap:1px}

/* ---------- task rows ---------- */
.mtx-row{
  display:flex;align-items:flex-start;gap:10px;padding:8px;margin:0 8px;
  border-radius:var(--mtx-r-sm);border:1px solid transparent;
  transition:background-color .14s ease,border-color .14s ease;
}
.mtx-row:hover{background:color-mix(in srgb,var(--mtx-accent) 55%,transparent)}
.mtx-row.done{opacity:.6}
.mtx-check{
  width:17px;height:17px;margin-top:2px;flex:none;padding:0;
  display:flex;align-items:center;justify-content:center;
  border:1.5px solid var(--mtx-border-strong);border-radius:5px;
  background:transparent;color:transparent;cursor:pointer;
  transition:background-color .14s ease,border-color .14s ease,color .14s ease;
}
.mtx-check:hover{border-color:var(--mtx-primary);color:color-mix(in srgb,var(--mtx-primary) 70%,transparent)}
.mtx-check.on{background:var(--mtx-primary);border-color:var(--mtx-primary);color:#fff}
.mtx-check.on:hover{color:#fff}
.mtx-row-main{flex:1;min-width:0;padding:1px 2px;border-radius:6px;cursor:pointer}
.mtx-row-title{font-size:13px;line-height:1.45;overflow-wrap:anywhere}
.mtx-row.done .mtx-row-title{text-decoration:line-through;color:var(--mtx-muted)}
.mtx-row-meta{display:flex;flex-wrap:wrap;gap:5px;margin-top:5px}
.mtx-chip{
  display:inline-flex;align-items:center;gap:3px;padding:1px 7px;border-radius:999px;
  border:1px solid var(--mtx-border);background:var(--mtx-surface);
  font-size:10.5px;font-weight:500;color:var(--mtx-muted);line-height:1.6;
}
.mtx-chip.late{color:var(--mtx-destructive);border-color:color-mix(in srgb,var(--mtx-destructive) 34%,transparent);background:color-mix(in srgb,var(--mtx-destructive) 8%,transparent)}
.mtx-chip.today{color:var(--mtx-primary);border-color:color-mix(in srgb,var(--mtx-primary) 32%,transparent);background:color-mix(in srgb,var(--mtx-primary) 8%,transparent)}
.mtx-row-actions{display:flex;gap:1px;flex:none;opacity:0;transition:opacity .14s ease}
.mtx-row:hover .mtx-row-actions,.mtx-row:focus-within .mtx-row-actions{opacity:1}
.mtx-iconbtn{
  display:flex;align-items:center;justify-content:center;
  width:26px;height:26px;padding:0;border:0;border-radius:6px;
  background:transparent;color:var(--mtx-muted);cursor:pointer;
  transition:background-color .14s ease,color .14s ease;
}
.mtx-iconbtn:hover{background:color-mix(in srgb,var(--mtx-accent) 85%,transparent);color:var(--mtx-fg)}
.mtx-iconbtn.star.on{color:#B0790F}
.mtx-iconbtn.sun.on{color:#C2700A;background:rgba(194,112,10,.12)}

/* ---------- states ---------- */
.mtx-mydayhead{display:flex;justify-content:space-between;align-items:baseline;padding:2px 20px 10px;font-size:11.5px;color:var(--mtx-muted)}
.mtx-mydayhead b{margin-left:6px;font-size:15px;font-weight:600;color:var(--mtx-fg)}
.mtx-mydayempty{margin:auto;max-width:360px;padding:56px 0;text-align:center;font-size:12.5px;line-height:1.7;color:var(--mtx-muted)}
.mtx-empty{margin:auto;padding:56px 0;text-align:center;font-size:12.5px;color:var(--mtx-muted)}
.mtx-state{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;height:100%;padding:32px;text-align:center}
.mtx-state h2{margin:0;font-size:16px;font-weight:600;letter-spacing:-.02em}
.mtx-state p{margin:0;max-width:420px;font-size:12.5px;line-height:1.65;color:var(--mtx-muted)}
.mtx-code{
  padding:10px 16px;border-radius:var(--mtx-r);cursor:pointer;
  border:1px dashed color-mix(in srgb,var(--mtx-primary) 45%,transparent);
  background:color-mix(in srgb,var(--mtx-primary) 6%,transparent);
  color:var(--mtx-primary);font-family:var(--ds-font-family-code,ui-monospace,SFMono-Regular,Menlo,monospace);
  font-size:22px;letter-spacing:.2em;
}
.mtx-link{color:var(--mtx-primary);font-size:12.5px;text-decoration:none;border-bottom:1px solid color-mix(in srgb,var(--mtx-primary) 40%,transparent)}
.mtx-toast{
  position:absolute;left:50%;bottom:20px;transform:translateX(-50%);
  max-width:80%;padding:8px 14px;border-radius:var(--mtx-r-sm);
  background:#18181B;color:#FAFAFA;font-size:12px;line-height:1.5;
  box-shadow:0 10px 30px rgba(0,0,0,.28);overflow-wrap:anywhere;
  animation:mtx-rise .18s ease-out;
}
.mtx-skel{height:40px;margin:0 8px;border-radius:var(--mtx-r-sm);background:var(--mtx-surface);animation:mtx-pulse 1.5s ease-in-out infinite}
@keyframes mtx-pulse{0%,100%{opacity:.55}50%{opacity:.9}}
@keyframes mtx-rise{from{opacity:0;transform:translate(-50%,6px)}to{opacity:1;transform:translate(-50%,0)}}

/* ---------- task detail sheet ---------- */
.mtx-mask{position:absolute;inset:0;z-index:90;background:rgba(9,9,11,.42);backdrop-filter:blur(2px);animation:mtx-fade .16s ease-out}
.mtx-drawer{
  position:absolute;top:0;right:0;bottom:0;z-index:95;display:flex;flex-direction:column;
  width:min(440px,94%);background:var(--mtx-bg);
  border-left:1px solid var(--mtx-border);
  box-shadow:-24px 0 48px rgba(15,23,42,.22);
  animation:mtx-slide .2s cubic-bezier(.22,.61,.36,1);
}
@keyframes mtx-fade{from{opacity:0}to{opacity:1}}
@keyframes mtx-slide{from{transform:translateX(16px);opacity:.6}to{transform:translateX(0);opacity:1}}
.mtx-drawer-head{display:flex;align-items:center;gap:10px;padding:14px 16px;border-bottom:1px solid var(--mtx-border);flex:none}
.mtx-drawer-title{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:14px;font-weight:600;letter-spacing:-.015em}
.mtx-drawer-body{flex:1;overflow:auto;padding:16px;display:flex;flex-direction:column;gap:16px}
.mtx-field{display:flex;flex-direction:column;gap:6px}
.mtx-field>label{font-size:12px;font-weight:500;color:var(--mtx-fg)}
.mtx-field input[type=text],.mtx-field input[type=date],.mtx-field input[type=datetime-local],.mtx-field select,.mtx-field textarea{
  width:100%;padding:7px 10px;border-radius:var(--mtx-r-sm);
  border:1px solid var(--mtx-border);background:var(--mtx-bg);color:var(--mtx-fg);
  font:inherit;font-size:12.5px;outline:none;
  transition:border-color .15s ease,box-shadow .15s ease;
}
.mtx-field input[type=text]{height:32px}
.mtx-field textarea{min-height:92px;resize:vertical;line-height:1.6}
.mtx-field input:focus,.mtx-field textarea:focus,.mtx-field select:focus{border-color:var(--mtx-primary);box-shadow:0 0 0 3px var(--mtx-ring)}
.mtx-inline{display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap}
.mtx-inline .mtx-field{flex:1;min-width:120px}
.mtx-weekdays{display:flex;gap:4px;flex-wrap:wrap}
.mtx-wd{
  width:28px;height:28px;padding:0;border-radius:999px;cursor:pointer;
  border:1px solid var(--mtx-border);background:transparent;color:inherit;
  font:inherit;font-size:11px;transition:background-color .14s ease,border-color .14s ease,color .14s ease;
}
.mtx-wd:hover{border-color:var(--mtx-border-strong)}
.mtx-wd.on{border-color:var(--mtx-primary);background:color-mix(in srgb,var(--mtx-primary) 12%,transparent);color:var(--mtx-primary);font-weight:600}
.mtx-note-small{font-size:11px;line-height:1.6;color:var(--mtx-muted)}
.mtx-steps{display:flex;flex-direction:column;gap:1px}
.mtx-step{display:flex;align-items:center;gap:8px;padding:5px 4px;border-radius:6px}
.mtx-step:hover{background:color-mix(in srgb,var(--mtx-accent) 62%,transparent)}
.mtx-step>input[type=checkbox]{width:15px;height:15px;margin:0;flex:none;border-radius:4px;accent-color:var(--mtx-primary)}
.mtx-step-name{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12.5px}
.mtx-step.done .mtx-step-name{text-decoration:line-through;color:var(--mtx-muted)}
.mtx-step .mtx-iconbtn{opacity:0}
.mtx-step:hover .mtx-iconbtn,.mtx-step:focus-within .mtx-iconbtn{opacity:1}
.mtx-files{display:flex;flex-direction:column;gap:1px}
.mtx-file{display:flex;align-items:center;gap:8px;padding:5px 4px;border-radius:6px;font-size:12.5px}
.mtx-file:hover{background:color-mix(in srgb,var(--mtx-accent) 62%,transparent)}
.mtx-file-name{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-file-size{font-size:10.5px;color:var(--mtx-muted);font-variant-numeric:tabular-nums}
.mtx-drawer-foot{display:flex;align-items:center;gap:8px;padding:12px 16px;border-top:1px solid var(--mtx-border);flex:none;background:var(--mtx-bg)}
.mtx-drawer-foot .spacer{flex:1}

/* ---------- group drawer ---------- */
.mtx-grow{display:flex;align-items:center;gap:8px}
.mtx-grow .mtx-input{height:32px;font-size:12.5px}
.mtx-assign{display:flex;flex-direction:column;gap:1px;margin-top:4px;padding-top:12px;border-top:1px solid var(--mtx-border)}
.mtx-assign-row{display:flex;align-items:center;gap:10px;padding:3px 0;font-size:12.5px}
.mtx-assign-row>span{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-select{
  max-width:180px;height:28px;padding:0 6px;border-radius:6px;
  border:1px solid var(--mtx-border);background:var(--mtx-bg);color:var(--mtx-fg);
  font:inherit;font-size:11.5px;
}
.mtx-note{margin:0;font-size:11.5px;line-height:1.6;color:var(--mtx-muted)}
.mtx-btn.danger{color:var(--mtx-destructive);border-color:color-mix(in srgb,var(--mtx-destructive) 26%,var(--mtx-border))}
.mtx-btn.danger:hover{background:color-mix(in srgb,var(--mtx-destructive) 9%,var(--mtx-bg));border-color:color-mix(in srgb,var(--mtx-destructive) 40%,var(--mtx-border))}
`

    /** Point a document-relative path at this plugin's API. */
    const url = (path) => new URL(`${API}${path}`, document.baseURI).href

    /** GET one API payload. */
    async function get(path) {
      const response = await fetch(url(path), { headers: { accept: 'application/json' } })
      const payload = await response.json()
      if (!payload.ok) throw new Error(payload.error || `HTTP ${response.status}`)
      return payload
    }

    /** POST one API payload. */
    async function post(path, body) {
      const response = await fetch(url(path), {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(body ?? {}),
      })
      const payload = await response.json()
      if (!payload.ok) throw new Error(payload.error || `HTTP ${response.status}`)
      return payload
    }

    const WEEKDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六']

    /** Short label for a due day. */
    function dueLabel(day, today) {
      if (day === undefined || day === null) return null
      if (day === today) return '今天'
      const at = new Date(`${day}T12:00:00`)
      if (Number.isNaN(at.getTime())) return day
      return `${at.getMonth() + 1}/${at.getDate()} ${WEEKDAYS[at.getDay()]}`
    }

    function IconCheck({ size = 12 }) {
      return h('svg', { width: size, height: size, viewBox: '0 0 16 16', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 2.4, strokeLinecap: 'round', strokeLinejoin: 'round' }, h('path', { d: 'M3 8.5l3.2 3.2L13 5' }))
    }

    function IconStar({ filled = false }) {
      return h('svg', { width: 15, height: 15, viewBox: '0 0 24 24', 'aria-hidden': true, fill: filled ? 'currentColor' : 'none', stroke: 'currentColor', strokeWidth: 1.7, strokeLinejoin: 'round' }, h('path', { d: 'M12 3.6l2.6 5.4 5.9.8-4.3 4.1 1 5.9-5.2-2.9-5.2 2.9 1-5.9L3.5 9.8l5.9-.8z' }))
    }

    function IconSun({ filled = false, size = 15 }) {
      return h('svg', {
        width: size,
        height: size,
        viewBox: '0 0 24 24',
        'aria-hidden': true,
        fill: 'none',
        stroke: 'currentColor',
        strokeWidth: 1.7,
        strokeLinecap: 'round',
      },
        h('circle', { cx: 12, cy: 12, r: 3.8, fill: filled ? 'currentColor' : 'none' }),
        h('path', { d: 'M12 2.6v2.2M12 19.2v2.2M2.6 12h2.2M19.2 12h2.2M5.4 5.4l1.6 1.6M17 17l1.6 1.6M18.6 5.4L17 7M7 17l-1.6 1.6' }))
    }

    function IconTrash() {
      return h('svg', { width: 15, height: 15, viewBox: '0 0 24 24', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 1.7, strokeLinecap: 'round' }, h('path', { d: 'M4 7h16M9 7V5h6v2M6.5 7l.8 12h9.4l.8-12M10 11v5M14 11v5' }))
    }

    function IconGear() {
      return h('svg', { width: 13, height: 13, viewBox: '0 0 24 24', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round', strokeLinejoin: 'round' },
        h('circle', { cx: 12, cy: 12, r: 3.2 }),
        h('path', { d: 'M19.5 14.6a1.6 1.6 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.6 1.6 0 0 0-1.8-.3 1.6 1.6 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.2a1.6 1.6 0 0 0-1-1.5 1.6 1.6 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.6 1.6 0 0 0 .3-1.8 1.6 1.6 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.2a1.6 1.6 0 0 0 1.5-1 1.6 1.6 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.6 1.6 0 0 0 1.8.3h.1a1.6 1.6 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.2a1.6 1.6 0 0 0 1 1.5 1.6 1.6 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.6 1.6 0 0 0-.3 1.8v.1a1.6 1.6 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.2a1.6 1.6 0 0 0-1.3 1z' }))
    }

    function IconRefresh() {
      return h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 1.8, strokeLinecap: 'round' }, h('path', { d: 'M20 11a8 8 0 1 0-2.3 5.7M20 5v6h-6' }))
    }

    function IconPlus() {
      return h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 2, strokeLinecap: 'round' }, h('path', { d: 'M12 5v14M5 12h14' }))
    }

    function IconClose() {
      return h('svg', { width: 14, height: 14, viewBox: '0 0 24 24', 'aria-hidden': true, fill: 'none', stroke: 'currentColor', strokeWidth: 1.9, strokeLinecap: 'round' }, h('path', { d: 'M6 6l12 12M18 6L6 18' }))
    }

    /**
     * The board: a left rail (brand, counters, lists) and the selected list's tasks on the right.
     * @param props - the panel registration's runtime share (unused).
     */
    function TodoPage() {
      const { t } = useText()
      const [phase, setPhase] = useState('loading')
      const [error, setError] = useState('')
      const [auth, setAuth] = useState(null)
      const [lists, setLists] = useState([])
      const [listId, setListId] = useState(() => {
        try { return window.localStorage.getItem(LS_LIST) || '' } catch { return '' }
      })
      const [tasks, setTasks] = useState([])
      const [tasksLoading, setTasksLoading] = useState(false)
      const [today, setToday] = useState(null)
      const [filter, setFilter] = useState('open')
      const [draft, setDraft] = useState('')
      const addRef = useRef(null)
      const [due, setDue] = useState('')
      const [busy, setBusy] = useState(false)
      const [toast, setToast] = useState('')
      const [groups, setGroups] = useState([])
      const [collapsed, setCollapsed] = useState({})
      const [sheet, setSheet] = useState(null)
      const [myday, setMyday] = useState(null)
      const [pane, setPane] = useState('list')
      const [detail, setDetail] = useState(null)
      const [meta, setMeta] = useState(null)
      const listIdRef = useRef(listId)
      listIdRef.current = listId
      const detailRef = useRef(detail)
      detailRef.current = detail

      const complain = useCallback((cause) => {
        const message = cause instanceof Error ? cause.message : String(cause)
        setToast(message)
        window.setTimeout(() => setToast(current => (current === message ? '' : current)), 6000)
      }, [])

      const loadTasks = useCallback(async (id) => {
        if (!id) { setTasks([]); return }
        setTasksLoading(true)
        try {
          const payload = await get(`tasks?listId=${encodeURIComponent(id)}&status=all`)
          setTasks(Array.isArray(payload.items) ? payload.items : [])
        } catch (cause) {
          complain(cause)
          setTasks([])
        } finally {
          setTasksLoading(false)
        }
      }, [complain])

      const loadToday = useCallback(async (force = false) => {
        try {
          setToday(await get(`today${force ? '?refresh=1' : ''}`))
        } catch (cause) {
          complain(cause)
        }
      }, [complain])

      const loadMyDay = useCallback(async (force = false) => {
        try { setMyday(await get(`myday${force ? '?refresh=1' : ''}`)) } catch (cause) { complain(cause) }
      }, [complain])

      const loadMeta = useCallback(async () => {
        try { setMeta(await get('meta')) } catch { /* the panel works without the badge */ }
      }, [])

      const bootstrap = useCallback(async () => {
        setPhase('loading')
        setError('')
        void loadMeta()
        try {
          const state = await get('state')
          setAuth(state.auth ?? null)
          if (state.auth?.state !== 'authenticated') { setPhase('unauth'); return }
          const catalog = Array.isArray(state.lists) ? state.lists : []
          setLists(catalog)
          setGroups(Array.isArray(state.groups) ? state.groups : [])
          const wanted = catalog.some(list => list.id === listIdRef.current) ? listIdRef.current : (catalog[0]?.id ?? '')
          setListId(wanted)
          try { window.localStorage.setItem(LS_LIST, wanted) } catch {}
          setPhase('ready')
          void loadTasks(wanted)
          void loadToday()
          void loadMyDay()
        } catch (cause) {
          setError(cause instanceof Error ? cause.message : String(cause))
          setPhase('error')
        }
      }, [loadMeta, loadMyDay, loadTasks, loadToday])

      useEffect(() => { void bootstrap() }, [bootstrap])

      const selectList = useCallback((id) => {
        setPane('list')
        setListId(id)
        try { window.localStorage.setItem(LS_LIST, id) } catch {}
        void loadTasks(id)
      }, [loadTasks])

      const refreshAll = useCallback(async () => {
        setBusy(true)
        try {
          await loadTasks(listIdRef.current)
          await loadToday(true)
        } finally {
          setBusy(false)
        }
      }, [loadTasks, loadToday])

      const completeTask = useCallback(async (task, owner) => {
        const listId = typeof owner === 'string' ? owner : listIdRef.current
        setBusy(true)
        try {
          await post('complete', { listId, taskId: task.id })
          await loadTasks(listIdRef.current)
          void loadToday(true)
          void loadMyDay(true)
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, loadMyDay, loadTasks, loadToday])

      const reopenTask = useCallback(async (task, owner) => {
        const listId = typeof owner === 'string' ? owner : listIdRef.current
        setBusy(true)
        try {
          await post('update', { listId, taskId: task.id, status: 'notStarted' })
          await loadTasks(listIdRef.current)
          void loadToday(true)
          void loadMyDay(true)
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, loadMyDay, loadTasks, loadToday])

      const toggleStar = useCallback(async (task, owner) => {
        const listId = typeof owner === 'string' ? owner : listIdRef.current
        const next = task.importance === 'high' ? 'normal' : 'high'
        setBusy(true)
        try {
          await post('update', { listId, taskId: task.id, importance: next })
          await loadTasks(listIdRef.current)
          void loadToday(true)
          void loadMyDay(true)
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, loadMyDay, loadTasks, loadToday])

      const removeTask = useCallback(async (task, owner) => {
        if (!window.confirm(t('deleteConfirm'))) return
        const listId = typeof owner === 'string' ? owner : listIdRef.current
        setBusy(true)
        try {
          await post('delete', { listId, taskId: task.id })
          await loadTasks(listIdRef.current)
          void loadToday(true)
          void loadMyDay(true)
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, loadMyDay, loadTasks, loadToday, t])

      const toggleMyDay = useCallback(async (task) => {
        const ids = Array.isArray(myday?.taskIds) ? myday.taskIds : []
        const on = !ids.includes(task.id)
        setBusy(true)
        try {
          setMyday(await post('myday', { taskId: task.id, on }))
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, myday])

      const openDetail = useCallback(async (task, owner) => {
        const target = typeof owner === 'string' && owner !== '' ? owner : listIdRef.current
        if (!target || !task || typeof task.id !== 'string') return
        setDetail({ listId: target, taskId: task.id, loading: true, task, notes: '', checklist: [], attachments: [] })
        try {
          const payload = await get(`detail?listId=${encodeURIComponent(target)}&taskId=${encodeURIComponent(task.id)}`)
          setDetail(current => current !== null && current.taskId === task.id
            ? {
                listId: target,
                taskId: task.id,
                loading: false,
                task: payload.task ?? task,
                notes: typeof payload.notes === 'string' ? payload.notes : '',
                checklist: Array.isArray(payload.checklist) ? payload.checklist : [],
                attachments: Array.isArray(payload.attachments) ? payload.attachments : [],
              }
            : current)
        } catch (cause) {
          complain(cause)
          setDetail(null)
        }
      }, [complain])

      /** Re-read the open detail after a change, keeping the drawer's fields. */
      const reloadDetail = useCallback(async () => {
        const current = detailRef.current
        if (current === null) return
        try {
          const payload = await get(`detail?listId=${encodeURIComponent(current.listId)}&taskId=${encodeURIComponent(current.taskId)}`)
          setDetail({
            listId: current.listId,
            taskId: current.taskId,
            loading: false,
            task: payload.task,
            notes: typeof payload.notes === 'string' ? payload.notes : '',
            checklist: Array.isArray(payload.checklist) ? payload.checklist : [],
            attachments: Array.isArray(payload.attachments) ? payload.attachments : [],
          })
        } catch (cause) { complain(cause) }
      }, [complain])

      const afterDetailChange = useCallback(() => {
        void loadTasks(listIdRef.current)
        void loadToday(true)
        void loadMyDay(true)
      }, [loadMyDay, loadTasks, loadToday])

      const addTask = useCallback(async (event) => {
        event.preventDefault()
        const title = draft.trim()
        if (title === '' || listIdRef.current === '') return
        setBusy(true)
        try {
          await post('create', { listId: listIdRef.current, title, dueDate: due })
          setDraft('')
          setDue('')
          await loadTasks(listIdRef.current)
          void loadToday(true)
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }, [complain, draft, due, loadTasks, loadToday])

      if (phase === 'loading') {
        return h('div', { className: 'mtx-page' },
          h('div', { className: 'mtx-state' },
            h('div', { className: 'mtx-logo' }, '✓'),
            h('h2', null, t('loading')),
            h('div', { style: { width: '240px', display: 'flex', flexDirection: 'column', gap: '6px' } },
              h('div', { className: 'mtx-skel' }), h('div', { className: 'mtx-skel' }))),
        )
      }

      if (phase === 'error') {
        return h('div', { className: 'mtx-page' },
          h('div', { className: 'mtx-state' },
            h('h2', null, 'Microsoft To Do'),
            h('p', null, error),
            h('button', { className: 'mtx-btn primary', onClick: () => { void bootstrap() } }, t('retry'))))
      }

      if (phase === 'unauth') {
        return h(SignInPanel, { t, onSignedIn: () => { void bootstrap() }, complain })
      }

      const todayKey = today?.today ?? ''
      const visible = tasks.filter(task => filter === 'all' ? true : filter === 'done' ? task.status === 'completed' : task.status !== 'completed')
      const counts = lists.reduce((acc, list) => {
        const row = today?.perList?.find(entry => entry.id === list.id)
        acc[list.id] = row ?? null
        return acc
      }, {})
      const recommendations = [...(today?.overdue ?? []), ...(today?.dueToday ?? [])].slice(0, 6)
      const pinned = [
        lists.find(list => list.wellknownListName === 'flaggedEmails'),
        lists.find(list => list.wellknownListName === 'defaultList'),
      ].filter(list => list !== undefined)
      const pinnedIds = new Set(pinned.map(list => list.id))
      const pinnedLabel = (list) => list.wellknownListName === 'flaggedEmails' ? t('flaggedEmails') : list.displayName
      const mydayItems = Array.isArray(myday?.items) ? myday.items : []
      const mydayIds = new Set(Array.isArray(myday?.taskIds) ? myday.taskIds : [])
      const rows = pane === 'myday'
        ? mydayItems.map(item => h(TaskRow, {
            key: item.id,
            task: { ...item, status: 'notStarted', dueDay: item.day ?? null },
            todayKey,
            t,
            busy,
            mydayOn: true,
            showList: true,
            onComplete: () => { void completeTask(item, item.listId) },
            onReopen: () => { void reopenTask(item, item.listId) },
            onStar: () => { void toggleStar(item, item.listId) },
            onDelete: () => { void removeTask(item, item.listId) },
            onMyDay: () => { void toggleMyDay(item) },
            onOpen: () => { void openDetail(item, item.listId) },
          }))
        : visible.map(task => h(TaskRow, {
            key: task.id,
            task,
            todayKey,
            t,
            busy,
            mydayOn: mydayIds.has(task.id),
            onComplete: () => { void completeTask(task) },
            onReopen: () => { void reopenTask(task) },
            onStar: () => { void toggleStar(task) },
            onDelete: () => { void removeTask(task) },
            onMyDay: () => { void toggleMyDay(task) },
            onOpen: () => { void openDetail(task) },
          }))

      const placed = new Set()
      const sections = []
      for (const group of groups) {
        const name = typeof group?.name === 'string' ? group.name.trim() : ''
        if (name === '') continue
        const ids = Array.isArray(group.listIds) ? group.listIds : []
        const members = lists.filter(list => ids.includes(list.id) && !pinnedIds.has(list.id))
        if (members.length === 0) continue
        for (const list of members) placed.add(list.id)
        sections.push({ key: `group:${name}`, name, lists: members })
      }
      const loose = lists.filter(list => !placed.has(list.id) && !pinnedIds.has(list.id))

      const renderRow = (list, sub, label) => {
        const row = counts[list.id]
        const text = label ?? list.displayName
        return h('button', {
          key: list.id,
          className: `mtx-listrow${sub ? ' sub' : ''}${list.id === listId ? ' on' : ''}`,
          onClick: () => selectList(list.id),
          title: text,
        },
          h('span', { className: `mtx-listdot${row && row.overdue > 0 ? ' late' : ''}` }),
          h('span', { className: 'mtx-listname' }, text),
          h('span', { className: 'mtx-listcount' }, row ? String(row.open) : ''))
      }

      const openSheet = () => setSheet({
        groups: groups.map(group => ({ name: group.name, listIds: [...(group.listIds ?? [])] })),
      })

      const current = lists.find(list => list.id === listId) ?? null
      const heading = pane === 'myday'
        ? t('myday')
        : current === null
          ? ''
          : current.wellknownListName === 'flaggedEmails' ? t('flaggedEmails') : current.displayName
      const headingCount = pane === 'myday'
        ? mydayItems.length
        : tasks.filter(task => task.status !== 'completed').length

      const compatState = typeof meta?.dsh?.compatible === 'boolean' ? meta.dsh.compatible : null
      const compatClass = compatState === true ? 'ok' : compatState === false ? 'bad' : 'unknown'
      const compatLabel = compatState === true ? t('compatOk') : compatState === false ? t('compatBad') : t('compatUnknown')
      const compatTitle = compatState === false
        ? `DSH ${meta.dsh.version} < ${meta.dsh.required}`
        : compatState === null ? t('compatUnknownHint') : `DSH ${meta.dsh.version}`

      return h('div', { className: 'mtx-page', style: { position: 'relative' } },
        h('div', { className: 'mtx-body' },
          h('aside', { className: 'mtx-aside' },
            h('div', { className: 'mtx-railhead' },
              h('div', { className: 'mtx-brand' },
                h('div', { className: 'mtx-logo' }, '✓'),
                h('div', { style: { minWidth: 0 } },
                  h('h1', { className: 'mtx-title' }, t('title')),
                  h('p', { className: 'mtx-sub' }, auth?.account?.name
                    ? `${t('account')} · ${auth.account.name}`
                    : t('account')))),
              h('button', {
                className: 'mtx-iconbtn',
                onClick: () => { void refreshAll() },
                disabled: busy,
                title: t('refresh'),
              }, h(IconRefresh))),
            h('div', { className: 'mtx-stats' },
              h(Stat, { label: t('today'), value: today?.dueToday?.length, tone: 'today' }),
              h(Stat, { label: t('overdue'), value: today?.overdue?.length, tone: 'overdue' }),
              h(Stat, { label: t('starred'), value: today?.starred?.length, tone: 'starred' })),
            h('div', { className: 'mtx-aside-head' }, h('span', null, t('lists')), h('span', null, String(lists.length))),
            h('button', {
              className: `mtx-listrow myday${pane === 'myday' ? ' on' : ''}`,
              onClick: () => { setPane('myday'); void loadMyDay(true) },
              title: t('myday'),
            },
              h(IconSun, { filled: pane === 'myday', size: 14 }),
              h('span', { className: 'mtx-listname' }, t('myday')),
              h('span', { className: 'mtx-listcount' }, String(mydayItems.length))),
            pinned.map(list => renderRow(list, false, pinnedLabel(list))),
            sections.map(section => {
              const open = collapsed[section.key] !== true
              return h('div', { key: section.key },
                h('button', {
                  className: 'mtx-group',
                  onClick: () => setCollapsed(prev => ({ ...prev, [section.key]: open })),
                  title: section.name,
                },
                  h('svg', {
                    className: `mtx-gchev${open ? ' open' : ''}`,
                    viewBox: '0 0 24 24',
                    'aria-hidden': true,
                    fill: 'none',
                    stroke: 'currentColor',
                    strokeWidth: 2.4,
                    strokeLinecap: 'round',
                    strokeLinejoin: 'round',
                  }, h('path', { d: 'M9 6l6 6-6 6' })),
                  h('span', { className: 'mtx-gname' }, section.name),
                  h('span', { className: 'mtx-gcount' }, String(section.lists.length))),
                open ? section.lists.map(list => renderRow(list, true)) : null)
            }),
            sections.length > 0 && loose.length > 0
              ? h('div', { className: 'mtx-aside-sub' }, t('ungrouped'))
              : null,
            loose.map(list => renderRow(list, false)),
            recommendations.length > 0
              ? h('div', { className: 'mtx-todaycard' },
                  h('h4', null, t('recommended')),
                  recommendations.map(item => h('button', {
                    key: `${item.listId}:${item.id}`,
                    className: 'mtx-mini',
                    onClick: () => { selectList(item.listId) },
                    title: item.title,
                  },
                    h('span', { className: `mtx-mini-mark${item.day && item.day < todayKey ? ' late' : ''}` }),
                    h('span', { className: 'mtx-mini-text' }, item.title),
                    h('span', { className: 'mtx-mini-list' }, item.listName))))
              : null,
            h('div', { className: 'mtx-aside-foot' },
              h('button', { className: 'mtx-footbtn', onClick: openSheet, title: t('groupSettings') },
                h(IconGear), t('groupSettings')),
              meta
                ? h('div', { className: 'mtx-verline', title: compatTitle },
                    h('span', null, `${t('version')} ${meta.plugin?.version ?? ''}`),
                    h('span', { className: `mtx-verchip ${compatClass}` }, compatLabel))
                : null)),
          h('section', { className: 'mtx-main' },
            compatState === false
              ? h('div', { className: 'mtx-warn' }, t('compatWarn')
                  .replace('{dsh}', meta.dsh.version ?? '?')
                  .replace('{required}', meta.dsh.required ?? ''))
              : null,
            pane === 'myday'
              ? null
              : h('div', { className: 'mtx-listhead' },
                  h('span', { className: 'mtx-listhead-title' }, heading),
                  h('span', { className: 'mtx-listhead-count' }, `${headingCount} ${t('open')}`)),
            pane === 'myday'
              ? h('div', { className: 'mtx-mydayhead' },
                  h('span', null, t('myday'), h('b', null, String(mydayItems.length))),
                  h('span', null, typeof myday?.date === 'string' ? myday.date : ''))
              : h('div', { className: 'mtx-toolbar' },
              h('div', { className: 'mtx-tabs' },
                h(Tab, { id: 'open', label: t('open'), current: filter, onPick: setFilter }),
                h(Tab, { id: 'done', label: t('done'), current: filter, onPick: setFilter }),
                h(Tab, { id: 'all', label: t('all'), current: filter, onPick: setFilter })),
              h('span', { className: 'mtx-hint' }, `${visible.length} / ${tasks.length}`)),
            pane === 'myday' ? null : h('form', { className: 'mtx-add', onSubmit: (event) => { void addTask(event) } },
              h('input', {
                ref: addRef,
                className: 'mtx-input',
                value: draft,
                placeholder: t('addPlaceholder'),
                onChange: (event) => setDraft(event.target.value),
                disabled: busy,
              }),
              h('input', {
                className: 'mtx-date',
                type: 'date',
                value: due,
                onChange: (event) => setDue(event.target.value),
                disabled: busy,
              }),
              h('button', { className: 'mtx-btn primary', type: 'submit', disabled: busy || draft.trim() === '' }, h(IconPlus), t('add'))),
            h('div', { className: 'mtx-list' },
              pane === 'myday'
                ? (mydayItems.length === 0
                    ? h('div', { className: 'mtx-mydayempty' },
                        h('div', null, t('mydayEmpty')),
                        h('div', null, t('mydayHint')))
                    : rows)
                : tasksLoading
                  ? [0, 1, 2].map(index => h('div', { key: index, className: 'mtx-skel' }))
                  : visible.length === 0
                    ? h('div', { className: 'mtx-empty' }, filter === 'open' ? t('emptyOpen') : t('empty'))
                    : rows)))
        , toast === '' ? null : h('div', { className: 'mtx-toast' }, toast)
        , sheet === null ? null : h(GroupSheet, {
            t,
            lists,
            value: sheet,
            onChange: setSheet,
            onClose: () => setSheet(null),
            onSave: async (next) => {
              setBusy(true)
              try {
                const payload = await post('groups', { groups: next })
                setGroups(Array.isArray(payload.groups) ? payload.groups : next)
                setSheet(null)
              } catch (cause) { complain(cause) } finally { setBusy(false) }
            },
          })
        , detail === null ? null : h(TaskDetail, {
            key: detail.taskId,
            t,
            detail,
            todayKey,
            mydayOn: mydayIds.has(detail.taskId),
            onClose: () => setDetail(null),
            onReload: reloadDetail,
            onAfter: afterDetailChange,
            onMyDay: () => { void toggleMyDay(detail.task ?? { id: detail.taskId }) },
            onComplete: () => {
              const task = detail.task ?? { id: detail.taskId }
              setDetail(null)
              void completeTask(task, detail.listId)
            },
            onDelete: () => {
              const task = detail.task ?? { id: detail.taskId }
              setDetail(null)
              void removeTask(task, detail.listId)
            },
            complain,
          }))
    }

    /**
     * Group editor: name groups, then move lists between them. The host stores
     * the result locally, because Graph cannot represent To Do list groups.
     */
    function GroupSheet({ t, lists, value, onChange, onClose, onSave }) {
      const groups = value.groups
      const clone = (next) => onChange({ groups: next })
      const rename = (index, name) => clone(groups.map((group, at) => at === index ? { ...group, name } : group))
      const drop = (index) => clone(groups.filter((_, at) => at !== index))
      const add = () => clone([...groups, { name: '', listIds: [] }])
      const assign = (listId, target) => clone(groups.map((group, at) => {
        const listIds = (group.listIds ?? []).filter(id => id !== listId)
        return at === target ? { ...group, listIds: [...listIds, listId] } : { ...group, listIds }
      }))
      const groupOf = (listId) => groups.findIndex(group => (group.listIds ?? []).includes(listId))

      useEffect(() => {
        const onKey = (event) => { if (event.key === 'Escape') onClose() }
        window.addEventListener('keydown', onKey)
        return () => { window.removeEventListener('keydown', onKey) }
      }, [onClose])

      return h('div', {
        className: 'mtx-mask',
        onMouseDown: (event) => { if (event.target === event.currentTarget) onClose() },
      },
        h('div', { className: 'mtx-drawer' },
          h('div', { className: 'mtx-drawer-head' },
            h('span', { className: 'mtx-drawer-title' }, t('groupSettings')),
            h('button', { className: 'mtx-iconbtn', onClick: onClose, title: t('cancel') }, h(IconClose))),
          h('div', { className: 'mtx-drawer-body' },
          h('p', { className: 'mtx-note' }, t('groupHint')),
          groups.map((group, index) => h('div', { className: 'mtx-grow', key: index },
            h('input', {
              className: 'mtx-input',
              value: group.name,
              placeholder: t('groupPlaceholder'),
              onChange: (event) => rename(index, event.target.value),
            }),
            h('button', { className: 'mtx-btn', onClick: () => drop(index), title: t('deleteGroup') }, h(IconTrash)))),
          h('div', { className: 'mtx-grow' },
            h('button', { className: 'mtx-btn', onClick: add }, h(IconPlus), t('newGroup'))),
          h('div', { className: 'mtx-assign' },
            lists.map(list => {
              const at = groupOf(list.id)
              return h('div', { className: 'mtx-assign-row', key: list.id },
                h('span', { title: list.displayName }, list.displayName),
                h('select', {
                  className: 'mtx-select',
                  value: at < 0 ? '' : String(at),
                  onChange: (event) => assign(list.id, event.target.value === '' ? -1 : Number(event.target.value)),
                },
                  h('option', { value: '' }, t('noGroup')),
                  groups.map((group, index) => h('option', { key: index, value: String(index) }, group.name.trim() === '' ? t('newGroup') : group.name))))
            })),
          ),
          h('div', { className: 'mtx-drawer-foot' },
            h('span', { className: 'spacer' }),
            h('button', { className: 'mtx-btn', onClick: onClose }, t('cancel')),
            h('button', {
              className: 'mtx-btn primary',
              onClick: () => { void onSave(groups.filter(group => group.name.trim() !== '').map(group => ({ name: group.name.trim(), listIds: group.listIds ?? [] }))) },
            }, t('save')))))
    }

    function Stat({ label, value, tone }) {
      return h('div', { className: `mtx-stat ${tone}` },
        h('div', { className: 'mtx-stat-label' }, label),
        h('div', { className: 'mtx-stat-value' }, value === undefined || value === null ? '—' : String(value)))
    }

    function Tab({ id, label, current, onPick }) {
      return h('button', {
        type: 'button',
        className: `mtx-tab${current === id ? ' on' : ''}`,
        onClick: () => onPick(id),
      }, label)
    }

    /** Monday-first order used by Graph's pattern.daysOfWeek. */
    const WDAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']

    /** A Graph weekday name in the shell's language (2024-01-01 was a Monday). */
    function dayShort(name) {
      const at = WDAYS.indexOf(name)
      const base = new Date(2024, 0, 1 + (at < 0 ? 0 : at))
      return new Intl.DateTimeFormat(undefined, { weekday: 'short' }).format(base)
    }

    /** The recurrence kind a task currently has, in the editor's vocabulary. */
    function recurrenceType(recurrence) {
      const type = recurrence?.pattern?.type
      if (type === 'daily' || type === 'weekly' || type === 'absoluteMonthly') return type
      if (type === 'relativeMonthly') return 'absoluteMonthly'
      return 'none'
    }

    /** Local file to bare base64, as the host's upload route expects it. */
    function fileToBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result ?? '').replace(/^data:[^,]*,/, ''))
        reader.onerror = () => reject(reader.error ?? new Error('could not read the file'))
        reader.readAsDataURL(file)
      })
    }

    function base64ToBlob(data, type) {
      const binary = atob(data)
      const bytes = new Uint8Array(binary.length)
      for (let at = 0; at < binary.length; at += 1) bytes[at] = binary.charCodeAt(at)
      return new Blob([bytes], { type })
    }

    function formatSize(size) {
      const bytes = Number(size ?? 0)
      if (!Number.isFinite(bytes) || bytes <= 0) return ''
      if (bytes < 1024) return `${bytes} B`
      if (bytes < 1024 * 1024) return `${Math.round(bytes / 102.4) / 10} KB`
      return `${Math.round(bytes / 104857.6) / 10} MB`
    }

    /**
     * One task's detail drawer: notes, dates, reminder, repeat, steps and files.
     * Steps and files go through the host's Graph bridge because the MCP tool
     * surface covers neither recurrence nor attachments.
     */
    function TaskDetail({ t, detail, todayKey, mydayOn, onClose, onReload, onAfter, onMyDay, onComplete, onDelete, complain }) {
      const task = detail.task ?? {}
      const [title, setTitle] = useState(typeof task.title === 'string' ? task.title : '')
      const [notes, setNotes] = useState(detail.notes ?? '')
      const [dueDate, setDueDate] = useState(() => String(task.dueDateTime?.dateTime ?? '').slice(0, 10))
      const [reminderAt, setReminderAt] = useState(() => {
        const at = String(task.reminderDateTime?.dateTime ?? '')
        return at === '' ? '' : at.slice(0, 16)
      })
      const [repeat, setRepeat] = useState(() => recurrenceType(task.recurrence))
      const [repeatEvery, setRepeatEvery] = useState(() => String(task.recurrence?.pattern?.interval ?? 1))
      const [repeatDays, setRepeatDays] = useState(() => new Set(Array.isArray(task.recurrence?.pattern?.daysOfWeek) ? task.recurrence.pattern.daysOfWeek : ['monday']))
      const [endKind, setEndKind] = useState(() => {
        const kind = task.recurrence?.range?.type
        return kind === 'endDate' || kind === 'numbered' ? kind : 'noEnd'
      })
      const [endDate, setEndDate] = useState(() => String(task.recurrence?.range?.endDate ?? '').slice(0, 10))
      const [stepDraft, setStepDraft] = useState('')
      const [saving, setSaving] = useState(false)
      const [uploading, setUploading] = useState(false)
      const fileInput = useRef(null)

      const save = async () => {
        setSaving(true)
        try {
          const payload = {
            listId: detail.listId,
            taskId: detail.taskId,
            title,
            bodyText: notes,
            dueDate,
            importance: ['low', 'normal', 'high'].includes(task.importance) ? task.importance : 'normal',
          }
          payload.reminderAt = reminderAt
          payload.recurrence = repeat === 'none'
            ? null
            : {
                pattern: {
                  type: repeat,
                  interval: Math.max(1, Math.min(99, Number(repeatEvery) || 1)),
                  daysOfWeek: [...repeatDays],
                },
                range: {
                  type: endKind,
                  startDate: dueDate !== '' ? dueDate : todayKey,
                  endDate: endKind === 'endDate' ? (endDate !== '' ? endDate : (dueDate || todayKey)) : undefined,
                },
              }
          await post('save', payload)
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) } finally { setSaving(false) }
      }

      const toggleRepeatDay = (name) => {
        setRepeatDays(previous => {
          const next = new Set(previous)
          if (next.has(name)) next.delete(name)
          else next.add(name)
          if (next.size === 0) next.add(name)
          return next
        })
      }

      const addStep = async () => {
        const displayName = stepDraft.trim()
        if (displayName === '') return
        try {
          await post('steps', { action: 'add', listId: detail.listId, taskId: detail.taskId, displayName })
          setStepDraft('')
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) }
      }

      const updateStep = async (item, patch) => {
        try {
          await post('steps', { action: 'update', listId: detail.listId, taskId: detail.taskId, itemId: item.id, ...patch })
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) }
      }

      const removeStep = async (item) => {
        try {
          await post('steps', { action: 'delete', listId: detail.listId, taskId: detail.taskId, itemId: item.id })
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) }
      }

      const promoteStep = async (item) => {
        try {
          await post('steps', { action: 'promote', listId: detail.listId, taskId: detail.taskId, itemId: item.id })
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) }
      }

      const uploadFile = async (event) => {
        const file = event.target.files && event.target.files[0]
        event.target.value = ''
        if (!file) return
        if (file.size > 25 * 1024 * 1024) { complain(new Error(t('fileTooLarge'))); return }
        setUploading(true)
        try {
          const dataBase64 = await fileToBase64(file)
          await post('attachment', {
            action: 'upload',
            listId: detail.listId,
            taskId: detail.taskId,
            name: file.name,
            contentType: file.type === '' ? 'application/octet-stream' : file.type,
            dataBase64,
          })
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) } finally { setUploading(false) }
      }

      const downloadFile = async (attachment) => {
        try {
          const payload = await get(`attachment?listId=${encodeURIComponent(detail.listId)}&taskId=${encodeURIComponent(detail.taskId)}&attachmentId=${encodeURIComponent(attachment.id)}`)
          const blob = base64ToBlob(payload.dataBase64 ?? '', payload.contentType ?? 'application/octet-stream')
          const url = URL.createObjectURL(blob)
          const anchor = document.createElement('a')
          anchor.href = url
          anchor.download = payload.name ?? attachment.name ?? 'attachment'
          anchor.click()
          window.setTimeout(() => URL.revokeObjectURL(url), 4000)
        } catch (cause) { complain(cause) }
      }

      const removeFile = async (attachment) => {
        try {
          await post('attachment', { action: 'delete', listId: detail.listId, taskId: detail.taskId, attachmentId: attachment.id })
          await onReload()
          onAfter()
        } catch (cause) { complain(cause) }
      }

      const checklist = Array.isArray(detail.checklist) ? detail.checklist : []
      const files = Array.isArray(detail.attachments) ? detail.attachments : []

      useEffect(() => {
        const onKey = (event) => { if (event.key === 'Escape') onClose() }
        window.addEventListener('keydown', onKey)
        return () => { window.removeEventListener('keydown', onKey) }
      }, [onClose])

      return h('div', { className: 'mtx-mask', onMouseDown: (event) => { if (event.target === event.currentTarget) onClose() } },
        h('div', { className: 'mtx-drawer' },
          h('div', { className: 'mtx-drawer-head' },
            h('span', { className: 'mtx-drawer-title', title }, title),
            detail.task?.status === 'completed' ? h('span', { className: 'mtx-chip' }, t('done')) : null,
            h('button', { type: 'button', className: 'mtx-iconbtn', onClick: onClose, title: t('cancel') }, '✕')),
          h('div', { className: 'mtx-drawer-body' },
            h('div', { className: 'mtx-field' },
              h('input', {
                type: 'text',
                value: title,
                onChange: (event) => setTitle(event.target.value),
                onKeyDown: (event) => { if (event.key === 'Enter') { event.preventDefault(); void save() } },
              })),
            h('div', { className: 'mtx-field' },
              h('label', null, t('notes')),
              h('textarea', {
                value: notes,
                placeholder: t('notesPlaceholder'),
                onChange: (event) => setNotes(event.target.value),
              })),
            h('div', { className: 'mtx-inline' },
              h('div', { className: 'mtx-field' },
                h('label', null, t('dueDate')),
                h('input', { type: 'date', value: dueDate, onChange: (event) => setDueDate(event.target.value) })),
              h('div', { className: 'mtx-field' },
                h('label', null, t('reminder')),
                h('input', { type: 'datetime-local', value: reminderAt, onChange: (event) => setReminderAt(event.target.value) }))),
            h('div', { className: 'mtx-field' },
              h('label', null, t('recurrence')),
              h('div', { className: 'mtx-inline' },
                h('select', {
                  value: repeat,
                  onChange: (event) => setRepeat(event.target.value),
                  style: { flex: 1 },
                },
                  h('option', { value: 'none' }, t('repeatNone')),
                  h('option', { value: 'daily' }, t('repeatDaily')),
                  h('option', { value: 'weekly' }, t('repeatWeekly')),
                  h('option', { value: 'absoluteMonthly' }, t('repeatMonthly'))),
                repeat === 'none' ? null : h('input', {
                  type: 'text',
                  inputMode: 'numeric',
                  value: repeatEvery,
                  onChange: (event) => setRepeatEvery(event.target.value.replace(/[^0-9]/g, '')),
                  style: { width: 54 },
                  title: t('repeatInterval'),
                })),
              repeat === 'weekly'
                ? h('div', { className: 'mtx-weekdays' },
                    WDAYS.map(name => h('button', {
                      key: name,
                      type: 'button',
                      className: `mtx-wd${repeatDays.has(name) ? ' on' : ''}`,
                      onClick: () => toggleRepeatDay(name),
                    }, dayShort(name))))
                : null,
              repeat === 'none'
                ? null
                : h('div', { className: 'mtx-inline' },
                    h('select', {
                      value: endKind,
                      onChange: (event) => setEndKind(event.target.value),
                      style: { flex: 1 },
                    },
                      h('option', { value: 'noEnd' }, t('repeatNoEnd')),
                      h('option', { value: 'endDate' }, t('repeatEndDate'))),
                    endKind === 'endDate'
                      ? h('input', { type: 'date', value: endDate, onChange: (event) => setEndDate(event.target.value), style: { flex: 1 } })
                      : null),
              repeat === 'none' ? null : h('div', { className: 'mtx-note-small' }, t('repeatHint'))),
            h('div', { className: 'mtx-field' },
              h('label', null, t('steps')),
              h('div', { className: 'mtx-steps' },
                checklist.map(item => h('div', { key: item.id, className: `mtx-step${item.isChecked ? ' done' : ''}` },
                  h('input', {
                    type: 'checkbox',
                    checked: item.isChecked === true,
                    onChange: () => { void updateStep(item, { isChecked: item.isChecked !== true }) },
                  }),
                  h('span', { className: 'mtx-step-name' }, item.displayName),
                  h('button', {
                    type: 'button',
                    className: 'mtx-iconbtn',
                    title: t('promote'),
                    onClick: () => { void promoteStep(item) },
                  }, '↗'),
                  h('button', {
                    type: 'button',
                    className: 'mtx-iconbtn',
                    title: t('remove'),
                    onClick: () => { void removeStep(item) },
                  }, h(IconTrash))))),
              h('div', { className: 'mtx-inline' },
                h('input', {
                  type: 'text',
                  value: stepDraft,
                  placeholder: t('stepPlaceholder'),
                  style: { flex: 1 },
                  onChange: (event) => setStepDraft(event.target.value),
                  onKeyDown: (event) => { if (event.key === 'Enter') { event.preventDefault(); void addStep() } },
                }),
                h('button', { type: 'button', className: 'mtx-btn', onClick: () => { void addStep() } }, t('add')))),
            h('div', { className: 'mtx-field' },
              h('label', null, t('attachments')),
              h('div', { className: 'mtx-files' },
                files.length === 0
                  ? h('div', { className: 'mtx-note-small' }, t('noAttachments'))
                  : files.map(file => h('div', { key: file.id, className: 'mtx-file' },
                      h('button', {
                        type: 'button',
                        className: 'mtx-link',
                        onClick: () => { void downloadFile(file) },
                        title: file.name,
                      }, h('span', { className: 'mtx-file-name' }, file.name)),
                      h('span', { className: 'mtx-file-size' }, formatSize(file.size)),
                      h('button', {
                        type: 'button',
                        className: 'mtx-iconbtn',
                        title: t('remove'),
                        onClick: () => { void removeFile(file) },
                      }, h(IconTrash))))),
              h('input', {
                ref: fileInput,
                type: 'file',
                style: { display: 'none' },
                onChange: (event) => { void uploadFile(event) },
              }),
              h('button', {
                type: 'button',
                className: 'mtx-btn',
                disabled: uploading,
                onClick: () => { if (fileInput.current !== null) fileInput.current.click() },
              }, uploading ? t('uploading') : t('attach')))),
          h('div', { className: 'mtx-drawer-foot' },
            h('button', {
              type: 'button',
              className: `mtx-btn${mydayOn ? ' primary' : ''}`,
              onClick: onMyDay,
              title: mydayOn ? t('removeMyDay') : t('addMyDay'),
            }, t('myday')),
            h('span', { className: 'spacer' }),
            h('button', { type: 'button', className: 'mtx-btn danger', onClick: onDelete, title: t('deleteTask') }, h(IconTrash), t('deleteTask')),
            h('button', { type: 'button', className: 'mtx-btn', onClick: onComplete }, t('doneLabel')),
            h('button', {
              type: 'button',
              className: 'mtx-btn primary',
              disabled: saving,
              onClick: () => { void save() },
            }, saving ? t('saving') : t('save')))))
    }

    function TaskRow({ task, todayKey, t, busy, onComplete, onReopen, onStar, onDelete, onMyDay, onOpen, mydayOn = false, showList = false }) {
      const done = task.status === 'completed'
      const day = typeof task.dueDay === 'string' ? task.dueDay : null
      const label = dueLabel(day, todayKey)
      const late = label !== null && day !== null && day < todayKey
      const noteText = String(task.body?.content ?? '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()
      return h('div', { className: `mtx-row${done ? ' done' : ''}` },
        h('button', {
          type: 'button',
          className: `mtx-check${done ? ' on' : ''}`,
          title: done ? t('reopen') : t('markDone'),
          disabled: busy,
          onClick: done ? onReopen : onComplete,
        }, h(IconCheck)),
        h('div', { className: 'mtx-row-main', onClick: onOpen, title: t('detail') },
          h('div', { className: 'mtx-row-title' }, task.title),
          (label !== null || task.importance === 'high' || noteText !== '' || (showList && task.listName))
            ? h('div', { className: 'mtx-row-meta' },
                showList && task.listName ? h('span', { className: 'mtx-chip' }, task.listName) : null,
                label === null ? null : h('span', { className: `mtx-chip${late ? ' late' : ''}${day === todayKey ? ' today' : ''}` }, label),
                task.importance === 'high' ? h('span', { className: 'mtx-chip today' }, '★ ' + t('starred')) : null,
                noteText === '' ? null : h('span', { className: 'mtx-chip' }, noteText.slice(0, 60)))
            : null),
        h('div', { className: 'mtx-row-actions' },
          h('button', {
            type: 'button',
            className: `mtx-iconbtn sun${mydayOn ? ' on' : ''}`,
            title: mydayOn ? t('removeMyDay') : t('addMyDay'),
            disabled: busy,
            onClick: onMyDay,
          }, h(IconSun, { filled: mydayOn })),
          h('button', {
            type: 'button',
            className: `mtx-iconbtn star${task.importance === 'high' ? ' on' : ''}`,
            title: task.importance === 'high' ? t('unstar') : t('star'),
            disabled: busy,
            onClick: onStar,
          }, h(IconStar, { filled: task.importance === 'high' })),
          h('button', {
            type: 'button',
            className: 'mtx-iconbtn',
            title: t('deleteTask'),
            disabled: busy,
            onClick: onDelete,
          }, h(IconTrash))))
    }

    /** Device-code sign-in, polling the host until the account is connected. */
    function SignInPanel({ t, onSignedIn, complain }) {
      const [pending, setPending] = useState(null)
      const [busy, setBusy] = useState(false)
      const [copied, setCopied] = useState(false)

      useEffect(() => {
        if (pending === null) return undefined
        const timer = window.setInterval(async () => {
          try {
            const state = await get('state')
            if (state.auth?.state === 'authenticated') onSignedIn()
          } catch {}
        }, 4000)
        return () => window.clearInterval(timer)
      }, [pending, onSignedIn])

      const start = async () => {
        setBusy(true)
        try {
          const payload = await post('sign-in', {})
          setPending(payload.signIn ?? {})
        } catch (cause) { complain(cause) } finally { setBusy(false) }
      }

      const copy = async () => {
        const code = pending?.userCode ?? pending?.user_code ?? ''
        if (code === '') return
        try { await navigator.clipboard.writeText(code); setCopied(true) } catch {}
      }

      const code = pending?.userCode ?? pending?.user_code ?? ''
      const link = pending?.verificationUri ?? pending?.verification_uri ?? 'https://microsoft.com/devicelogin'

      return h('div', { className: 'mtx-page' },
        h('div', { className: 'mtx-state' },
          h('div', { className: 'mtx-logo' }, '✓'),
          h('h2', null, t('signInTitle')),
          h('p', null, t('signInBody')),
          pending === null
            ? h('button', { className: 'mtx-btn primary', onClick: () => { void start() }, disabled: busy }, t('signInButton'))
            : h('div', { style: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px' } },
                h('p', null, t('signInHint')),
                h('div', { className: 'mtx-code', onClick: () => { void copy() }, title: t('codeCopied') }, code),
                h('a', { className: 'mtx-link', href: link, target: '_blank', rel: 'noreferrer' }, link),
                h('p', null, t('waiting')))))
    }

    /** Sidebar glyph: a check-list, drawn in the sidebar's currentColor. */
    function PanelIcon({ size = 20 }) {
      return h('svg', {
        width: size,
        height: size,
        viewBox: '0 0 24 24',
        'aria-hidden': true,
        fill: 'none',
        stroke: 'currentColor',
        strokeWidth: 1.9,
        strokeLinecap: 'round',
        strokeLinejoin: 'round',
      },
        h('path', { d: 'M4 6.5l1.6 1.6L8.4 5' }),
        h('path', { d: 'M11.5 6.5H20' }),
        h('path', { d: 'M4 12.5l1.6 1.6 2.8-3.1' }),
        h('path', { d: 'M11.5 12.5H20' }),
        h('path', { d: 'M4 18.5l1.6 1.6 2.8-3.1' }),
        h('path', { d: 'M11.5 18.5H20' }))
    }

    /** Shell-locale binding, captured once at apply time. */
    let boundT = null

    /** Translator bound to the shell locale, falling back to Chinese. */
    function useText() {
      const t = useCallback((key) => {
        if (boundT !== null) {
          try {
            const value = boundT(key)
            if (typeof value === 'string' && value !== '' && value !== key) return value
          } catch {}
        }
        return DICT.zh[key] ?? key
      }, [])
      return { t }
    }

    return {
      inject: ['slots', 'locale'],
      apply(ctx) {
        const style = document.createElement('style')
        style.dataset.plugin = 'dsh-microsoft-todo'
        style.dataset.pluginCss = 'dsh-microsoft-todo/panel.css'
        style.textContent = CSS
        document.head.appendChild(style)

        let bound = null
        if (typeof ctx.locale?.register === 'function') {
          ctx.effect(() => ctx.locale.register(NS, DICT), 'microsoft-todo: dictionaries')
          try { bound = ctx.locale.bind(NS) } catch { bound = null }
        }
        boundT = bound
        const label = () => (bound ? bound('panel') : DICT.zh.panel)

        ctx.slots.inject('main', () => ctx.slots.register({
          name: 'main',
          key: PANEL_ID,
          locale: NS,
        }, TodoPage))

        ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
          name: 'sidebar.panellist',
          id: PANEL_ID,
          order: 20,
          locale: NS,
          label,
        }, PanelIcon))

        ctx.effect(() => () => { style.remove() }, 'microsoft-todo: panel stylesheet')
      },
    }
  },
})

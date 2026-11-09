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
      zh: { panel: '微软待办', title: '微软待办', addPlaceholder: '添加任务，回车保存…', add: '添加', refresh: '刷新', today: '今日', overdue: '逾期', starred: '重要', lists: '清单', open: '未完成', done: '已完成', all: '全部', empty: '这个清单还没有任务', emptyOpen: '没有未完成的任务，清爽！', loading: '正在连接 Microsoft To Do…', retry: '重试', signInTitle: '连接你的 Microsoft To Do', signInBody: '需要在微软登录一次，之后由本机保存的令牌自动续期。', signInButton: '开始登录', signInHint: '打开下面的链接，输入代码完成授权：', codeCopied: '代码已复制', waiting: '等待你在浏览器里完成授权…', deleteConfirm: '删除这条任务？', failed: '操作失败', account: '已连接', star: '标为重要', unstar: '取消重要', markDone: '标记完成', reopen: '恢复未完成', recommended: '今日优先', groups: '分组', ungrouped: '未分组', groupSettings: '分组设置', newGroup: '新建分组', groupPlaceholder: '分组名称', save: '保存', cancel: '取消', deleteGroup: '删除分组', noGroup: '不分组', groupHint: '分组只保存在本机：微软 Graph 不提供 To Do 清单的分组信息，因此这里的分组不会同步到其他设备。', myday: '我的一天', mydayEmpty: '今天还没挑任务', mydayHint: '点任务右侧的太阳图标，把任务放进「我的一天」；每天自动清空。', addMyDay: '加入我的一天', removeMyDay: '移出我的一天' },
      en: { panel: 'To Do', title: 'Microsoft To Do', addPlaceholder: 'Add a task, press Enter…', add: 'Add', refresh: 'Refresh', today: 'Today', overdue: 'Overdue', starred: 'Starred', lists: 'Lists', open: 'Open', done: 'Completed', all: 'All', empty: 'No tasks in this list yet', emptyOpen: 'Nothing open here — all clear.', loading: 'Connecting to Microsoft To Do…', retry: 'Retry', signInTitle: 'Connect your Microsoft To Do', signInBody: 'Sign in once; the token cached on this machine refreshes itself afterwards.', signInButton: 'Sign in', signInHint: 'Open the link below and enter the code:', codeCopied: 'Code copied', waiting: 'Waiting for you to finish signing in…', deleteConfirm: 'Delete this task?', failed: 'Action failed', account: 'Connected', star: 'Mark as important', unstar: 'Remove importance', markDone: 'Mark complete', reopen: 'Reopen', recommended: 'Today first', groups: 'Groups', ungrouped: 'Ungrouped', groupSettings: 'Manage groups', newGroup: 'New group', groupPlaceholder: 'Group name', save: 'Save', cancel: 'Cancel', deleteGroup: 'Delete group', noGroup: 'No group', groupHint: 'Groups are stored on this machine only: the Microsoft Graph To Do API does not expose list groups, so they do not sync to other devices.', myday: 'My Day', mydayEmpty: 'Nothing picked for today yet', mydayHint: 'Use the sun button on a task to add it to My Day; the list resets every day.', addMyDay: 'Add to My Day', removeMyDay: 'Remove from My Day' },
    }

    const CSS = `
.mtx-page{display:flex;flex-direction:column;height:100%;min-height:0;background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);font-size:13px}
.mtx-head{display:flex;align-items:center;gap:16px;padding:18px 22px;border-bottom:1px solid var(--dsw-alias-border-l2);flex:none}
.mtx-brand{display:flex;align-items:center;gap:12px;min-width:0}
.mtx-logo{width:34px;height:34px;border-radius:9px;background:#2564CF;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:17px;flex:none}
.mtx-title{font-size:16px;font-weight:650;line-height:1.3;margin:0}
.mtx-sub{color:var(--dsw-alias-label-secondary);font-size:11.5px;margin:2px 0 0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-stats{display:flex;gap:8px;margin-left:auto}
.mtx-stat{border:1px solid var(--dsw-alias-border-l2);border-radius:10px;padding:6px 12px;min-width:66px;background:var(--dsw-alias-bg-layer-2)}
.mtx-stat-label{color:var(--dsw-alias-label-secondary);font-size:10.5px;letter-spacing:.04em}
.mtx-stat-value{font-size:18px;font-weight:680;line-height:1.15;font-variant-numeric:tabular-nums}
.mtx-stat.today .mtx-stat-value{color:#2564CF}
.mtx-stat.overdue .mtx-stat-value{color:#C4314B}
.mtx-stat.starred .mtx-stat-value{color:#B07C0A}
.mtx-btn{border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);border-radius:8px;padding:5px 11px;font:inherit;font-size:12px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.mtx-btn:hover{background:var(--dsw-alias-interactive-bg-hover);border-color:var(--dsw-alias-border-l4)}
.mtx-btn:disabled{opacity:.45;cursor:default}
.mtx-btn.primary{background:#2564CF;border-color:#2564CF;color:#fff}
.mtx-btn.primary:hover{background:#1d4fa6}
.mtx-body{flex:1;display:flex;min-height:0}
.mtx-aside{width:242px;flex:none;border-right:1px solid var(--dsw-alias-border-l2);display:flex;flex-direction:column;min-height:0;overflow-y:auto;padding:12px 10px 16px}
.mtx-aside-head{color:var(--dsw-alias-label-secondary);font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;padding:4px 8px 8px;display:flex;justify-content:space-between}
.mtx-listrow{display:flex;align-items:center;gap:8px;width:100%;text-align:left;border:0;background:transparent;color:inherit;font:inherit;padding:7px 9px;border-radius:8px;cursor:pointer}
.mtx-listrow:hover{background:var(--dsw-alias-interactive-bg-hover)}
.mtx-listrow.on{background:color-mix(in srgb,#2564CF 14%,transparent);color:var(--dsw-alias-label-primary);font-weight:600}
.mtx-listname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-listcount{color:var(--dsw-alias-label-secondary);font-size:11px;font-variant-numeric:tabular-nums;flex:none}
.mtx-listdot{width:6px;height:6px;border-radius:50%;background:#2564CF;flex:none;opacity:.55}
.mtx-listdot.late{background:#C4314B;opacity:1}
.mtx-todaycard{margin-top:14px;border-top:1px solid var(--dsw-alias-border-l2);padding-top:10px}
.mtx-todaycard h4{margin:0 0 6px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--dsw-alias-label-secondary);padding:0 8px}
.mtx-mini{display:flex;align-items:flex-start;gap:7px;width:100%;text-align:left;border:0;background:transparent;color:inherit;font:inherit;font-size:12px;padding:5px 8px;border-radius:7px;cursor:pointer;line-height:1.4}
.mtx-mini:hover{background:var(--dsw-alias-interactive-bg-hover)}
.mtx-mini-mark{width:5px;height:5px;border-radius:50%;margin-top:6px;flex:none;background:#2564CF}
.mtx-mini-mark.late{background:#C4314B}
.mtx-mini-text{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-mini-list{color:var(--dsw-alias-label-secondary);font-size:10.5px;flex:none;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-main{flex:1;min-width:0;display:flex;flex-direction:column;min-height:0}
.mtx-toolbar{display:flex;align-items:center;gap:12px;padding:12px 22px 6px;flex:none}
.mtx-tabs{display:inline-flex;background:var(--dsw-alias-bg-layer-2);border:1px solid var(--dsw-alias-border-l2);border-radius:9px;padding:2px;gap:2px}
.mtx-tab{border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:12px;padding:4px 11px;border-radius:7px;cursor:pointer}
.mtx-tab.on{background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-primary);font-weight:600;box-shadow:0 1px 3px rgba(0,0,0,.12)}
.mtx-hint{color:var(--dsw-alias-label-secondary);font-size:11.5px;margin-left:auto}
.mtx-add{display:flex;gap:8px;padding:6px 22px 12px;flex:none}
.mtx-input{flex:1;min-width:0;border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);border-radius:8px;padding:8px 11px;font:inherit;font-size:13px;outline:none}
.mtx-input:focus{border-color:#2564CF;box-shadow:0 0 0 3px color-mix(in srgb,#2564CF 18%,transparent)}
.mtx-date{border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-secondary);border-radius:8px;padding:0 8px;font:inherit;font-size:12px}
.mtx-list{flex:1;min-height:0;overflow-y:auto;padding:0 22px 28px;display:flex;flex-direction:column;gap:6px}
.mtx-row{display:flex;align-items:flex-start;gap:10px;border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);border-radius:10px;padding:9px 11px}
.mtx-row:hover{border-color:var(--dsw-alias-border-l4)}
.mtx-row.done{opacity:.62}
.mtx-check{width:19px;height:19px;border-radius:50%;border:1.6px solid var(--dsw-alias-border-l4);background:transparent;cursor:pointer;flex:none;margin-top:1px;display:flex;align-items:center;justify-content:center;color:transparent;padding:0}
.mtx-check:hover{border-color:#2564CF;color:#2564CF}
.mtx-check.on{background:#2564CF;border-color:#2564CF;color:#fff}
.mtx-row-main{flex:1;min-width:0}
.mtx-row-title{line-height:1.45;overflow-wrap:anywhere}
.mtx-row.done .mtx-row-title{text-decoration:line-through;color:var(--dsw-alias-label-secondary)}
.mtx-row-meta{display:flex;flex-wrap:wrap;gap:6px;margin-top:5px}
.mtx-chip{font-size:10.5px;border-radius:999px;padding:1px 8px;background:var(--dsw-alias-bg-layer-1);color:var(--dsw-alias-label-secondary);border:1px solid var(--dsw-alias-border-l2)}
.mtx-chip.late{color:#C4314B;border-color:color-mix(in srgb,#C4314B 40%,transparent);background:color-mix(in srgb,#C4314B 9%,transparent)}
.mtx-chip.today{color:#2564CF;border-color:color-mix(in srgb,#2564CF 40%,transparent);background:color-mix(in srgb,#2564CF 9%,transparent)}
.mtx-row-actions{display:flex;gap:2px;flex:none}
.mtx-iconbtn{border:0;background:transparent;color:var(--dsw-alias-label-secondary);cursor:pointer;padding:4px;border-radius:6px;display:flex}
.mtx-iconbtn:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}
.mtx-iconbtn.star.on{color:#B07C0A}
.mtx-iconbtn.sun.on{color:#C2700A;background:rgba(194,112,10,.12)}
.mtx-myday{font-weight:600}
.mtx-myday .mtx-listname{font-weight:600}
.mtx-mydayhead{color:var(--dsw-alias-label-secondary);font-size:11px;padding:2px 2px 10px;display:flex;justify-content:space-between;align-items:baseline}
.mtx-mydayhead b{color:var(--dsw-alias-label-primary);font-size:15px;font-weight:600;margin-left:6px}
.mtx-mydayempty{margin:auto;max-width:340px;color:var(--dsw-alias-label-secondary);font-size:12.5px;text-align:center;line-height:1.7;padding:48px 0}
.mtx-empty{margin:auto;color:var(--dsw-alias-label-secondary);font-size:12.5px;text-align:center;padding:40px 0}
.mtx-state{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;height:100%;padding:32px;text-align:center}
.mtx-state h2{margin:0;font-size:16px}
.mtx-state p{margin:0;color:var(--dsw-alias-label-secondary);max-width:420px;line-height:1.6;font-size:12.5px}
.mtx-code{font-family:var(--ds-font-family-code,ui-monospace,monospace);font-size:22px;letter-spacing:.22em;padding:10px 16px;border:1px dashed #2564CF;border-radius:10px;color:#2564CF;cursor:pointer;background:color-mix(in srgb,#2564CF 7%,transparent)}
.mtx-link{color:#2564CF;text-decoration:none;font-size:12.5px;border-bottom:1px solid color-mix(in srgb,#2564CF 40%,transparent)}
.mtx-toast{position:absolute;bottom:18px;left:50%;transform:translateX(-50%);background:#C4314B;color:#fff;border-radius:8px;padding:7px 14px;font-size:12px;box-shadow:0 6px 20px rgba(0,0,0,.24);max-width:80%;overflow-wrap:anywhere}
.mtx-skel{height:44px;border-radius:10px;background:linear-gradient(90deg,var(--dsw-alias-bg-layer-2),var(--dsw-alias-bg-layer-1),var(--dsw-alias-bg-layer-2));background-size:200% 100%;animation:mtx-shimmer 1.3s linear infinite}
@keyframes mtx-shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}
.mtx-group{display:flex;align-items:center;gap:6px;width:100%;border:0;background:transparent;color:inherit;font:inherit;padding:6px 8px;border-radius:7px;cursor:pointer;margin-top:6px}
.mtx-group:hover{background:var(--dsw-alias-interactive-bg-hover)}
.mtx-gchev{width:11px;height:11px;flex:none;opacity:.65;transition:transform .15s ease}
.mtx-gchev.open{transform:rotate(90deg)}
.mtx-gname{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;font-weight:600;color:var(--dsw-alias-label-secondary);text-align:left}
.mtx-gcount{font-size:10.5px;color:var(--dsw-alias-label-secondary);opacity:.75;font-variant-numeric:tabular-nums}
.mtx-listrow.sub{padding-left:24px}
.mtx-aside-sub{color:var(--dsw-alias-label-secondary);font-size:10px;letter-spacing:.08em;text-transform:uppercase;padding:10px 8px 4px}
.mtx-aside-foot{margin-top:auto;padding-top:10px}
.mtx-footbtn{display:flex;align-items:center;gap:6px;width:100%;border:0;background:transparent;color:var(--dsw-alias-label-secondary);font:inherit;font-size:11.5px;padding:6px 8px;border-radius:7px;cursor:pointer}
.mtx-footbtn:hover{background:var(--dsw-alias-interactive-bg-hover);color:var(--dsw-alias-label-primary)}
.mtx-modal{position:fixed;inset:0;background:rgba(0,0,0,.44);display:flex;align-items:center;justify-content:center;z-index:80;padding:20px}
.mtx-sheet{width:min(540px,100%);max-height:84vh;overflow:auto;background:var(--dsw-alias-bg-layer-1);border:1px solid var(--dsw-alias-border-l2);border-radius:14px;padding:18px;display:flex;flex-direction:column;gap:11px;box-shadow:0 18px 50px rgba(0,0,0,.32)}
.mtx-sheet h3{margin:0;font-size:14px;font-weight:650}
.mtx-grow{display:flex;gap:8px;align-items:center}
.mtx-grow .mtx-input{padding:6px 9px;font-size:12.5px}
.mtx-assign{display:flex;flex-direction:column;gap:2px;border-top:1px solid var(--dsw-alias-border-l2);padding-top:10px;margin-top:2px}
.mtx-assign-row{display:flex;align-items:center;gap:10px;font-size:12px;padding:3px 0}
.mtx-assign-row>span{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.mtx-select{border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-2);color:var(--dsw-alias-label-primary);border-radius:7px;padding:4px 6px;font:inherit;font-size:11.5px;max-width:180px}
.mtx-sheet-foot{display:flex;justify-content:flex-end;gap:8px;border-top:1px solid var(--dsw-alias-border-l2);padding-top:12px}
.mtx-note{color:var(--dsw-alias-label-secondary);font-size:11px;margin:0;line-height:1.5}
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

    /**
     * The board: lists on the left, the selected list's tasks on the right.
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
      const [due, setDue] = useState('')
      const [busy, setBusy] = useState(false)
      const [toast, setToast] = useState('')
      const [groups, setGroups] = useState([])
      const [collapsed, setCollapsed] = useState({})
      const [sheet, setSheet] = useState(null)
      const [myday, setMyday] = useState(null)
      const [pane, setPane] = useState('list')
      const listIdRef = useRef(listId)
      listIdRef.current = listId

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

      const bootstrap = useCallback(async () => {
        setPhase('loading')
        setError('')
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
      }, [loadMyDay, loadTasks, loadToday])

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
          }))

      const placed = new Set()
      const sections = []
      for (const group of groups) {
        const name = typeof group?.name === 'string' ? group.name.trim() : ''
        if (name === '') continue
        const ids = Array.isArray(group.listIds) ? group.listIds : []
        const members = lists.filter(list => ids.includes(list.id))
        if (members.length === 0) continue
        for (const list of members) placed.add(list.id)
        sections.push({ key: `group:${name}`, name, lists: members })
      }
      const loose = lists.filter(list => !placed.has(list.id))

      const renderRow = (list, sub) => {
        const row = counts[list.id]
        return h('button', {
          key: list.id,
          className: `mtx-listrow${sub ? ' sub' : ''}${list.id === listId ? ' on' : ''}`,
          onClick: () => selectList(list.id),
          title: list.displayName,
        },
          h('span', { className: `mtx-listdot${row && row.overdue > 0 ? ' late' : ''}` }),
          h('span', { className: 'mtx-listname' }, list.displayName),
          h('span', { className: 'mtx-listcount' }, row ? String(row.open) : ''))
      }

      const openSheet = () => setSheet({
        groups: groups.map(group => ({ name: group.name, listIds: [...(group.listIds ?? [])] })),
      })

      return h('div', { className: 'mtx-page', style: { position: 'relative' } },
        h('header', { className: 'mtx-head' },
          h('div', { className: 'mtx-brand' },
            h('div', { className: 'mtx-logo' }, '✓'),
            h('div', { style: { minWidth: 0 } },
              h('h1', { className: 'mtx-title' }, t('title')),
              h('p', { className: 'mtx-sub' }, auth?.account?.name
                ? `${t('account')} · ${auth.account.name} · ${auth.account.username ?? ''}`
                : t('account')))),
          h('div', { className: 'mtx-stats' },
            h(Stat, { label: t('today'), value: today?.dueToday?.length, tone: 'today' }),
            h(Stat, { label: t('overdue'), value: today?.overdue?.length, tone: 'overdue' }),
            h(Stat, { label: t('starred'), value: today?.starred?.length, tone: 'starred' })),
          h('button', { className: 'mtx-btn', onClick: () => { void refreshAll() }, disabled: busy, title: t('refresh') },
            h(IconRefresh), t('refresh'))),
        h('div', { className: 'mtx-body' },
          h('aside', { className: 'mtx-aside' },
            h('div', { className: 'mtx-aside-head' }, h('span', null, t('lists')), h('span', null, String(lists.length))),
            h('button', {
              className: `mtx-listrow myday${pane === 'myday' ? ' on' : ''}`,
              onClick: () => { setPane('myday'); void loadMyDay(true) },
              title: t('myday'),
            },
              h(IconSun, { filled: pane === 'myday', size: 14 }),
              h('span', { className: 'mtx-listname' }, t('myday')),
              h('span', { className: 'mtx-listcount' }, String(mydayItems.length))),
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
                h(IconGear), t('groupSettings'))))),
          h('section', { className: 'mtx-main' },
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
                    : rows))
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

      return h('div', {
        className: 'mtx-modal',
        onMouseDown: (event) => { if (event.target === event.currentTarget) onClose() },
      },
        h('div', { className: 'mtx-sheet' },
          h('h3', null, t('groupSettings')),
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
          h('div', { className: 'mtx-sheet-foot' },
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

    function TaskRow({ task, todayKey, t, busy, onComplete, onReopen, onStar, onDelete, onMyDay, mydayOn = false, showList = false }) {
      const done = task.status === 'completed'
      const day = typeof task.dueDay === 'string' ? task.dueDay : null
      const label = dueLabel(day, todayKey)
      const late = label !== null && day !== null && day < todayKey
      return h('div', { className: `mtx-row${done ? ' done' : ''}` },
        h('button', {
          type: 'button',
          className: `mtx-check${done ? ' on' : ''}`,
          title: done ? t('reopen') : t('markDone'),
          disabled: busy,
          onClick: done ? onReopen : onComplete,
        }, h(IconCheck)),
        h('div', { className: 'mtx-row-main' },
          h('div', { className: 'mtx-row-title' }, task.title),
          (label !== null || task.importance === 'high' || (task.body?.content ?? '') !== '' || (showList && task.listName))
            ? h('div', { className: 'mtx-row-meta' },
                showList && task.listName ? h('span', { className: 'mtx-chip' }, task.listName) : null,
                label === null ? null : h('span', { className: `mtx-chip${late ? ' late' : ''}${day === todayKey ? ' today' : ''}` }, label),
                task.importance === 'high' ? h('span', { className: 'mtx-chip today' }, '★ ' + t('starred')) : null,
                (task.body?.content ?? '') === '' ? null : h('span', { className: 'mtx-chip' }, String(task.body.content).slice(0, 60)))
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
            title: '删除',
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

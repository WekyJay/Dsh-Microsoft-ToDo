/**
 * Microsoft To Do bridge, host half.
 *
 * Keeps ONE `microsoft-todo-mcp` process alive over MCP stdio (it owns the
 * device-code sign-in and the MSAL token cache) and serves the panel's JSON
 * API under `/microsoft-todo/api/*`. The browser half (`client.js`) renders
 * the board and talks to these routes document-relatively.
 */
import { spawn } from 'node:child_process'
import { readFile, readdir, writeFile, rename } from 'node:fs/promises'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

/** Absolute route prefix owned by this plugin. */
const ROUTE = '/microsoft-todo'

/**
 * Where the local list groups live.
 *
 * Microsoft Graph exposes no group/folder relationship for To Do lists, so the
 * To Do client's own grouping is invisible to the API. Groups are therefore
 * kept locally, next to this plugin, and merged into the panel's catalog.
 */
const GROUPS_FILE = fileURLToPath(new URL('./groups.json', import.meta.url))

/**
 * Where the local "My Day" picks live.
 *
 * Microsoft Graph has no My Day resource for To Do tasks (only Planner exposes
 * `myDayTasks`), so the day's short list is kept here: the task ids picked for
 * one calendar day, reset when the day changes, exactly like the To Do client.
 */
const MYDAY_FILE = fileURLToPath(new URL('./myday.json', import.meta.url))

/**
 * The zone the user's day boundaries are read in.
 *
 * A constant `Asia/Shanghai` was wrong for everyone outside UTC+8: "Today",
 * "Overdue", "My Day" and every due date this bridge writes sat a day off for
 * them, by exactly the offset between the two zones. The device's own zone is
 * the answer. `DSH_TODO_TZONE` (then `TZ`) overrides it for a machine whose
 * zone the OS reports wrongly — containers and CI usually claim UTC — and
 * `UTC` is the floor: `Intl` throws on an unknown zone at format time, which
 * would fail every read instead of degrading one field.
 */
function resolveTimeZone() {
  const candidates = [
    process.env.DSH_TODO_TZONE,
    process.env.TZ,
    Intl.DateTimeFormat().resolvedOptions().timeZone,
  ]
  for (const zone of candidates) {
    if (typeof zone !== 'string' || zone.trim() === '') continue
    try {
      new Intl.DateTimeFormat('en-US', { timeZone: zone })
      return zone.trim()
    }
    catch { /* not a zone this runtime knows: try the next candidate */ }
  }
  return 'UTC'
}

const TZONE = resolveTimeZone()

/** Graph, reached directly for the few fields the MCP tool surface omits. */
const GRAPH = 'https://graph.microsoft.com/v1.0'

/**
 * The MSAL cache the MCP server writes after each silent refresh.
 *
 * The bridge keeps a token fresh as a side effect of talking to Graph, so the
 * direct Graph calls here borrow that token instead of running a second
 * sign-in flow.
 */
const TOKEN_CACHE = join(homedir(), '.config', 'microsoft-todo-mcp', 'token-cache.json')

/**
 * The upstream server, pinned to the version this bridge is written against.
 *
 * `npx -y microsoft-todo-mcp` asks npm for `latest` on every start, so two
 * machines showing the same plugin version could be running different server
 * code, and a bad upstream release would reach every install without this
 * repository changing. Moving this pin is the deliberate act of re-testing.
 */
const MCP_PKG = 'microsoft-todo-mcp@0.2.0'

/** npm's launcher carries a `.cmd` suffix on Windows and is not runnable without it. */
const NPX = process.platform === 'win32' ? 'npx.cmd' : 'npx'

/**
 * The environment the upstream server is allowed to see.
 *
 * `env: process.env` handed a third-party child process every secret the host
 * carried in its environment — API keys, tokens, whatever the operator had
 * exported. The server needs none of that: it needs to find `node`, to reach
 * its own cache under HOME, a temp directory, and the locale and proxy
 * settings that decide whether its own network calls work.
 */
const CHILD_ENV_KEYS = [
  'PATH', 'HOME', 'USERPROFILE', 'TMPDIR', 'TMP', 'TEMP',
  'LANG', 'LC_ALL', 'LC_CTYPE', 'TZ',
  'HTTP_PROXY', 'HTTPS_PROXY', 'NO_PROXY', 'http_proxy', 'https_proxy', 'no_proxy',
  'NODE_EXTRA_CA_CERTS', 'SSL_CERT_FILE', 'SSL_CERT_DIR',
  'npm_config_cache', 'npm_config_registry', 'npm_config_proxy', 'npm_config_https_proxy',
  // Windows: without these, npm cannot resolve its own installation.
  'SystemRoot', 'COMSPEC', 'PATHEXT', 'APPDATA', 'LOCALAPPDATA', 'PROGRAMDATA',
]

/** The child's environment, built from the allow-list above. */
function childEnv() {
  const env = {}
  for (const key of CHILD_ENV_KEYS) {
    const value = process.env[key]
    if (typeof value === 'string' && value !== '') env[key] = value
  }
  return env
}

/** Upper bound on a request body, in bytes. */
const BODY_LIMIT = 64 * 1024

/** Uploads carry the file as base64, so they need a much wider bound. */
const UPLOAD_LIMIT = 34 * 1024 * 1024

/** How long `/api/today` is allowed to reuse its aggregate. */
const TODAY_TTL_MS = 30_000

/* ---------- version + compatibility ---------- */

/**
 * This plugin's manifest, read once for the version badge and the DSH range.
 *
 * `engines.dsh` is the requirement dsh itself enforces at install time; the
 * panel shows the same range so a running install can explain a mismatch
 * without the user having to read package.json.
 */
async function pluginManifest() {
  // Re-read every call: the panel and the marketplace pin the version they
  // see, and a running install may be one bundle behind the manifest on disk.
  let manifest = { name: 'dsh-microsoft-todo', version: '0.0.0', dshRange: '' }
  try {
    const raw = JSON.parse(await readFile(fileURLToPath(new URL('./package.json', import.meta.url)), 'utf8'))
    const range = raw?.engines?.dsh
    manifest = {
      name: typeof raw?.name === 'string' ? raw.name : manifest.name,
      version: typeof raw?.version === 'string' ? raw.version : manifest.version,
      dshRange: typeof range === 'string' ? range : '',
    }
  } catch { /* a packaged bundle without package.json still renders the panel */ }
  return manifest
}

/** Parsed semantic version, or null when the text is not one. */
function parseVersion(text) {
  const match = /^v?(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$/.exec(String(text).trim())
  if (match === null) return null
  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
    pre: typeof match[4] === 'string' ? match[4].split('.') : [],
  }
}

function comparePart(left, right) {
  return left === right ? 0 : (left < right ? -1 : 1)
}

/** Semver precedence, with prerelease ordering (0.2.2-rc.1 < 0.2.2). */
function compareVersions(left, right) {
  for (const key of ['major', 'minor', 'patch']) {
    const order = comparePart(left[key], right[key])
    if (order !== 0) return order
  }
  if (left.pre.length === 0 && right.pre.length === 0) return 0
  if (left.pre.length === 0) return 1
  if (right.pre.length === 0) return -1
  for (let index = 0; index < Math.max(left.pre.length, right.pre.length); index += 1) {
    const l = left.pre[index]
    const r = right.pre[index]
    if (l === undefined) return -1
    if (r === undefined) return 1
    const lNum = /^\d+$/.test(l)
    const rNum = /^\d+$/.test(r)
    if (lNum && rNum) {
      const order = comparePart(Number(l), Number(r))
      if (order !== 0) return order
    } else if (lNum) return -1
    else if (rNum) return 1
    else {
      const order = comparePart(l, r)
      if (order !== 0) return order
    }
  }
  return 0
}

/** One whitespace-separated comparator (`>=0.2.0`, `^0.2.0`, `~0.2.0`, `0.2.0`). */
function satisfiesComparator(version, token) {
  const match = /^(>=|<=|>|<|=|\^|~)?\s*(.+)$/.exec(token)
  if (match === null) return false
  const operator = match[1] ?? '='
  const bound = parseVersion(match[2])
  if (bound === null) return false
  const order = compareVersions(version, bound)
  if (operator === '>=') return order >= 0
  if (operator === '<=') return order <= 0
  if (operator === '>') return order > 0
  if (operator === '<') return order < 0
  if (operator === '=') return order === 0
  if (order < 0) return false
  const upper = operator === '~' || bound.major === 0
    ? { major: bound.major, minor: bound.minor + 1, patch: 0, pre: [] }
    : { major: bound.major + 1, minor: 0, patch: 0, pre: [] }
  return compareVersions(version, upper) < 0
}

/** Whether `version` satisfies a plain AND range such as `>=0.2.0 <0.5.0`. */
function satisfiesRange(versionText, range) {
  const version = parseVersion(versionText)
  if (version === null) return null
  const tokens = String(range).trim().split(/\s+/).filter(Boolean)
  if (tokens.length === 0) return null
  return tokens.every(token => satisfiesComparator(version, token))
}

/**
 * The running dsh version.
 *
 * A plugin bundle cannot import dsh's own packages, so this reads the same
 * package dsh reads for its compatibility check (`dsh-app-boot`), from this
 * machine's profiles. Returns null when the layout is something else (a
 * packaged binary, a container) — the panel then says the version is unknown
 * instead of guessing.
 */
async function runningDshVersion() {
  const fromEnv = process.env.DSH_VERSION
  if (typeof fromEnv === 'string' && parseVersion(fromEnv) !== null) return fromEnv.trim()
  const homes = []
  if (typeof process.env.DSH_HOME === 'string' && process.env.DSH_HOME !== '') homes.push(process.env.DSH_HOME)
  homes.push(join(homedir(), '.dsh'))
  for (const home of homes) {
    let profiles = []
    try { profiles = await readdir(join(home, 'profiles')) } catch { continue }
    for (const profile of profiles) {
      try {
        const raw = JSON.parse(await readFile(join(home, 'profiles', profile, 'node_modules', '@deepseek-ai', 'dsh-app-boot', 'package.json'), 'utf8'))
        if (typeof raw?.version === 'string' && parseVersion(raw.version) !== null) return raw.version
      } catch { /* profile without dsh packages; keep looking */ }
    }
  }
  return null
}

/** Version + compatibility payload used by the panel header. */
async function metaPayload() {
  const manifest = await pluginManifest()
  const version = await runningDshVersion()
  const compatible = manifest.dshRange === '' || version === null
    ? null
    : satisfiesRange(version, manifest.dshRange)
  return {
    plugin: { name: manifest.name, version: manifest.version },
    dsh: { version, required: manifest.dshRange || null, compatible },
  }
}


/** Calendar day (YYYY-MM-DD) of an instant, in the user's zone. */
function dayOf(date) {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: TZONE, year: 'numeric', month: '2-digit', day: '2-digit',
  }).format(date)
}

/** Offset of one IANA zone from UTC at an instant, in milliseconds. */
function zoneOffset(zone, instantMs) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: zone, hourCycle: 'h23',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(new Date(instantMs))
  const at = {}
  for (const part of parts) at[part.type] = part.value
  const wall = Date.UTC(Number(at.year), Number(at.month) - 1, Number(at.day), Number(at.hour), Number(at.minute), Number(at.second))
  return wall - instantMs
}

/** The instant described by a wall clock written in one IANA zone. */
function instantOf(wall, zone) {
  const guess = Date.parse(`${wall.slice(0, 19).replace(' ', 'T')}Z`)
  if (Number.isNaN(guess)) return null
  return guess - zoneOffset(zone, guess)
}

/** Calendar day carried by a Graph task due date, or null when it has none. */
function dueDay(task) {
  const raw = task?.dueDateTime
  if (raw === undefined || raw === null) return null
  const value = typeof raw === 'string' ? raw : raw.dateTime
  if (typeof value !== 'string' || value.length < 10) return null
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value
  const zone = typeof raw === 'object' && typeof raw.timeZone === 'string' ? raw.timeZone : TZONE
  const instant = instantOf(value, zone)
  return instant === null ? value.slice(0, 10) : dayOf(new Date(instant))
}

/** The text payload of a tool result, parsed when it is JSON. */
function unwrap(result) {
  const content = result?.content
  const text = Array.isArray(content) ? content.find(part => part?.type === 'text')?.text : undefined
  if (typeof text !== 'string') return result
  try {
    return JSON.parse(text)
  } catch {
    return text
  }
}

/**
 * One long-lived MCP stdio session with the upstream Microsoft To Do server.
 * Newline-delimited JSON-RPC as the MCP stdio transport prescribes.
 */
class TodoBridge {
  constructor() {
    this.child = null
    this.buffer = ''
    this.stderrTail = ''
    this.nextId = 1
    this.initId = null
    this.initialized = false
    this.pending = new Map()
    this.readyWaiters = []
  }

  /** Spawn the upstream server and send `initialize`. */
  start() {
    if (this.child !== null) return
    const child = spawn(NPX, ['-y', MCP_PKG], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: childEnv(),
    })
    this.child = child
    this.buffer = ''
    this.initialized = false
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk) => { this.ingest(chunk) })
    child.stderr.setEncoding('utf8')
    child.stderr.on('data', (chunk) => { this.stderrTail = (this.stderrTail + chunk).slice(-4000) })
    child.on('error', (error) => { this.fail(new Error(`microsoft-todo-mcp cannot start: ${error.message}`)) })
    child.on('exit', (code) => { this.fail(new Error(`microsoft-todo-mcp exited (${String(code)})`)) })
    child.stdin.on('error', () => {})
    this.initId = this.nextId++
    this.send({
      jsonrpc: '2.0',
      id: this.initId,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'dsh-microsoft-todo', version: '0.1.0' },
      },
    })
  }

  /** Settle every waiter and every in-flight call after the child is gone. */
  fail(error) {
    const detail = this.stderrTail.trim()
    const wrapped = new Error(detail === '' ? error.message : `${error.message}: ${detail.slice(-600)}`)
    this.child = null
    this.initialized = false
    this.initId = null
    const waiters = this.readyWaiters
    this.readyWaiters = []
    for (const waiter of waiters) waiter.reject(wrapped)
    for (const entry of this.pending.values()) {
      clearTimeout(entry.timer)
      entry.reject(wrapped)
    }
    this.pending.clear()
  }

  /** Write one JSON-RPC frame. */
  send(frame) {
    const child = this.child
    if (child === null) return
    try {
      child.stdin.write(`${JSON.stringify(frame)}\n`)
    } catch {
      // The exit handler reports the real failure.
    }
  }

  /** Consume stdout, splitting newline-delimited frames. */
  ingest(chunk) {
    this.buffer += chunk
    let index = this.buffer.indexOf('\n')
    while (index >= 0) {
      const line = this.buffer.slice(0, index).trim()
      this.buffer = this.buffer.slice(index + 1)
      if (line !== '') this.dispatch(line)
      index = this.buffer.indexOf('\n')
    }
  }

  /** Route one incoming frame. */
  dispatch(line) {
    let message
    try {
      message = JSON.parse(line)
    } catch {
      return
    }
    if (message.id === this.initId) {
      if (message.error !== undefined) {
        this.fail(new Error(`initialize failed: ${message.error?.message ?? 'unknown error'}`))
        return
      }
      this.initialized = true
      this.initId = null
      this.send({ jsonrpc: '2.0', method: 'notifications/initialized' })
      const waiters = this.readyWaiters
      this.readyWaiters = []
      for (const waiter of waiters) waiter.resolve()
      return
    }
    const entry = this.pending.get(message.id)
    if (entry === undefined) return
    this.pending.delete(message.id)
    clearTimeout(entry.timer)
    if (message.error !== undefined) entry.reject(new Error(message.error?.message ?? 'tool call failed'))
    else entry.resolve(message.result)
  }

  /** Resolve once the session is initialized, spawning it when needed. */
  ready() {
    if (this.initialized) return Promise.resolve()
    this.start()
    return new Promise((resolve, reject) => { this.readyWaiters.push({ resolve, reject }) })
  }

  /** Call one MCP tool and return its parsed result. */
  async call(name, args, timeoutMs = 60_000) {
    await this.ready()
    const id = this.nextId++
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`${name} timed out after ${Math.round(timeoutMs / 1000)}s`))
      }, timeoutMs)
      timer.unref?.()
      this.pending.set(id, {
        timer,
        resolve: (result) => { unwrapResultInto(resolve, result) },
        reject,
      })
      this.send({ jsonrpc: '2.0', id, method: 'tools/call', params: { name, arguments: args } })
    })
  }

  /** Stop the child process. */
  dispose() {
    const child = this.child
    this.child = null
    if (child !== null) {
      try { child.stdin.end() } catch {}
      try { child.kill() } catch {}
    }
    this.fail(new Error('Microsoft To Do bridge closed'))
  }
}

/** Resolve a tool result through {@link unwrap}. */
function unwrapResultInto(resolve, result) {
  resolve(unwrap(result))
}

/**
 * Read the list catalog, plus the best account evidence available.
 *
 * A cached access token may look expired to `auth_status` while the server
 * silently refreshes it on the next Graph call, so the catalog is the source
 * of truth for whether the connection works.
 */
async function readState(bridge) {
  let auth = null
  try { auth = unwrap(await bridge.call('auth_status', {}, 20_000)) } catch {}
  try {
    const page = unwrap(await bridge.call('list_task_lists', { maxResults: 100 }, 30_000))
    const lists = Array.isArray(page?.items) ? page.items : []
    return { auth: { ...(auth ?? {}), state: 'authenticated' }, lists, ...(await readGroups()) }
  } catch (error) {
    const state = auth?.state === undefined || auth?.state === 'token_expired' ? 'unauthenticated' : String(auth.state)
    return { auth: { ...(auth ?? {}), state, detail: failureOf(error) }, lists: [], groups: [] }
  }
}

/** Read one list's tasks, each tagged with its due day in the user's zone. */
async function readTasks(bridge, listId, status) {
  if (listId === '') throw new Error('listId is required')
  const page = unwrap(await bridge.call('list_tasks', { listId, status, maxResults: 200 }, 45_000))
  const items = Array.isArray(page?.items) ? page.items : []
  return { items: items.map(task => ({ ...task, dueDay: dueDay(task) })) }
}

let todayCache = { at: 0, value: null }

/** Aggregate every list's open work into the day's picture. */
async function readToday(bridge, force = false) {
  const now = Date.now()
  if (!force && todayCache.value !== null && now - todayCache.at < TODAY_TTL_MS) return todayCache.value
  const { lists } = await readState(bridge)
  const today = dayOf(new Date())
  const perList = []
  const overdue = []
  const dueToday = []
  const starred = []
  const undated = []
  const allOpen = []
  const settled = await Promise.all(lists.map(async (list) => {
    try {
      const page = unwrap(await bridge.call('list_tasks', {
        listId: list.id, status: 'notStarted', maxResults: 200,
      }, 45_000))
      return { list, tasks: Array.isArray(page?.items) ? page.items : [] }
    } catch {
      return { list, tasks: [] }
    }
  }))
  for (const { list, tasks } of settled) {
    let late = 0
    let due = 0
    for (const task of tasks) {
      const entry = {
        listId: list.id,
        listName: list.displayName,
        id: task.id,
        title: task.title,
        dueDateTime: task.dueDateTime ?? null,
        importance: task.importance ?? 'normal',
        createdDateTime: task.createdDateTime ?? null,
      }
      const day = dueDay(task)
      const open = { ...entry, day }
      allOpen.push(open)
      if (day === null) undated.push(entry)
      else if (day < today) { late += 1; overdue.push(open) }
      else if (day === today) { due += 1; dueToday.push(open) }
      if (entry.importance === 'high') starred.push(open)
    }
    perList.push({
      id: list.id,
      name: list.displayName,
      open: tasks.length,
      overdue: late,
      dueToday: due,
    })
  }
  const value = { today, perList, overdue, dueToday, starred, undated, allOpen, fetchedAt: new Date().toISOString() }
  todayCache = { at: now, value }
  return value
}

/** Escape one Graph path segment. */
function seg(value) {
  return encodeURIComponent(String(value))
}

/** The freshest access token in the MCP server's MSAL cache. */
async function cachedToken() {
  let cache
  try {
    cache = JSON.parse(await readFile(TOKEN_CACHE, 'utf8'))
  } catch {
    throw new Error('未找到微软登录令牌，请先在面板里完成一次设备码登录')
  }
  let best = null
  for (const entry of Object.values(cache?.AccessToken ?? {})) {
    if (entry === null || typeof entry !== 'object' || typeof entry.secret !== 'string' || entry.secret === '') continue
    const expires = Number(entry.expires_on ?? 0)
    if (best === null || expires > best.expires) best = { secret: entry.secret, expires }
  }
  if (best === null) throw new Error('未找到微软登录令牌，请先在面板里完成一次设备码登录')
  return best
}

/** Ask the MCP server to refresh the cached token, then read it again. */
async function refreshToken(bridge) {
  try {
    await bridge.call('list_task_lists', { maxResults: 1 }, 30_000)
  } catch {}
  return cachedToken()
}

/** One authenticated Graph call, refreshing the token once on a 401. */
async function graph(bridge, path, init = {}) {
  const headers = { 'content-type': 'application/json', ...(init.headers ?? {}) }
  const attempt = async (token) => fetch(GRAPH + path, {
    ...init,
    headers: { ...headers, authorization: `Bearer ${token}` },
  })
  let token = await cachedToken()
  if (token.expires * 1000 - Date.now() < 90_000) token = await refreshToken(bridge)
  let res = await attempt(token.secret)
  if (res.status === 401) {
    token = await refreshToken(bridge)
    res = await attempt(token.secret)
  }
  const text = await res.text()
  let body = null
  if (text !== '') {
    try { body = JSON.parse(text) } catch { body = text }
  }
  if (!res.ok) {
    const message = body?.error?.message ?? (typeof body === 'string' ? body : res.statusText)
    throw new Error(`Graph ${res.status}: ${String(message).slice(0, 400)}`)
  }
  return body
}

/** Graph's task body only accepts HTML, so notes round-trip through this. */
function toHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>')
}

/** HTML task body back to plain text for the notes editor. */
function fromHtml(html) {
  return String(html ?? '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .trim()
}

/** One task's full detail: notes, dates, recurrence, steps and files. */
async function readDetail(bridge, listId, taskId) {
  const task = await graph(bridge, `/me/todo/lists/${seg(listId)}/tasks/${seg(taskId)}?$expand=checklistItems`)
  let attachments = []
  try {
    const page = await graph(bridge, `/me/todo/lists/${seg(listId)}/tasks/${seg(taskId)}/attachments?$select=id,name,size,contentType`)
    attachments = Array.isArray(page?.value) ? page.value : []
  } catch {}
  return {
    listId,
    task,
    notes: fromHtml(task?.body?.content),
    checklist: Array.isArray(task?.checklistItems) ? task.checklistItems : [],
    attachments,
  }
}

/** Map the detail sheet's recurrence draft onto the Graph recurrence object. */
function recurrenceOf(input) {
  const pattern = input?.pattern ?? {}
  const type = ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly'].includes(pattern.type) ? pattern.type : 'weekly'
  const interval = Math.max(1, Math.min(99, Number(pattern.interval ?? 1) || 1))
  const out = { type, interval }
  if (type === 'weekly' || type === 'relativeMonthly') {
    const days = Array.isArray(pattern.daysOfWeek) && pattern.daysOfWeek.length > 0
      ? pattern.daysOfWeek.filter(day => typeof day === 'string')
      : ['monday']
    out.daysOfWeek = days
    if (type === 'weekly') out.firstDayOfWeek = 'sunday'
    if (type === 'relativeMonthly') {
      out.index = ['first', 'second', 'third', 'fourth', 'last'].includes(pattern.index) ? pattern.index : 'first'
    }
  }
  if (type === 'absoluteMonthly') out.dayOfMonth = Math.max(1, Math.min(31, Number(pattern.dayOfMonth ?? 1) || 1))
  const range = input?.range ?? {}
  const kind = ['endDate', 'noEnd', 'numbered'].includes(range.type) ? range.type : 'noEnd'
  const startDate = typeof range.startDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(range.startDate)
    ? range.startDate
    : dayOf(new Date())
  const result = { pattern: out, range: { type: kind, startDate, recurrenceTimeZone: TZONE } }
  if (kind === 'endDate') {
    result.range.endDate = typeof range.endDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(range.endDate) ? range.endDate : startDate
  }
  if (kind === 'numbered') {
    result.range.numberOfOccurrences = Math.max(1, Math.min(999, Number(range.numberOfOccurrences ?? 10) || 10))
  }
  return result
}

/** Apply one detail-sheet edit to a task, in one Graph PATCH. */
async function saveDetail(bridge, body) {
  const listId = String(body.listId ?? '')
  const taskId = String(body.taskId ?? '')
  if (listId === '' || taskId === '') throw new Error('listId and taskId are required')
  const patch = {}
  if (typeof body.title === 'string' && body.title.trim() !== '') patch.title = body.title.trim()
  if (typeof body.bodyText === 'string') patch.body = { content: toHtml(body.bodyText), contentType: 'html' }
  if (Object.hasOwn(body, 'dueDate')) {
    patch.dueDateTime = typeof body.dueDate === 'string' && body.dueDate !== ''
      ? { dateTime: `${body.dueDate}T17:00:00.0000000`, timeZone: TZONE }
      : null
  }
  if (Object.hasOwn(body, 'reminderAt')) {
    const at = typeof body.reminderAt === 'string' ? body.reminderAt : ''
    patch.isReminderOn = at !== ''
    patch.reminderDateTime = at === ''
      ? null
      : { dateTime: at.length === 16 ? `${at}:00.0000000` : at, timeZone: TZONE }
  }
  if (typeof body.importance === 'string' && ['low', 'normal', 'high'].includes(body.importance)) {
    patch.importance = body.importance
  }
  if (Object.hasOwn(body, 'recurrence')) {
    patch.recurrence = body.recurrence === null ? null : recurrenceOf(body.recurrence)
  }
  if (Object.keys(patch).length === 0) return null
  const updated = await graph(bridge, `/me/todo/lists/${seg(listId)}/tasks/${seg(taskId)}`, {
    method: 'PATCH',
    body: JSON.stringify(patch),
  })
  todayCache = { at: 0, value: null }
  return updated
}

/** Steps (checklist items) for one task, through the MCP server. */
async function stepAction(bridge, body) {
  const action = String(body.action ?? '')
  const listId = String(body.listId ?? '')
  const taskId = String(body.taskId ?? '')
  if (listId === '' || taskId === '') throw new Error('listId and taskId are required')
  if (action === 'add') {
    const displayName = String(body.displayName ?? '').trim()
    if (displayName === '') throw new Error('displayName is required')
    return { item: unwrap(await bridge.call('add_checklist_item', { listId, taskId, displayName }, 30_000)) }
  }
  if (action === 'update') {
    const args = { listId, taskId, itemId: String(body.itemId ?? '') }
    if (typeof body.displayName === 'string') args.displayName = body.displayName
    if (typeof body.isChecked === 'boolean') args.isChecked = body.isChecked
    return { item: unwrap(await bridge.call('update_checklist_item', args, 30_000)) }
  }
  if (action === 'delete') {
    await bridge.call('delete_checklist_item', { listId, taskId, itemId: String(body.itemId ?? '') }, 30_000)
    return { ok: true }
  }
  if (action === 'promote') {
    const itemId = String(body.itemId ?? '')
    const page = unwrap(await bridge.call('list_checklist_items', { listId, taskId }, 30_000))
    const items = Array.isArray(page?.items) ? page.items : []
    const item = items.find(entry => entry.id === itemId)
    if (item === undefined) throw new Error('checklist item not found')
    const created = unwrap(await bridge.call('create_task', { listId, title: String(item.displayName) }, 45_000))
    await bridge.call('delete_checklist_item', { listId, taskId, itemId }, 30_000)
    todayCache = { at: 0, value: null }
    return { task: created }
  }
  throw new Error(`unknown checklist action: ${action}`)
}

/** Add or remove the files attached to one task. */
async function attachmentAction(bridge, body) {
  const action = String(body.action ?? '')
  const listId = String(body.listId ?? '')
  const taskId = String(body.taskId ?? '')
  if (listId === '' || taskId === '') throw new Error('listId and taskId are required')
  const base = `/me/todo/lists/${seg(listId)}/tasks/${seg(taskId)}/attachments`
  if (action === 'upload') {
    const data = String(body.dataBase64 ?? '')
    if (data === '') throw new Error('dataBase64 is required')
    const bytes = Buffer.from(data, 'base64')
    const name = String(body.name ?? 'attachment')
    const contentType = String(body.contentType ?? 'application/octet-stream')
    if (bytes.length <= 3 * 1024 * 1024) {
      const attachment = await graph(bridge, base, {
        method: 'POST',
        body: JSON.stringify({
          '@odata.type': '#microsoft.graph.taskFileAttachment',
          name,
          contentBytes: data,
          contentType,
        }),
      })
      return {
        attachment: {
          id: attachment?.id,
          name: attachment?.name,
          size: attachment?.size,
          contentType: attachment?.contentType,
        },
      }
    }
    if (bytes.length > 25 * 1024 * 1024) throw new Error('file is larger than 25 MB')
    // Files above the 3 MB inline limit go through a pre-authenticated
    // upload session: 4 MB chunks, no Authorization header on the PUTs.
    const session = await graph(bridge, `${base}/createUploadSession`, {
      method: 'POST',
      body: JSON.stringify({ AttachmentItem: { attachmentType: 'file', name, size: bytes.length } }),
    })
    const uploadUrl = String(session?.uploadUrl ?? '')
    if (uploadUrl === '') throw new Error('Graph did not return an upload session')
    const chunk = 4 * 1024 * 1024
    for (let offset = 0; offset < bytes.length; offset += chunk) {
      const end = Math.min(offset + chunk, bytes.length)
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'content-type': contentType,
          'content-range': `bytes ${offset}-${end - 1}/${bytes.length}`,
        },
        body: bytes.subarray(offset, end),
      })
      if (!res.ok) {
        const text = await res.text()
        throw new Error(`upload session failed (${res.status}): ${text.slice(0, 300)}`)
      }
    }
    return { attachment: { name, size: bytes.length, contentType } }
  }
  if (action === 'delete') {
    const attachmentId = String(body.attachmentId ?? '')
    if (attachmentId === '') throw new Error('attachmentId is required')
    await graph(bridge, `${base}/${seg(attachmentId)}`, { method: 'DELETE' })
    return { ok: true }
  }
  throw new Error(`unknown attachment action: ${action}`)
}

/** Read one attachment back as base64 so the browser can save it. */
async function downloadAttachment(bridge, listId, taskId, attachmentId) {
  const one = await graph(bridge, `/me/todo/lists/${seg(listId)}/tasks/${seg(taskId)}/attachments/${seg(attachmentId)}`)
  return {
    name: one?.name ?? 'attachment',
    contentType: one?.contentType ?? 'application/octet-stream',
    dataBase64: one?.contentBytes ?? '',
  }
}

/** Read a JSON request body. */
async function readBody(req, limit = BODY_LIMIT) {
  const chunks = []
  let size = 0
  for await (const chunk of req) {
    size += chunk.length
    if (size > limit) throw new Error('request body too large')
    chunks.push(chunk)
  }
  const text = Buffer.concat(chunks).toString('utf8').trim()
  if (text === '') return {}
  const parsed = JSON.parse(text)
  if (parsed === null || typeof parsed !== 'object') throw new Error('request body must be a JSON object')
  return parsed
}

/** Write one JSON response. */
function sendJson(res, status, payload) {
  const body = JSON.stringify(payload)
  res.statusCode = status
  res.setHeader('content-type', 'application/json; charset=utf-8')
  res.setHeader('cache-control', 'no-store')
  res.end(body)
}

/** Strip Graph/MCP noise so a failure reads as one sentence. */
function failureOf(error) {
  const message = error instanceof Error ? error.message : String(error)
  return message.length > 600 ? `${message.slice(0, 600)}…` : message
}

/** Coerce one stored group into shape, or null when it is unusable. */
function asGroup(raw) {
  if (raw === null || typeof raw !== 'object') return null
  const name = typeof raw.name === 'string' ? raw.name.trim() : ''
  if (name === '') return null
  const listIds = Array.isArray(raw.listIds) ? raw.listIds.filter(id => typeof id === 'string' && id !== '') : []
  return { name, listIds }
}

/** Read the locally stored list groups; an absent file is an empty catalog. */
async function readGroups() {
  try {
    const parsed = JSON.parse(await readFile(GROUPS_FILE, 'utf8'))
    const raw = Array.isArray(parsed?.groups) ? parsed.groups : []
    const groups = []
    const seen = new Set()
    for (const entry of raw) {
      const group = asGroup(entry)
      if (group === null || seen.has(group.name)) continue
      seen.add(group.name)
      groups.push(group)
    }
    return { groups }
  } catch (error) {
    if (error?.code === 'ENOENT') return { groups: [] }
    throw error
  }
}

/** Replace the stored groups atomically. */
async function writeGroups(groups) {
  const temp = `${GROUPS_FILE}.tmp`
  await writeFile(temp, `${JSON.stringify({ groups }, null, 2)}\n`, 'utf8')
  await rename(temp, GROUPS_FILE)
}

/** Coerce stored My Day picks into shape; a different day starts empty. */
function asMyDay(raw, today) {
  const date = typeof raw?.date === 'string' ? raw.date : ''
  const taskIds = Array.isArray(raw?.taskIds)
    ? raw.taskIds.filter(id => typeof id === 'string' && id !== '')
    : []
  if (date !== today) return { date: today, taskIds: [] }
  return { date: today, taskIds: [...new Set(taskIds)] }
}

/** Read today's picks; a missing or unreadable file means nothing picked yet. */
async function readMyDay(today) {
  try {
    return asMyDay(JSON.parse(await readFile(MYDAY_FILE, 'utf8')), today)
  } catch (error) {
    if (error?.code === 'ENOENT' || error instanceof SyntaxError) return { date: today, taskIds: [] }
    throw error
  }
}

/** Replace the stored picks atomically. */
async function writeMyDay(value) {
  const temp = `${MYDAY_FILE}.tmp`
  await writeFile(temp, `${JSON.stringify(value, null, 2)}\n`, 'utf8')
  await rename(temp, MYDAY_FILE)
}

/**
 * Today's picks resolved to the tasks they point at.
 * Picks of tasks that were completed or deleted elsewhere simply drop out.
 */
async function readMyDayView(bridge, force) {
  const today = dayOf(new Date())
  const store = await readMyDay(today)
  const data = await readToday(bridge, force)
  const known = new Map((data.allOpen ?? []).map(entry => [entry.id, entry]))
  const items = store.taskIds.map(id => known.get(id)).filter(Boolean)
  return { date: today, taskIds: store.taskIds, items }
}

/** Serve one API request. */
async function serve(bridge, req, res) {
  const url = new URL(String(req.url), 'http://localhost')
  const path = url.pathname.slice(ROUTE.length).replace(/\/$/, '') || '/'
  try {
    if (path === '/api/meta' && req.method === 'GET') {
      sendJson(res, 200, { ok: true, ...(await metaPayload()) })
      return
    }
    if (path === '/api/state' && req.method === 'GET') {
      const state = await readState(bridge)
      sendJson(res, 200, { ok: true, ...state })
      return
    }
    if (path === '/api/tasks' && req.method === 'GET') {
      const tasks = await readTasks(bridge, url.searchParams.get('listId') ?? '', url.searchParams.get('status') ?? 'all')
      sendJson(res, 200, { ok: true, ...tasks })
      return
    }
    if (path === '/api/today' && req.method === 'GET') {
      const today = await readToday(bridge, url.searchParams.get('refresh') === '1')
      sendJson(res, 200, { ok: true, ...today })
      return
    }
    if (path === '/api/groups' && req.method === 'GET') {
      sendJson(res, 200, { ok: true, ...(await readGroups()) })
      return
    }
    if (path === '/api/myday' && req.method === 'GET') {
      const view = await readMyDayView(bridge, url.searchParams.get('refresh') === '1')
      sendJson(res, 200, { ok: true, ...view })
      return
    }
    if (path === '/api/myday' && req.method === 'POST') {
      const body = await readBody(req)
      const taskId = typeof body.taskId === 'string' ? body.taskId.trim() : ''
      if (taskId === '') throw new Error('taskId is required')
      const on = body.on !== false
      const today = dayOf(new Date())
      const store = await readMyDay(today)
      const taskIds = on
        ? (store.taskIds.includes(taskId) ? store.taskIds : [...store.taskIds, taskId])
        : store.taskIds.filter(id => id !== taskId)
      await writeMyDay({ date: today, taskIds })
      todayCache = { at: 0, value: null }
      const view = await readMyDayView(bridge, true)
      sendJson(res, 200, { ok: true, ...view })
      return
    }
    if (path === '/api/groups' && req.method === 'POST') {
      const body = await readBody(req)
      if (!Array.isArray(body.groups)) throw new Error('groups must be an array')
      if (body.groups.length > 40) throw new Error('too many groups')
      const groups = []
      const seen = new Set()
      for (const entry of body.groups) {
        const group = asGroup(entry)
        if (group === null) continue
        if (seen.has(group.name)) continue
        seen.add(group.name)
        groups.push(group)
      }
      await writeGroups(groups)
      sendJson(res, 200, { ok: true, groups })
      return
    }
    if (path === '/api/detail' && req.method === 'GET') {
      const listId = url.searchParams.get('listId') ?? ''
      const taskId = url.searchParams.get('taskId') ?? ''
      if (listId === '' || taskId === '') throw new Error('listId and taskId are required')
      sendJson(res, 200, { ok: true, ...(await readDetail(bridge, listId, taskId)) })
      return
    }
    if (path === '/api/save' && req.method === 'POST') {
      const body = await readBody(req)
      const updated = await saveDetail(bridge, body)
      const detail = await readDetail(bridge, String(body.listId), String(body.taskId))
      sendJson(res, 200, { ok: true, updated, ...detail })
      return
    }
    if (path === '/api/steps' && req.method === 'POST') {
      const body = await readBody(req)
      const result = await stepAction(bridge, body)
      sendJson(res, 200, { ok: true, ...result })
      return
    }
    if (path === '/api/attachment' && req.method === 'POST') {
      const body = await readBody(req, UPLOAD_LIMIT)
      const result = await attachmentAction(bridge, body)
      sendJson(res, 200, { ok: true, ...result })
      return
    }
    if (path === '/api/attachment' && req.method === 'GET') {
      const listId = url.searchParams.get('listId') ?? ''
      const taskId = url.searchParams.get('taskId') ?? ''
      const attachmentId = url.searchParams.get('attachmentId') ?? ''
      if (listId === '' || taskId === '' || attachmentId === '') throw new Error('listId, taskId and attachmentId are required')
      sendJson(res, 200, { ok: true, ...(await downloadAttachment(bridge, listId, taskId, attachmentId)) })
      return
    }
    if (path === '/api/sign-in' && req.method === 'POST') {
      const result = unwrap(await bridge.call('sign_in', {}, 60_000))
      sendJson(res, 200, { ok: true, signIn: result })
      return
    }
    if (path === '/api/complete' && req.method === 'POST') {
      const body = await readBody(req)
      await bridge.call('complete_task', { listId: String(body.listId ?? ''), taskId: String(body.taskId ?? '') })
      todayCache = { at: 0, value: null }
      sendJson(res, 200, { ok: true })
      return
    }
    if (path === '/api/create' && req.method === 'POST') {
      const body = await readBody(req)
      const args = { listId: String(body.listId ?? ''), title: String(body.title ?? '').trim() }
      if (args.title === '') throw new Error('title is required')
      if (typeof body.dueDate === 'string' && body.dueDate !== '') args.dueDateTime = { dateTime: `${body.dueDate}T17:00:00`, timeZone: TZONE }
      if (body.importance === 'high' || body.importance === 'low') args.importance = body.importance
      const created = unwrap(await bridge.call('create_task', args, 45_000))
      todayCache = { at: 0, value: null }
      sendJson(res, 200, { ok: true, task: created })
      return
    }
    if (path === '/api/update' && req.method === 'POST') {
      const body = await readBody(req)
      const args = { listId: String(body.listId ?? ''), taskId: String(body.taskId ?? '') }
      if (typeof body.title === 'string' && body.title.trim() !== '') args.title = body.title.trim()
      if (body.status === 'notStarted' || body.status === 'inProgress' || body.status === 'completed') args.status = body.status
      if (body.importance === 'low' || body.importance === 'normal' || body.importance === 'high') args.importance = body.importance
      if (typeof body.dueDate === 'string') {
        args.dueDateTime = body.dueDate === '' ? null : { dateTime: `${body.dueDate}T17:00:00`, timeZone: TZONE }
      }
      await bridge.call('update_task', args, 45_000)
      todayCache = { at: 0, value: null }
      sendJson(res, 200, { ok: true })
      return
    }
    if (path === '/api/delete' && req.method === 'POST') {
      const body = await readBody(req)
      await bridge.call('delete_task', { listId: String(body.listId ?? ''), taskId: String(body.taskId ?? '') })
      todayCache = { at: 0, value: null }
      sendJson(res, 200, { ok: true })
      return
    }
    sendJson(res, 404, { ok: false, error: `no route for ${req.method} ${path}` })
  } catch (error) {
    sendJson(res, 500, { ok: false, error: failureOf(error) })
  }
}

export const inject = ['webServer']

/**
 * Serve the Microsoft To Do API while this plugin is mounted.
 * @param ctx - host plugin context.
 */
export function apply(ctx) {
  const bridge = new TodoBridge()
  ctx.effect(() => () => { bridge.dispose() }, 'microsoft-todo: upstream session')
  ctx.effect(() => ctx.webServer.register({
    kind: 'prefix',
    path: ROUTE,
    handler: (req, res) => { void serve(bridge, req, res) },
  }), `microsoft-todo: ${ROUTE}/api/*`)
}

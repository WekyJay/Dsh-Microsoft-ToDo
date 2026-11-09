/**
 * Microsoft To Do bridge, host half.
 *
 * Keeps ONE `microsoft-todo-mcp` process alive over MCP stdio (it owns the
 * device-code sign-in and the MSAL token cache) and serves the panel's JSON
 * API under `/microsoft-todo/api/*`. The browser half (`client.js`) renders
 * the board and talks to these routes document-relatively.
 */
import { spawn } from 'node:child_process'
import { readFile, writeFile, rename } from 'node:fs/promises'
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

/** The zone the user's day boundaries are read in. */
const TZONE = 'Asia/Shanghai'

/** Upper bound on a request body, in bytes. */
const BODY_LIMIT = 64 * 1024

/** How long `/api/today` is allowed to reuse its aggregate. */
const TODAY_TTL_MS = 30_000

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
    const child = spawn('npx', ['-y', 'microsoft-todo-mcp'], {
      stdio: ['pipe', 'pipe', 'pipe'],
      env: process.env,
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

/** Read a JSON request body. */
async function readBody(req) {
  const chunks = []
  let size = 0
  for await (const chunk of req) {
    size += chunk.length
    if (size > BODY_LIMIT) throw new Error('request body too large')
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

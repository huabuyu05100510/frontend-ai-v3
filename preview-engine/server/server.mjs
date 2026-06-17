// ============================================================================
// server — 零依赖协同服务端（http upgrade → WebSocket → CRDT 中转）
//   端口 PORT（默认 8787）。协议：
//     client→server: {t:'join',room} / {t:'op',update} / {t:'awareness',...}
//     server→client: {t:'snapshot',snapshot} / {t:'op',update} / {t:'awareness',...}
//   仅用 Node 内置模块（http/crypto），不依赖任何 npm 包。
// ============================================================================

import http from 'node:http'
import { acceptKey, encodeTextFrame, decodeFrames } from './wsFrame.mjs'
import { createRoom, applyUpdate, snapshot } from './room.mjs'
import { convertLegacy } from './convert.mjs'

const PORT = Number(process.env.PORT) || 8787

/** Room ID 合法字符：字母/数字/连字符/下划线，长度 1-64 */
const ROOM_ID_RE = /^[a-zA-Z0-9_:.-]{1,64}$/
/** 最多同时存在的 room 数量，防止资源耗尽 */
const MAX_ROOMS = 1000
/** 单条 WebSocket 消息最大字节数（1 MB），防止 OOM 攻击 */
const MAX_MSG_BYTES = 1 * 1024 * 1024

function isValidRoomId(id) {
  return typeof id === 'string' && ROOM_ID_RE.test(id)
}

/** roomId -> { room, clients:Set<socket> } */
const rooms = new Map()
function getRoom(id) {
  let r = rooms.get(id)
  if (!r) {
    r = { room: createRoom(), clients: new Set() }
    rooms.set(id, r)
  }
  return r
}

const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Headers', 'content-type')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  if (req.method === 'OPTIONS') {
    res.writeHead(204)
    res.end()
    return
  }
  const url = new URL(req.url, 'http://localhost')
  if (req.method === 'POST' && url.pathname === '/convert') {
    const ext = (url.searchParams.get('ext') || '').toLowerCase()
    const chunks = []
    let total = 0
    req.on('data', (c) => {
      total += c.length
      if (total > 600 * 1024 * 1024) req.destroy() // 600MB 上限
      else chunks.push(c)
    })
    req.on('end', () => {
      try {
        const result = convertLegacy(Buffer.concat(chunks), ext)
        res.writeHead(result.ok ? 200 : 422, { 'content-type': 'application/json; charset=utf-8' })
        res.end(JSON.stringify(result))
      } catch (e) {
        res.writeHead(500, { 'content-type': 'application/json; charset=utf-8' })
        res.end(JSON.stringify({ ok: false, reason: String(e && e.message ? e.message : e) }))
      }
    })
    return
  }
  res.writeHead(200, { 'content-type': 'text/plain; charset=utf-8' })
  res.end('collab-server ok')
})

server.on('upgrade', (req, socket) => {
  const key = req.headers['sec-websocket-key']
  if (!key) {
    socket.destroy()
    return
  }
  socket.write(
    [
      'HTTP/1.1 101 Switching Protocols',
      'Upgrade: websocket',
      'Connection: Upgrade',
      `Sec-WebSocket-Accept: ${acceptKey(key)}`,
      '\r\n',
    ].join('\r\n'),
  )

  let buffer = Buffer.alloc(0)
  let bufferSize = 0
  let joined = null
  const send = (obj) => {
    if (!socket.writable) return
    socket.write(encodeTextFrame(JSON.stringify(obj)))
  }
  const broadcast = (obj, except) => {
    if (!joined) return
    const frame = encodeTextFrame(JSON.stringify(obj))
    for (const c of joined.clients) {
      if (c !== except && c.writable) c.write(frame)
    }
  }

  socket.on('data', (chunk) => {
    bufferSize += chunk.length
    if (bufferSize > MAX_MSG_BYTES) {
      // 单条消息超限，关闭连接防止 OOM
      socket.destroy()
      return
    }
    buffer = Buffer.concat([buffer, chunk])
    const { messages, rest } = decodeFrames(buffer)
    buffer = rest
    bufferSize = rest.length
    for (const m of messages) {
      if (m.opcode === 0x8) {
        socket.end()
        return
      }
      if (m.opcode !== 0x1) continue // 只处理文本
      let msg
      try {
        msg = JSON.parse(m.payload.toString('utf8'))
      } catch {
        continue
      }
      if (msg.t === 'join') {
        const roomId = String(msg.room || 'default')
        // 校验 room ID 格式
        if (!isValidRoomId(roomId)) {
          send({ t: 'error', reason: 'invalid room id' })
          socket.destroy()
          return
        }
        // 防止 room 数量无限增长
        if (!rooms.has(roomId) && rooms.size >= MAX_ROOMS) {
          send({ t: 'error', reason: 'server at capacity' })
          socket.destroy()
          return
        }
        joined = getRoom(roomId)
        joined.clients.add(socket)
        send({ t: 'snapshot', snapshot: snapshot(joined.room) })
      } else if (msg.t === 'op' && joined && msg.update) {
        const changed = applyUpdate(joined.room, msg.update)
        if (changed) broadcast({ t: 'op', update: msg.update }, socket)
      } else if (msg.t === 'awareness' && joined) {
        broadcast({ t: 'awareness', from: msg.from, state: msg.state }, socket)
      }
    }
  })

  const cleanup = () => {
    if (!joined) return
    joined.clients.delete(socket)
    // room 最后一个客户端离开后自动回收，防止空 room 无限积累
    if (joined.clients.size === 0) {
      for (const [id, r] of rooms) {
        if (r === joined) { rooms.delete(id); break }
      }
    }
  }
  socket.on('close', cleanup)
  socket.on('error', cleanup)
})

server.listen(PORT, () => {
  console.log(`[collab-server] listening on ws://localhost:${PORT}`)
})

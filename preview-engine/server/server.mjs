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
    buffer = Buffer.concat([buffer, chunk])
    const { messages, rest } = decodeFrames(buffer)
    buffer = rest
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
        joined = getRoom(String(msg.room || 'default'))
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
    if (joined) joined.clients.delete(socket)
  }
  socket.on('close', cleanup)
  socket.on('error', cleanup)
})

server.listen(PORT, () => {
  console.log(`[collab-server] listening on ws://localhost:${PORT}`)
})

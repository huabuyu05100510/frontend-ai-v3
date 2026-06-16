// ============================================================================
// convertClient — 旧版二进制 Office 的服务端转换客户端
//   POST 原始字节到本地 /convert，得到 OOXML 字节(base64) 或直接的 sheet model。
// ============================================================================

export interface ConvertSheetModel {
  name: string
  rows: number
  cols: number
  cells: { r: number; c: number; text: string }[]
}

export interface ConvertResult {
  ok: boolean
  format?: 'ooxml' | 'model'
  realType?: string
  base64?: string
  kind?: string
  model?: ConvertSheetModel
  via?: string
  reason?: string
  install?: string
}

const DEFAULT_ENDPOINT =
  (typeof location !== 'undefined' && (location as Location).protocol === 'https:'
    ? 'https://localhost:8787/convert'
    : 'http://localhost:8787/convert')

export async function convertLegacy(bytes: Uint8Array, ext: string, endpoint = DEFAULT_ENDPOINT): Promise<ConvertResult> {
  const res = await fetch(`${endpoint}?ext=${encodeURIComponent(ext)}`, {
    method: 'POST',
    body: bytes as BodyInit,
  })
  return (await res.json()) as ConvertResult
}

export function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64)
  const out = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i)
  return out
}

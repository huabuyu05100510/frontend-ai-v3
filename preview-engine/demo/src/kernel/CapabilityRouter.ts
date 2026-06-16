import type { ProbeResult, DeviceProfile, RouteDecision, RenderPath } from './types'

// ============================================================================
// CapabilityRouter — Native / WASM / Server 三态决策
//   不写死「客户端 or 服务端」，由「文件特征 × 设备能力」动态决定。
// ============================================================================

/** 浏览器通常可原生播放的媒体真实类型 */
const NATIVE_MEDIA = new Set(['mp4', 'm4v', 'mov', 'mp3', 'wav', 'aac', 'm4a'])
/** 仅需转封装即可（容器不支持，编码本身现代浏览器可解） */
const REMUX_MEDIA = new Set(['mkv', 'flv', 'avi', 'ts', 'mxf', 'wmv'])
/** 编码本身需软解 */
const TRANSCODE_MEDIA = new Set(['amr', 'wma', 's48', 'pcm'])
/** 有客户端解析器（WASM/JS），否则服务端兜底 */
const CLIENT_PARSE = new Set(['pdf', 'docx', 'xlsx', 'pptx'])
/** 老格式，无客户端解析器，必须服务端转换 */
const OLD_OFFICE = new Set(['doc', 'ppt', 'xls'])
/** 原生图片 */
const NATIVE_IMAGE = new Set(['png', 'jpg', 'gif', 'webp'])

function mediaMime(realType: string): string {
  const audio = new Set(['mp3', 'wav', 'aac', 'm4a', 'amr', 'wma', 's48', 'pcm'])
  return `${audio.has(realType) ? 'audio' : 'video'}/${realType}`
}

function decide(path: RenderPath, reason: string): RouteDecision {
  return { path, reason }
}

export function route(probe: ProbeResult, device: DeviceProfile): RouteDecision {
  const { realType, category } = probe

  if (category === 'unknown') return decide('server', '未知类型，服务端兜底转换')

  // 纯文本：流式渲染，无需解码器
  if (realType === 'txt' || realType === 'csv') return decide('native', '纯文本流式渲染')

  // 图片
  if (category === 'raster') {
    if (NATIVE_IMAGE.has(realType)) return decide('native', '浏览器原生解码图片')
    // bmp 等：原生失败用 WASM，弱机/无 WASM 服务端兜底
    if (device.canPlayType?.(`image/${realType}`)) return decide('native', '浏览器可解码')
    if (device.wasmEnabled) return decide('wasm', 'WASM 解码图片 + 瓦片切分')
    return decide('server', '无 WASM，服务端转 PNG')
  }

  // 媒体
  if (category === 'media') {
    if (NATIVE_MEDIA.has(realType) && device.canPlayType?.(mediaMime(realType)))
      return decide('native', '浏览器原生可播放')
    if (REMUX_MEDIA.has(realType)) {
      return device.wasmEnabled
        ? decide('wasm', 'WASM 转封装为 fMP4 + MSE 喂入')
        : decide('server', '无 WASM，服务端转 HLS')
    }
    if (TRANSCODE_MEDIA.has(realType)) {
      return device.wasmEnabled
        ? decide('wasm', 'ffmpeg.wasm 软解')
        : decide('server', '无 WASM，服务端转码')
    }
    // 原生集合但当前设备不支持 → 降级
    return device.wasmEnabled ? decide('wasm', 'WASM 解码兜底') : decide('server', '服务端转码兜底')
  }

  // 文档（paged / flow / sheet）
  if (OLD_OFFICE.has(realType)) return decide('server', '老 Office 格式，服务端转 OOXML/PDF')
  if (CLIENT_PARSE.has(realType)) {
    return device.wasmEnabled
      ? decide('wasm', '客户端解析（PDF.js / OOXML Parser）')
      : decide('server', '无 WASM，服务端转 PDF/图')
  }

  return decide('server', '默认服务端兜底')
}

// ============================================================================
// convert — 旧版二进制 Office 转换
//   优先 LibreOffice headless（高保真，doc/xls/ppt 通吃）；
//   无 LibreOffice 时，对 .xls 走内置零依赖 BIFF 解析回退。
// ============================================================================

import { spawnSync } from 'node:child_process'
import { mkdtempSync, writeFileSync, readFileSync, readdirSync, rmSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { parseXls } from './xlsLegacy.mjs'

const TARGET = { xls: 'xlsx', doc: 'docx', ppt: 'pptx' }
const CANDIDATES = ['soffice', 'libreoffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice']

export function findSoffice() {
  for (const c of CANDIDATES) {
    try {
      const r = spawnSync(c, ['--version'], { timeout: 5000 })
      if (r.status === 0) return c
    } catch {
      // 忽略，试下一个
    }
  }
  return null
}

export function convertWithSoffice(bin, bytes, ext) {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'cvt-'))
  try {
    const inFile = path.join(dir, 'in.' + ext)
    writeFileSync(inFile, Buffer.from(bytes))
    const target = TARGET[ext]
    const r = spawnSync(bin, ['--headless', '--convert-to', target, '--outdir', dir, inFile], { timeout: 60000 })
    if (r.status !== 0) throw new Error('LibreOffice 转换失败: ' + (r.stderr ? r.stderr.toString() : ''))
    const out = readdirSync(dir).find((f) => f.endsWith('.' + target))
    if (!out) throw new Error('LibreOffice 未产出目标文件')
    const data = readFileSync(path.join(dir, out))
    return { realType: target, base64: data.toString('base64') }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** 转换分发：返回结构化结果（ooxml 字节 base64 / 直接 sheet model / 不支持） */
export function convertLegacy(bytes, ext) {
  const e = (ext || '').toLowerCase()
  if (!TARGET[e]) return { ok: false, reason: `不支持的旧格式：.${e}` }

  const soffice = findSoffice()
  if (soffice) {
    const { realType, base64 } = convertWithSoffice(soffice, bytes, e)
    return { ok: true, format: 'ooxml', realType, base64, via: 'libreoffice' }
  }
  if (e === 'xls') {
    const model = parseXls(bytes)
    return { ok: true, format: 'model', kind: 'sheet', model, via: 'builtin-biff' }
  }
  return {
    ok: false,
    reason: `.${e} 为旧版二进制（CFB），高保真转换需服务端 LibreOffice（本机未安装）。`,
    install: 'brew install --cask libreoffice',
  }
}

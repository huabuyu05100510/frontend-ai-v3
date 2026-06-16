import { parseXml, findAll, firstChild, attr, firstText } from './xml'
import { ZipArchive } from './zip'

// ============================================================================
// xlsx — SpreadsheetML → SheetModel（纯逻辑，可测）
//   单元格值来源：t="s"→共享字符串，t="inlineStr"→内联，其余→数值/字符串。
// ============================================================================

export interface Cell {
  r: number
  c: number
  text: string
}
export interface SheetModel {
  name: string
  rows: number
  cols: number
  cells: Map<string, Cell> // key = "r,c"
}

export function colToIndex(col: string): number {
  let n = 0
  for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64)
  return n - 1
}

export function parseA1(ref: string): { r: number; c: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  if (!m) return { r: 0, c: 0 }
  return { r: parseInt(m[2], 10) - 1, c: colToIndex(m[1]) }
}

/** sharedStrings.xml → 字符串表（si 可含多个 run，需拼接） */
export function parseSharedStrings(xml: string): string[] {
  const root = parseXml(xml)
  return findAll(root, 'si').map((si) =>
    findAll(si, 't')
      .map(firstText)
      .join(''),
  )
}

export function parseSheet(name: string, sheetXml: string, shared: string[]): SheetModel {
  const root = parseXml(sheetXml)
  const cells = new Map<string, Cell>()
  let maxR = -1
  let maxC = -1
  for (const c of findAll(root, 'c')) {
    const ref = attr(c, 'r')
    if (!ref) continue
    const { r, col } = (() => {
      const p = parseA1(ref)
      return { r: p.r, col: p.c }
    })()
    const t = attr(c, 't')
    let text = ''
    if (t === 's') {
      const v = firstChild(c, 'v')
      const idx = v ? parseInt(firstText(v), 10) : NaN
      text = Number.isFinite(idx) ? (shared[idx] ?? '') : ''
    } else if (t === 'inlineStr') {
      const is = firstChild(c, 'is')
      text = is
        ? findAll(is, 't')
            .map(firstText)
            .join('')
        : ''
    } else {
      const v = firstChild(c, 'v')
      text = v ? firstText(v) : ''
    }
    if (text === '') continue
    cells.set(`${r},${col}`, { r, c: col, text })
    if (r > maxR) maxR = r
    if (col > maxC) maxC = col
  }
  return { name, rows: maxR + 1, cols: maxC + 1, cells }
}

/** 从 XLSX 字节加载首个工作表 */
export async function loadXlsx(bytes: Uint8Array): Promise<SheetModel> {
  const zip = await ZipArchive.open(bytes)
  const shared = zip.has('xl/sharedStrings.xml') ? parseSharedStrings(await zip.text('xl/sharedStrings.xml')) : []
  const sheetName =
    zip.names().filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0] ?? 'xl/worksheets/sheet1.xml'
  const sheetXml = await zip.text(sheetName)
  return parseSheet('Sheet1', sheetXml, shared)
}

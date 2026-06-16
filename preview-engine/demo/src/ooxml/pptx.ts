import { parseXml, findAll, firstChild, attr, firstText, type XmlNode } from './xml'
import { ZipArchive } from './zip'

// ============================================================================
// pptx — PresentationML(slideN.xml) → Slide[]（纯逻辑，可测）
//   每个形状(p:sp)提取文本(a:t，段落用换行拼接)与位置(a:off，EMU 单位)。
// ============================================================================

export interface SlideText {
  text: string
  x: number
  y: number
}
export interface Slide {
  index: number
  texts: SlideText[]
}

function shapeText(sp: XmlNode): string {
  const txBody = firstChild(sp, 'p:txBody')
  if (!txBody) return ''
  return findAll(txBody, 'a:p')
    .map((p) =>
      findAll(p, 'a:t')
        .map(firstText)
        .join(''),
    )
    .join('\n')
}

function shapeOffset(sp: XmlNode): { x: number; y: number } {
  const off = findAll(sp, 'a:off')[0]
  if (!off) return { x: 0, y: 0 }
  return { x: parseInt(attr(off, 'x') ?? '0', 10) || 0, y: parseInt(attr(off, 'y') ?? '0', 10) || 0 }
}

export function parseSlide(index: number, slideXml: string): Slide {
  const root = parseXml(slideXml)
  const texts: SlideText[] = []
  for (const sp of findAll(root, 'p:sp')) {
    const text = shapeText(sp)
    if (!text) continue
    const { x, y } = shapeOffset(sp)
    texts.push({ text, x, y })
  }
  return { index, texts }
}

/** 从 PPTX 字节加载全部幻灯片（按编号排序） */
export async function loadPptx(bytes: Uint8Array): Promise<Slide[]> {
  const zip = await ZipArchive.open(bytes)
  const slideNames = zip
    .names()
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(/slide(\d+)/.exec(a)![1], 10)
      const nb = parseInt(/slide(\d+)/.exec(b)![1], 10)
      return na - nb
    })
  const slides: Slide[] = []
  for (let i = 0; i < slideNames.length; i++) {
    slides.push(parseSlide(i, await zip.text(slideNames[i])))
  }
  return slides
}

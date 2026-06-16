import { describe, it, expect } from 'vitest'
import { parseSlide } from '../pptx'

const shape = (text: string, x = 0, y = 0) =>
  `<p:sp><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="100" cy="50"/></a:xfrm></p:spPr>` +
  `<p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

const slide = (shapes: string) =>
  `<p:sld xmlns:p="ns" xmlns:a="ns"><p:cSld><p:spTree>${shapes}</p:spTree></p:cSld></p:sld>`

describe('parseSlide', () => {
  it('提取每个文本框的文本与坐标', () => {
    const s = parseSlide(3, slide(shape('标题', 100, 200) + shape('正文', 300, 400)))
    expect(s.index).toBe(3)
    expect(s.texts).toHaveLength(2)
    expect(s.texts[0]).toEqual({ text: '标题', x: 100, y: 200 })
    expect(s.texts[1]).toEqual({ text: '正文', x: 300, y: 400 })
  })

  it('多段落文本用换行拼接', () => {
    const sp =
      '<p:sp><p:txBody><a:p><a:r><a:t>第一行</a:t></a:r></a:p><a:p><a:r><a:t>第二行</a:t></a:r></a:p></p:txBody></p:sp>'
    const s = parseSlide(1, slide(sp))
    expect(s.texts[0].text).toBe('第一行\n第二行')
  })

  it('一个段落多 run 直接拼接', () => {
    const sp = '<p:sp><p:txBody><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>World</a:t></a:r></a:p></p:txBody></p:sp>'
    const s = parseSlide(1, slide(sp))
    expect(s.texts[0].text).toBe('Hello World')
  })

  it('忽略无文本的形状', () => {
    const empty = '<p:sp><p:spPr><a:xfrm><a:off x="1" y="2"/></a:xfrm></p:spPr></p:sp>'
    const s = parseSlide(1, slide(empty + shape('有字')))
    expect(s.texts).toHaveLength(1)
    expect(s.texts[0].text).toBe('有字')
  })

  it('无坐标默认 0', () => {
    const sp = '<p:sp><p:txBody><a:p><a:r><a:t>X</a:t></a:r></a:p></p:txBody></p:sp>'
    const s = parseSlide(1, slide(sp))
    expect(s.texts[0]).toEqual({ text: 'X', x: 0, y: 0 })
  })
})

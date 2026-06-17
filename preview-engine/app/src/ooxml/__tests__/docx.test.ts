import { describe, it, expect } from 'vitest'
import { parseDocx, parseRels } from '../docx'

const doc = (body: string) =>
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`

describe('parseDocx', () => {
  it('普通段落与文本拼接', () => {
    const blocks = parseDocx(doc('<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>'))
    expect(blocks).toHaveLength(1)
    expect(blocks[0].type).toBe('paragraph')
    if (blocks[0].type === 'paragraph') {
      expect(blocks[0].runs.map((r) => r.text).join('')).toBe('Hello World')
    }
  })

  it('标题（pStyle=Heading1）识别等级', () => {
    const blocks = parseDocx(doc('<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>章节</w:t></w:r></w:p>'))
    expect(blocks[0].type).toBe('heading')
    if (blocks[0].type === 'heading') {
      expect(blocks[0].level).toBe(2)
      expect(blocks[0].runs[0].text).toBe('章节')
    }
  })

  it('加粗/斜体 run 样式', () => {
    const blocks = parseDocx(
      doc('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>粗</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>斜</w:t></w:r></w:p>'),
    )
    if (blocks[0].type === 'paragraph') {
      expect(blocks[0].runs[0].bold).toBe(true)
      expect(blocks[0].runs[1].italic).toBe(true)
    }
  })

  it('w:b w:val="false" 不算加粗', () => {
    const blocks = parseDocx(doc('<w:p><w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>x</w:t></w:r></w:p>'))
    if (blocks[0].type === 'paragraph') {
      expect(blocks[0].runs[0].bold).toBeFalsy()
    }
  })

  it('表格解析为行列单元格', () => {
    const tbl =
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const blocks = parseDocx(doc(tbl))
    expect(blocks[0].type).toBe('table')
    if (blocks[0].type === 'table') {
      expect(blocks[0].rows.length).toBe(2)
      expect(blocks[0].rows[0].length).toBe(2)
      const cellText = (cell: { text: string }[]) => cell.map((r) => r.text).join('')
      expect(cellText(blocks[0].rows[0][0])).toBe('A1')
      expect(cellText(blocks[0].rows[1][1])).toBe('B2')
    }
  })

  it('保持块顺序', () => {
    const blocks = parseDocx(
      doc(
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>正文</w:t></w:r></w:p>',
      ),
    )
    expect(blocks.map((b) => b.type)).toEqual(['heading', 'paragraph'])
  })

  it('空段落产出空 runs 段落', () => {
    const blocks = parseDocx(doc('<w:p/>'))
    expect(blocks[0].type).toBe('paragraph')
    if (blocks[0].type === 'paragraph') expect(blocks[0].runs).toEqual([])
  })

  it('解析段落对齐（center / 居中）', () => {
    const blocks = parseDocx(doc('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>中</w:t></w:r></w:p>'))
    expect(blocks[0].type).toBe('paragraph')
    if (blocks[0].type === 'paragraph') expect(blocks[0].align).toBe('center')
  })

  it('w:jc=both 归一为 justify', () => {
    const blocks = parseDocx(doc('<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'))
    if (blocks[0].type === 'paragraph') expect(blocks[0].align).toBe('justify')
  })

  it('提取内嵌图片为 image 块（rId + 尺寸 EMU→px）', () => {
    const drawing =
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/>' +
      '<a:graphic><a:graphicData><pic:pic><pic:blipFill>' +
      '<a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r></w:p>'
    const blocks = parseDocx(doc(drawing))
    const img = blocks.find((b) => b.type === 'image')
    expect(img).toBeTruthy()
    if (img && img.type === 'image') {
      expect(img.rId).toBe('rId7')
      expect(img.width).toBe(200) // 1905000 / 9525
      expect(img.height).toBe(100)
    }
  })

  it('图文混排：文本块在前，图片块随后', () => {
    const mixed =
      '<w:p><w:r><w:t>说明</w:t></w:r>' +
      '<w:r><w:drawing><a:blip r:embed="rId3"/></w:drawing></w:r></w:p>'
    const blocks = parseDocx(doc(mixed))
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'image'])
  })
})

describe('parseRels', () => {
  it('解析关系表 Id→Target', () => {
    const xml =
      '<Relationships><Relationship Id="rId7" Type="image" Target="media/image1.png"/>' +
      '<Relationship Id="rId1" Type="styles" Target="styles.xml"/></Relationships>'
    expect(parseRels(xml)).toEqual({ rId7: 'media/image1.png', rId1: 'styles.xml' })
  })
})

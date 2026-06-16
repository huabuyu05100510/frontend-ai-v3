import { useEffect, useState } from 'react'
import type { SourceHandle } from '../../kernel/SourceHandle'
import { loadDocx, type Block } from '../../ooxml/docx'
import { loadXlsx, type SheetModel } from '../../ooxml/xlsx'
import { loadPptx, type Slide } from '../../ooxml/pptx'
import {
  extractDocxUnits,
  extractSheetUnits,
  extractSlideUnits,
  translateAll,
} from '../../translation/translate'

type Loaded =
  | { kind: 'docx'; blocks: Block[] }
  | { kind: 'sheet'; sheet: SheetModel }
  | { kind: 'pptx'; slides: Slide[] }

export type Mode = 'source' | 'both' | 'target'

const SLIDE_W = 720
const SLIDE_H = 405
const DEFAULT_EMU_W = 12192000 // 13.33in 16:9

export function OfficeView({ source, realType }: { source: SourceHandle; realType: string }) {
  const [data, setData] = useState<Loaded | null>(null)
  const [trans, setTrans] = useState<Record<string, string>>({})
  const [mode, setMode] = useState<Mode>('source')
  const [err, setErr] = useState<string | null>(null)

  useEffect(() => {
    let off = false
    setData(null)
    setErr(null)
    source
      .blob()
      .arrayBuffer()
      .then(async (buf) => {
        const bytes = new Uint8Array(buf)
        if (realType === 'docx') {
          const blocks = await loadDocx(bytes)
          if (off) return
          setData({ kind: 'docx', blocks })
          setTrans(translateAll(extractDocxUnits(blocks)))
        } else if (realType === 'xlsx') {
          const sheet = await loadXlsx(bytes)
          if (off) return
          setData({ kind: 'sheet', sheet })
          setTrans(translateAll(extractSheetUnits(sheet)))
        } else if (realType === 'pptx') {
          const slides = await loadPptx(bytes)
          if (off) return
          setData({ kind: 'pptx', slides })
          setTrans(translateAll(extractSlideUnits(slides)))
        }
      })
      .catch((e) => !off && setErr(String(e?.message || e)))
    return () => {
      off = true
    }
  }, [source, realType])

  if (err) return <div className="panel" style={{ color: 'var(--red)' }}>解析失败：{err}</div>
  if (!data) return <div className="panel">解析 OOXML 中…</div>

  return (
    <div>
      <div className="panel" style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
        <span className="kv" style={{ marginRight: 4 }}>视图</span>
        {(['source', 'both', 'target'] as Mode[]).map((m) => (
          <button
            key={m}
            onClick={() => setMode(m)}
            style={{
              borderColor: mode === m ? 'var(--accent)' : 'var(--border)',
              color: mode === m ? 'var(--accent)' : undefined,
            }}
          >
            {m === 'source' ? '原文' : m === 'target' ? '译文' : '双栏对比'}
          </button>
        ))}
        <span className="kv" style={{ marginLeft: 'auto' }}>结构化解析 · 文本可复制 · 译文原位回填</span>
      </div>
      {data.kind === 'docx' && <DocxBody blocks={data.blocks} trans={trans} mode={mode} />}
      {data.kind === 'sheet' && <SheetBody sheet={data.sheet} trans={trans} mode={mode} />}
      {data.kind === 'pptx' && <PptxBody slides={data.slides} trans={trans} mode={mode} />}
    </div>
  )
}

// ---------------------------------------------------------------- DOCX --------
function runStyle(r: { bold?: boolean; italic?: boolean }): React.CSSProperties {
  return { fontWeight: r.bold ? 700 : undefined, fontStyle: r.italic ? 'italic' : undefined }
}

function DocxBlockView({ block, target }: { block: Block; target?: string }) {
  if (block.type === 'table') {
    return (
      <table style={{ borderCollapse: 'collapse', width: '100%', margin: '8px 0', fontSize: 13 }}>
        <tbody>
          {block.rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td key={ci} style={{ border: '1px solid var(--border)', padding: '4px 8px' }}>
                  {cell.map((r) => r.text).join('')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    )
  }
  if (block.type === 'image') {
    const justify = block.align === 'center' ? 'center' : block.align === 'right' ? 'flex-end' : 'flex-start'
    return (
      <div style={{ display: 'flex', justifyContent: justify, margin: '8px 0' }}>
        {block.src ? (
          <img
            src={block.src}
            style={{ maxWidth: '100%', width: block.width ? Math.min(block.width, 640) : undefined, height: 'auto', borderRadius: 4 }}
          />
        ) : (
          <div className="kv" style={{ border: '1px dashed var(--border)', padding: 12, borderRadius: 4 }}>
            🖼 图片（{block.width ?? '?'}×{block.height ?? '?'}）
          </div>
        )}
      </div>
    )
  }
  const textAlign = block.align as React.CSSProperties['textAlign'] | undefined
  if (block.type === 'heading') {
    const size = [22, 19, 17, 15, 14, 13][Math.min(5, block.level - 1)]
    const text = target !== undefined ? target : block.runs.map((r) => r.text).join('')
    return <div style={{ fontSize: size, fontWeight: 700, margin: '14px 0 6px', textAlign }}>{text}</div>
  }
  if (target !== undefined) return <p style={{ margin: '6px 0', lineHeight: 1.7, textAlign }}>{target}</p>
  return (
    <p style={{ margin: '6px 0', lineHeight: 1.7, textAlign }}>
      {block.runs.map((r, i) => (
        <span key={i} style={runStyle(r)}>
          {r.text}
        </span>
      ))}
    </p>
  )
}

function DocxBody({ blocks, trans, mode }: { blocks: Block[]; trans: Record<string, string>; mode: Mode }) {
  const Col = ({ useTarget }: { useTarget: boolean }) => (
    <div className="viewport" style={{ height: 560, padding: '8px 18px', background: '#fff', color: '#111' }}>
      {blocks.map((b, i) => (
        <DocxBlockView key={i} block={b} target={useTarget ? trans[`b${i}`] ?? '' : undefined} />
      ))}
    </div>
  )
  if (mode === 'both') {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
        <div>
          <div className="kv" style={{ marginBottom: 4 }}>原文</div>
          <Col useTarget={false} />
        </div>
        <div>
          <div className="kv" style={{ marginBottom: 4 }}>译文</div>
          <Col useTarget />
        </div>
      </div>
    )
  }
  return <Col useTarget={mode === 'target'} />
}

// ---------------------------------------------------------------- XLSX --------
function SheetTable({ sheet, trans, useTarget }: { sheet: SheetModel; trans: Record<string, string>; useTarget: boolean }) {
  const maxRows = Math.min(sheet.rows, 500)
  const maxCols = Math.min(sheet.cols, 60)
  const rows = Array.from({ length: maxRows }, (_, r) => r)
  const cols = Array.from({ length: maxCols }, (_, c) => c)
  return (
    <div className="viewport" style={{ height: 560, background: '#fff', color: '#111' }}>
      <table style={{ borderCollapse: 'collapse', fontSize: 12 }}>
        <tbody>
          {rows.map((r) => (
            <tr key={r}>
              {cols.map((c) => {
                const cell = sheet.cells.get(`${r},${c}`)
                const text = cell ? (useTarget ? trans[`${r},${c}`] ?? cell.text : cell.text) : ''
                return (
                  <td
                    key={c}
                    style={{
                      border: '1px solid #e2e6ee',
                      padding: '2px 6px',
                      minWidth: 64,
                      maxWidth: 220,
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                  >
                    {text}
                  </td>
                )
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export function SheetBody({ sheet, trans, mode }: { sheet: SheetModel; trans: Record<string, string>; mode: Mode }) {
  const note = (
    <div className="kv" style={{ marginBottom: 4 }}>
      {sheet.name} · {sheet.rows.toLocaleString()} 行 × {sheet.cols} 列
      {(sheet.rows > 500 || sheet.cols > 60) && '（已截断渲染，生产环境走视口虚拟化）'}
    </div>
  )
  if (mode === 'both') {
    return (
      <div>
        {note}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <div>
            <div className="kv">原文</div>
            <SheetTable sheet={sheet} trans={trans} useTarget={false} />
          </div>
          <div>
            <div className="kv">译文</div>
            <SheetTable sheet={sheet} trans={trans} useTarget />
          </div>
        </div>
      </div>
    )
  }
  return (
    <div>
      {note}
      <SheetTable sheet={sheet} trans={trans} useTarget={mode === 'target'} />
    </div>
  )
}

// ---------------------------------------------------------------- PPTX --------
function SlideFrame({ slide, trans, useTarget }: { slide: Slide; trans: Record<string, string>; useTarget: boolean }) {
  const scale = SLIDE_W / DEFAULT_EMU_W
  return (
    <div
      style={{
        position: 'relative',
        width: SLIDE_W,
        height: SLIDE_H,
        background: '#fff',
        color: '#111',
        border: '1px solid var(--border)',
        borderRadius: 6,
        overflow: 'hidden',
        margin: '0 auto',
      }}
    >
      {slide.texts.map((t, i) => {
        const text = useTarget ? trans[`s${slide.index}.t${i}`] ?? t.text : t.text
        return (
          <div
            key={i}
            style={{
              position: 'absolute',
              left: Math.max(0, t.x * scale),
              top: Math.max(0, t.y * scale),
              maxWidth: SLIDE_W - 16,
              fontSize: 14,
              lineHeight: 1.4,
              whiteSpace: 'pre-wrap',
            }}
          >
            {text}
          </div>
        )
      })}
    </div>
  )
}

function PptxBody({ slides, trans, mode }: { slides: Slide[]; trans: Record<string, string>; mode: Mode }) {
  return (
    <div className="viewport" style={{ height: 560 }}>
      {slides.map((s) => (
        <div key={s.index} style={{ marginBottom: 18 }}>
          <div className="kv" style={{ marginBottom: 6 }}>第 {s.index + 1} 页</div>
          {mode === 'both' ? (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
              <SlideFrame slide={s} trans={trans} useTarget={false} />
              <SlideFrame slide={s} trans={trans} useTarget />
            </div>
          ) : (
            <SlideFrame slide={s} trans={trans} useTarget={mode === 'target'} />
          )}
        </div>
      ))}
    </div>
  )
}

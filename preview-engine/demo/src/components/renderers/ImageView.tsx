import { useEffect, useRef, useState } from 'react'
import type { SourceHandle } from '../../kernel/SourceHandle'
import { fitScale, clampScale } from '../../renderers/image/fit'

// 真实图片渲染：createImageBitmap 解码 → Canvas 绘制 + 滚轮缩放
export function ImageView({ source }: { source: SourceHandle }) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const bitmapRef = useRef<ImageBitmap | null>(null)
  const [scale, setScale] = useState(1)
  const [info, setInfo] = useState('')
  const [err, setErr] = useState('')

  useEffect(() => {
    let disposed = false
    ;(async () => {
      try {
        const bmp = await createImageBitmap(source.blob())
        if (disposed) {
          bmp.close()
          return
        }
        bitmapRef.current = bmp
        const vp = wrapRef.current!.getBoundingClientRect()
        const s = fitScale({ width: bmp.width, height: bmp.height }, { width: vp.width - 24, height: 520 })
        setScale(s)
        setInfo(`${bmp.width}×${bmp.height}px`)
      } catch (e) {
        setErr('图片解码失败：' + String(e))
      }
    })()
    return () => {
      disposed = true
      bitmapRef.current?.close()
      bitmapRef.current = null
    }
  }, [source])

  useEffect(() => {
    const bmp = bitmapRef.current
    const canvas = canvasRef.current
    if (!bmp || !canvas) return
    const dpr = Math.min(window.devicePixelRatio || 1, 2)
    const w = bmp.width * scale
    const h = bmp.height * scale
    canvas.width = Math.round(w * dpr)
    canvas.height = Math.round(h * dpr)
    canvas.style.width = `${w}px`
    canvas.style.height = `${h}px`
    const ctx = canvas.getContext('2d')!
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, w, h)
    ctx.drawImage(bmp, 0, 0, w, h)
  }, [scale, info])

  const onWheel = (e: React.WheelEvent) => {
    e.preventDefault()
    setScale((s) => clampScale(s * (e.deltaY < 0 ? 1.1 : 0.9), 0.05, 8))
  }

  if (err) return <div className="panel" style={{ color: 'var(--red)' }}>{err}</div>
  return (
    <div>
      <div className="kv">
        {info} · 缩放 <b>{Math.round(scale * 100)}%</b>（滚轮缩放）
      </div>
      <div
        ref={wrapRef}
        onWheel={onWheel}
        style={{ height: 540, overflow: 'auto', background: '#0b0e13', borderRadius: 8, display: 'flex', justifyContent: 'center', alignItems: 'flex-start', padding: 12 }}
      >
        <canvas ref={canvasRef} style={{ borderRadius: 4 }} />
      </div>
    </div>
  )
}

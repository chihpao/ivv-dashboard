import React, { useMemo, useRef } from 'react'
import { useSheets } from './useSheets'
import {
  LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer,
  BarChart, Bar, Legend, AreaChart, Area, FunnelChart, Funnel, LabelList
} from 'recharts'
import html2canvas from 'html2canvas'
import Papa from 'papaparse'

export default function Dashboard() {
  const { kpi, trend, top5, sla, categoryTrend, loading } = useSheets()
  // 不要在 loading 時提前 return，以避免 Hooks 順序變動

  const k = kpi[0] ?? {}

  // 有些表頭命名可能不同（例如：未結案件數/未結案數、SLA達成/SLA達成率）
  // 這裡提供同義鍵的 fallback，避免顯示成 "-"。
  const pick = (obj: Record<string, any>, candidates: string[], fallback: any = '-') => {
    for (const key of candidates) {
      if (key in obj && obj[key] !== undefined && obj[key] !== null && obj[key] !== '') return obj[key]
    }
    return fallback
  }

  const trendRows = trend.map((r: any) => ({
    // 預設第一欄為日期、第二欄為數量
    date: String(r[Object.keys(r)[0]]),
    count: Number(r[Object.keys(r)[1]]) || 0,
  }))

  // 7/30 日移動平均線
  const trendWithMA = useMemo(() => {
    const ma = (arr: number[], window: number) => arr.map((_, i) => {
      if (i + 1 < window) return undefined
      let s = 0
      for (let j = i - window + 1; j <= i; j++) s += arr[j]
      return +(s / window).toFixed(2)
    })
    const counts = trendRows.map(d => d.count)
    const ma7 = ma(counts, 7)
    const ma30 = ma(counts, 30)
    return trendRows.map((d, i) => ({ ...d, ma7: ma7[i], ma30: ma30[i] }))
  }, [trendRows])

  const topRows = top5.map((r: any) => {
    const keys = Object.keys(r)
    return { name: String(r[keys[0]]), value: Number(r[keys[1]]) || 0 }
  })

  // SLA：可吃原始 resolve_minute 或已彙總資料
  const slaRows = useMemo(() => {
    if (!sla || sla.length === 0) return [] as { name: string; value: number }[]
    const sample = sla[0]
    const keys = Object.keys(sample)
    if (keys.includes('resolve_minute')) {
      const minutes = sla
        .map((r: any) => Number(r['resolve_minute']))
        .filter((n: number) => Number.isFinite(n) && n >= 0)
      const bins = [30, 60, 120, 240, 480]
      const counts: Record<string, number> = {}
      const label = (min: number, max?: number) => (max ? `${min}-${max} 分` : `>${min} 分`)
      for (const n of minutes) {
        let placed = false
        let prev = 0
        for (const b of bins) {
          if (n <= b) {
            const l = label(prev === 0 ? 0 : prev + 1, b)
            counts[l] = (counts[l] ?? 0) + 1
            placed = true
            break
          }
          prev = b
        }
        if (!placed) {
          const l = label(bins[bins.length - 1])
          counts[l] = (counts[l] ?? 0) + 1
        }
      }
      return Object.entries(counts).map(([name, value]) => ({ name, value }))
    }
    // 已彙總：取前兩欄
    return sla.map((r: any) => ({ name: String(r[keys[0]]), value: Number(r[keys[1]]) || 0 }))
  }, [sla])

  // 分類堆疊面積圖：第一欄日期、其餘為分類
  const categoryCfg = useMemo(() => {
    if (!categoryTrend || categoryTrend.length === 0) return { dateKey: '', keys: [] as string[], rows: [] as any[] }
    const first = categoryTrend[0]
    const keys = Object.keys(first)
    const dateKey = keys[0]
    const seriesKeys = keys.slice(1)
    const rows = categoryTrend.map((r: any) => ({
      [dateKey]: String(r[dateKey]),
      ...Object.fromEntries(seriesKeys.map(k => [k, Number(r[k]) || 0])),
    }))
    return { dateKey, keys: seriesKeys, rows }
  }, [categoryTrend])

  // 匯出：CSV/PNG
  const downloadCSV = (filename: string, rows: any[]) => {
    const csv = Papa.unparse(rows)
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    URL.revokeObjectURL(url)
  }

  const downloadPNG = async (el: HTMLElement | null, filename: string) => {
    if (!el) return
    const canvas = await html2canvas(el, { backgroundColor: '#fff', scale: 2 })
    const link = document.createElement('a')
    link.download = filename
    link.href = canvas.toDataURL('image/png')
    link.click()
  }

  const trendRef = useRef<HTMLDivElement>(null)
  const topRef = useRef<HTMLDivElement>(null)
  const slaRef = useRef<HTMLDivElement>(null)
  const catRef = useRef<HTMLDivElement>(null)

  return (
    <div style={{ padding: 16, display: 'grid', gap: 16 }}>
      {loading && (
        <div style={{ padding: 8, color: '#666' }}>資料載入中…</div>
      )}

      {/* KPI */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 12 }}>
        <Kpi
          title="本月總件數"
          value={pick(k, [
            '本月總件數', '本月件數', '本月總計', '本月新增數', '本月新增件數',
          ])}
        />
        <Kpi
          title="解決率"
          value={pick(k, [
            '解決率', '解決率%', '本月解決率',
          ])}
        />
        <Kpi
          title="平均處理時長"
          value={pick(k, [
            '平均處理時長', '平均處理時長(分)', '平均處理時長(分鐘)', '平均處理時間(分)', '平均處理時間(分鐘)',
          ])}
        />
        <Kpi
          title="未結案件數"
          value={pick(k, [
            '未結案件數', '未結案數', '未結數', '未處理案件數',
          ])}
        />
        <Kpi
          title="SLA達成率"
          value={pick(k, [
            'SLA達成率', 'SLA達成', 'SLA達成%', 'SLA達成率(%)',
          ])}
        />
      </div>

      {/* 每日趨勢 + 7/30 日均線 */}
      <Card
        title="每日趨勢（含7/30日均線）"
        actions={<Actions onPNG={() => downloadPNG(trendRef.current, 'trend.png')} onCSV={() => downloadCSV('trend.csv', trendWithMA)} />}
      >
        <div ref={trendRef}>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={trendWithMA}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="date" />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="count" name="每日" stroke="#8884d8" dot={false} />
              <Line type="monotone" dataKey="ma7" name="7日均" stroke="#ff7300" dot={false} />
              <Line type="monotone" dataKey="ma30" name="30日均" stroke="#00c49f" dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* 模組 Top 5 */}
      <Card
        title="模組 Top 5"
        actions={<Actions onPNG={() => downloadPNG(topRef.current, 'top5.png')} onCSV={() => downloadCSV('top5.csv', topRows)} />}
      >
        <div ref={topRef}>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={topRows}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Bar dataKey="value" fill="#8884d8" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Card>

      {/* SLA 漏斗圖 */}
      {slaRows.length > 0 ? (
        <Card
          title="SLA 處理時效分佈（Funnel）"
          actions={<Actions onPNG={() => downloadPNG(slaRef.current, 'sla-funnel.png')} onCSV={() => downloadCSV('sla-funnel.csv', slaRows)} />}
        >
          <div ref={slaRef}>
            <ResponsiveContainer width="100%" height={320}>
              <FunnelChart>
                <Tooltip />
                <Funnel dataKey="value" data={slaRows} nameKey="name" isAnimationActive>
                  <LabelList position="right" fill="#000" stroke="none" dataKey="name" />
                </Funnel>
              </FunnelChart>
            </ResponsiveContainer>
          </div>
        </Card>
      ) : (
        <Card title="SLA 處理時效分佈（Funnel）">
          <div style={{ padding: 8, color: '#666' }}>尚未設定 SLA CSV 來源（到 src/useSheets.ts 設定 SLA_URL）</div>
        </Card>
      )}

      {/* 分類堆疊面積圖 */}
      {categoryCfg.rows.length > 0 ? (
        <Card
          title="分類堆疊面積圖"
          actions={<Actions onPNG={() => downloadPNG(catRef.current, 'category-stacked.png')} onCSV={() => downloadCSV('category-stacked.csv', categoryCfg.rows)} />}
        >
          <div ref={catRef}>
            <ResponsiveContainer width="100%" height={320}>
              <AreaChart data={categoryCfg.rows}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey={categoryCfg.dateKey} />
                <YAxis allowDecimals={false} />
                <Tooltip />
                <Legend />
                {categoryCfg.keys.map((k, i) => (
                  <Area key={k} type="monotone" dataKey={k} stackId="1" stroke={palette[i % palette.length]} fill={palette[i % palette.length]} />
                ))}
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </Card>
      ) : (
        <Card title="分類堆疊面積圖">
          <div style={{ padding: 8, color: '#666' }}>尚未設定分類趨勢 CSV 來源（到 src/useSheets.ts 設定 CATEGORY_TREND_URL）</div>
        </Card>
      )}
    </div>
  )
}

function Kpi({ title, value }: { title: string; value: any }) {
  return (
    <div style={{ padding: 12, border: '1px solid #eee', borderRadius: 12 }}>
      <div style={{ fontSize: 12, opacity: 0.7 }}>{title}</div>
      <div style={{ fontSize: 24, fontWeight: 700, marginTop: 6 }}>{String(value ?? '-')}</div>
    </div>
  )
}

function Card({ title, children, actions }: { title: string; children: React.ReactNode; actions?: React.ReactNode }) {
  return (
    <div style={{ border: '1px solid #eee', borderRadius: 12, padding: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, marginBottom: 8 }}>
        <div style={{ fontSize: 14, fontWeight: 600 }}>{title}</div>
        {actions}
      </div>
      {children}
    </div>
  )
}

function Actions({ onPNG, onCSV }: { onPNG?: () => void; onCSV?: () => void }) {
  return (
    <div style={{ display: 'flex', gap: 8 }}>
      {onPNG && (
        <button onClick={onPNG} style={btnStyle}>
          下載 PNG
        </button>
      )}
      {onCSV && (
        <button onClick={onCSV} style={btnStyle}>
          下載 CSV
        </button>
      )}
    </div>
  )
}

const btnStyle: React.CSSProperties = {
  padding: '6px 10px',
  fontSize: 12,
  border: '1px solid #ddd',
  borderRadius: 8,
  background: '#fff',
  cursor: 'pointer',
}

const palette = ['#8884d8', '#82ca9d', '#ffc658', '#ff7f50', '#00c49f', '#a28fd0', '#8dd1e1', '#d0ed57', '#a4de6c']

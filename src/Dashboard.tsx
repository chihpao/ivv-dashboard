import React from 'react'
import { useSheets } from "./useSheets"
import { LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer, BarChart, Bar } from "recharts"

export default function Dashboard() {
  const { kpi, trend, top5, loading } = useSheets()
  if (loading) return <div style={{padding:16}}>載入中…</div>

  // KPI：假設 kpi CSV 第一列就包含欄位：本月總件數, 解決率, 平均處理時長, 未結案數, SLA 達成率
  const k = kpi[0] ?? {}

  // trend CSV：應有欄位「日期」「件數」（我們在公式已經 label 過）
  const trendRows = trend.map(r => ({ date: String(r["日期"]), count: Number(r["件數"]) || 0 }))

  // top5 CSV：應有兩欄（依你在 module_top5 的 QUERY 設定）
  // 若你的欄位名是「module, count」，就改成 r.module / r.count
  const topRows = top5.map(r => {
    const keys = Object.keys(r)
    return { name: String(r[keys[0]]), value: Number(r[keys[1]]) || 0 }
  })

  return (
    <div style={{padding:16, display:"grid", gap:16}}>
      {/* KPI 區 */}
      <div style={{display:"grid", gridTemplateColumns:"repeat(5, 1fr)", gap:12}}>
        <Kpi title="本月總件數" value={k["本月總件數"] ?? "-"} />
        <Kpi title="解決率" value={k["解決率"] ?? "-"} />
        <Kpi title="平均處理時長(分)" value={k["平均處理時長"] ?? "-"} />
        <Kpi title="未結案數" value={k["未結案數"] ?? "-"} />
        <Kpi title="SLA 達成率" value={k["SLA 達成率"] ?? "-"} />
      </div>

      {/* 日趨勢折線圖 */}
      <Card title="日趨勢（件數）">
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={trendRows}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="date" />
            <YAxis allowDecimals={false} />
            <Tooltip />
            <Line type="monotone" dataKey="count" />
          </LineChart>
        </ResponsiveContainer>
      </Card>

      {/* 模組 Top5 長條圖 */}
      <Card title="模組別 Top 5（件數）">
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={topRows}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis allowDecimals={false} />
            <Tooltip />
            <Bar dataKey="value" />
          </BarChart>
        </ResponsiveContainer>
      </Card>
    </div>
  )
}

function Kpi({ title, value }: { title: string; value: any }) {
  return (
    <div style={{padding:12, border:"1px solid #eee", borderRadius:12}}>
      <div style={{fontSize:12, opacity:.7}}>{title}</div>
      <div style={{fontSize:24, fontWeight:700, marginTop:6}}>{String(value ?? "-")}</div>
    </div>
  )
}

function Card({ title, children }: any) {
  return (
    <div style={{border:"1px solid #eee", borderRadius:12, padding:12}}>
      <div style={{fontSize:14, fontWeight:600, marginBottom:8}}>{title}</div>
      {children}
    </div>
  )
}

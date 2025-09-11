import { useMemo, useRef, useState } from "react";
import { useSheets } from "./useSheets";
import dayjs from "dayjs";
import html2canvas from "html2canvas";
import {
  LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer,
  BarChart, Bar, Legend
} from "recharts";
import "./Dashboard.css";

type TrendRow = { date: string; count: number; ma7?: number | null; ma30?: number | null };
type TopRow   = { name: string; value: number };

const toCsv = (rows: any[]) => {
  if (!rows?.length) return "";
  const headers = Object.keys(rows[0]);
  const lines = [headers.join(","), ...rows.map(r => headers.map(h => JSON.stringify(r[h] ?? "")).join(","))];
  return "\uFEFF" + lines.join("\n");
};
const download = (name: string, blob: Blob) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = name; a.click();
  URL.revokeObjectURL(url);
};
const monthKey = (iso: string) => {
  const d = dayjs(iso); return d.isValid() ? d.format("YYYY-MM") : String(iso).slice(0, 7);
};
const movingAvg = (data: TrendRow[], key: keyof TrendRow, win: number) => {
  let sum = 0; const out: TrendRow[] = [];
  for (let i = 0; i < data.length; i++) {
    sum += Number(data[i][key] ?? 0);
    if (i >= win) sum -= Number(data[i - win][key] ?? 0);
    out.push({ ...data[i], [`ma${win}`]: i >= win - 1 ? +(sum / win).toFixed(2) : null });
  }
  return out;
};

export default function Dashboard() {
  const { kpi, trend, top5, loading } = useSheets();
  const [month, setMonth] = useState<string>("ALL");

  // 可選月份來源：trend 「日期」欄
  const months = useMemo(() => {
    const s = new Set<string>();
    for (const r of trend ?? []) {
      const d = r["日期"] ?? r["date"] ?? r["Date"];
      if (d) s.add(monthKey(String(d)));
    }
    return Array.from(s).sort();
  }, [trend]);

  const trendRows: TrendRow[] = useMemo(() => {
    const rows: TrendRow[] = (trend ?? [])
      .map(r => ({ date: String(r["日期"] ?? r["date"] ?? r["Date"]), count: Number(r["件數"] ?? r["count"] ?? 0) }))
      .filter(r => r.date)
      .sort((a, b) => a.date.localeCompare(b.date));
    const filtered = month === "ALL" ? rows : rows.filter(r => monthKey(r.date) === month);
    return movingAvg(movingAvg(filtered, "count", 7), "count", 30);
  }, [trend, month]);

  const topRows: TopRow[] = useMemo(() => {
    // 目前仍顯示彙整版 Top5；若要「按月 Top5」，改從 calls 即時計算
    return (top5 ?? []).map((r: any) => {
      const keys = Object.keys(r);
      return { name: String(r[keys[0]]), value: Number(r[keys[1]] ?? 0) };
    });
  }, [top5]);

  const k = kpi?.[0] ?? {};
  const trendRef = useRef<HTMLDivElement | null>(null);
  const topRef = useRef<HTMLDivElement | null>(null);

  const png = async (ref: React.RefObject<HTMLDivElement | null>, name: string) => {
    if (!ref.current) return;
    const canvas = await html2canvas(ref.current, { backgroundColor: "#fff", scale: 2 });
    canvas.toBlob(b => b && download(name, b));
  };

  if (loading) return <div className="page"><div className="loading">載入中…</div></div>;

  return (
    <div className="page">
      {/* 固定工具列 */}
      <div className="toolbar">
        <div className="title">IV&V / 客服 數據儀表板</div>
        <div className="tool-actions">
          <label className="label">月份</label>
          <select className="select" value={month} onChange={e => setMonth(e.target.value)}>
            <option value="ALL">全部</option>
            {months.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
          {/* 你可以在這裡加上「全部匯出 PNG/CSV」 */}
        </div>
      </div>

      {/* KPI 區：桌面 5 欄、平板 3 欄、手機 1 欄 */}
      <section className="kpi-grid">
        <Kpi label="本月總件數"       value={k["本月總件數"] ?? "-"} />
        <Kpi label="解決率"           value={k["解決率"] ?? "-"} />
        <Kpi label="平均處理時長(分)" value={k["平均處理時長"] ?? k["平均處理時長(分)"] ?? "-"} />
        <Kpi label="未結案數"         value={k["未結案數"] ?? "-"} />
        <Kpi label="SLA達成率"       value={k["SLA達成率"] ?? "-"} />
      </section>

      {/* 圖表網格：桌面兩欄 */}
      <section className="cards">
        {/* 日趨勢 */}
        <div className="card" ref={trendRef}>
          <div className="card-head">
            <div className="card-title">日趨勢（件數）{month !== "ALL" ? ` - ${month}` : ""}</div>
            <div className="actions">
              <button className="btn" onClick={() => png(trendRef, `trend-${month}.png`)}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`trend-${month}.csv`, new Blob([toCsv(trendRows)], { type: "text/csv;charset=utf-8" }))}
                disabled={!trendRows.length}
              >
                匯出 CSV
              </button>
            </div>
          </div>
          <div className="chart">
            <ResponsiveContainer>
              <LineChart data={trendRows}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="date" minTickGap={24} />
                <YAxis allowDecimals={false} />
                <Tooltip />
                <Legend />
                <Line type="monotone" dataKey="count" name="每日件數" dot={false} strokeWidth={2} />
                <Line type="monotone" dataKey="ma7"   name="MA7" dot={false} strokeWidth={1} strokeDasharray="5 3" />
                <Line type="monotone" dataKey="ma30"  name="MA30" dot={false} strokeWidth={1} strokeDasharray="2 4" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 模組 Top5 */}
        <div className="card" ref={topRef}>
          <div className="card-head">
            <div className="card-title">模組別 Top 5（件數）</div>
            <div className="actions">
              <button className="btn" onClick={() => png(topRef, `module-top5-${month}.png`)}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`module-top5-${month}.csv`, new Blob([toCsv(topRows)], { type: "text/csv;charset=utf-8" }))}
                disabled={!topRows.length}
              >
                匯出 CSV
              </button>
            </div>
          </div>
          <div className="chart">
            <ResponsiveContainer>
              <BarChart data={topRows}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" interval={0} angle={-10} textAnchor="end" height={58} />
                <YAxis allowDecimals={false} />
                <Tooltip />
                <Bar dataKey="value" name="件數" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      <footer className="footer">最後更新：{dayjs().format("YYYY-MM-DD HH:mm")}</footer>
    </div>
  );
}

function Kpi({ label, value }: { label: string; value: any }) {
  return (
    <div className="kpi">
      <div className="kpi-label">{label}</div>
      <div className="kpi-value">{String(value ?? "-")}</div>
    </div>
  );
}

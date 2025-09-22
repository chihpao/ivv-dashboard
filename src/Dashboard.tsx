import { useEffect, useMemo, useRef, useState, type ReactNode } from "react";
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

const numberFormatter = new Intl.NumberFormat("zh-Hant");

const formatNumber = (value: number | null | undefined) => {
  if (value == null || Number.isNaN(value)) return "-";
  return numberFormatter.format(value);
};

const formatPercent = (value: number | null | undefined, digits = 1) => {
  if (value == null || Number.isNaN(value)) return "-";
  const fixed = Number.isFinite(value) ? value.toFixed(digits) : value;
  const prefix = value > 0 ? "+" : "";
  return `${prefix}${fixed}%`;
};

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
const yearFromMonth = (month: string) => month.slice(0, 4);
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
  const { trend, moduleByMonth, avgCallDuration, loading } = useSheets();
  const [year, setYear] = useState<string>("ALL");
  const [month, setMonth] = useState<string>("ALL");
  const isYearAll = year === "ALL";

  // 趨勢資料（含移動平均）
  const trendAll: TrendRow[] = useMemo(() => {
    return (trend ?? [])
      .map(r => {
        const rawDate = r["日期"] ?? r["date"] ?? r["Date"];
        const date = dayjs(rawDate).isValid() ? dayjs(rawDate).format("YYYY-MM-DD") : String(rawDate ?? "");
        return {
          date,
          count: Number(r["件數"] ?? r["count"] ?? 0),
        };
      })
      .filter(r => r.date)
      .sort((a, b) => a.date.localeCompare(b.date));
  }, [trend]);

  const monthsByYear = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const row of trendAll) {
      const mk = monthKey(row.date);
      if (!mk || mk.length < 7) continue;
      const yr = yearFromMonth(mk);
      if (!map.has(yr)) map.set(yr, []);
      const list = map.get(yr)!;
      if (!list.includes(mk)) list.push(mk);
    }
    for (const list of map.values()) {
      list.sort();
    }
    return map;
  }, [trendAll]);

  const years = useMemo(() => {
    return Array.from(monthsByYear.keys()).sort();
  }, [monthsByYear]);

  const months = useMemo(() => {
    if (year === "ALL") return [];
    return monthsByYear.get(year) ?? [];
  }, [year, monthsByYear]);

  const avgDurationRows = useMemo(() => {
    const rows: Array<{ month: string; year: string; value: number }> = [];
    for (const r of avgCallDuration ?? []) {
      const rawMonth = (r as any)["month"] ?? (r as any)["月份"] ?? (r as any)["Month"];
      const mk = monthKey(String(rawMonth ?? ""));
      if (!mk || mk.length < 7) continue;
      const rawDuration =
        (r as any)["avg_call_duration"] ??
        (r as any)["平均處理時長"] ??
        (r as any)["平均處理時長(分)"] ??
        (r as any)["AvgCallDuration"] ??
        (r as any)["avg"];
      if (rawDuration === "" || rawDuration === undefined || rawDuration === null) continue;
      const duration = Number(rawDuration);
      if (!Number.isFinite(duration)) continue;
      rows.push({ month: mk, year: yearFromMonth(mk), value: duration });
    }
    return rows.sort((a, b) => a.month.localeCompare(b.month));
  }, [avgCallDuration]);

  useEffect(() => {
    if (year === "ALL") {
      if (month !== "ALL") setMonth("ALL");
      return;
    }
    const list = monthsByYear.get(year) ?? [];
    if (month !== "ALL" && !list.includes(month)) {
      setMonth("ALL");
    }
  }, [year, month, monthsByYear]);

  const filteredTrend = useMemo(() => {
    if (year === "ALL" && month === "ALL") return trendAll;
    return trendAll.filter(r => {
      const mk = monthKey(r.date);
      if (!mk || mk.length < 7) return false;
      if (year !== "ALL" && yearFromMonth(mk) !== year) return false;
      if (month !== "ALL" && mk !== month) return false;
      return true;
    });
  }, [trendAll, year, month]);

  const trendRows: TrendRow[] = useMemo(() => {
    return movingAvg(movingAvg(filteredTrend, "count", 7), "count", 30);
  }, [filteredTrend]);

  const monthlyTotals = useMemo(() => {
    const map = new Map<string, {
      month: string;
      total: number;
      days: number;
      max?: TrendRow;
      min?: TrendRow;
    }>();

    for (const row of trendAll) {
      const key = monthKey(row.date);
      if (!map.has(key)) {
        map.set(key, { month: key, total: 0, days: 0, max: row, min: row });
      }
      const stat = map.get(key)!;
      stat.total += row.count || 0;
      stat.days += 1;
      if (!stat.max || row.count > (stat.max.count || 0)) stat.max = row;
      if (!stat.min || row.count < (stat.min.count || 0)) stat.min = row;
    }

    return Array.from(map.values()).sort((a, b) => a.month.localeCompare(b.month));
  }, [trendAll]);

  const filteredMonthlyTotals = useMemo(() => {
    if (year === "ALL") return monthlyTotals;
    return monthlyTotals.filter(m => yearFromMonth(m.month) === year);
  }, [monthlyTotals, year]);

  const selectedMonthKey = useMemo(() => {
    if (month !== "ALL") return month;
    if (filteredMonthlyTotals.length) return filteredMonthlyTotals[filteredMonthlyTotals.length - 1].month;
    return monthlyTotals.length ? monthlyTotals[monthlyTotals.length - 1].month : null;
  }, [month, filteredMonthlyTotals, monthlyTotals]);

  const selectedMonthly = useMemo(() => {
    if (!selectedMonthKey) return null;
    return monthlyTotals.find(m => m.month === selectedMonthKey) ?? null;
  }, [monthlyTotals, selectedMonthKey]);

  const previousMonthly = useMemo(() => {
    if (!selectedMonthKey) return null;
    const baseList = year === "ALL" && month === "ALL" ? monthlyTotals : filteredMonthlyTotals;
    const idx = baseList.findIndex(m => m.month === selectedMonthKey);
    if (idx <= 0) return null;
    return baseList[idx - 1];
  }, [monthlyTotals, filteredMonthlyTotals, selectedMonthKey, year, month]);

  const yoyMonthly = useMemo(() => {
    if (!selectedMonthKey) return null;
    const base = dayjs(`${selectedMonthKey}-01`);
    if (!base.isValid()) return null;
    const yoyKey = base.subtract(1, "year").format("YYYY-MM");
    return monthlyTotals.find(m => m.month === yoyKey) ?? null;
  }, [monthlyTotals, selectedMonthKey]);

  // KPI：本月總件數（ALL = 全部月份總和）
  const monthTotalCount = useMemo(() => {
    return filteredTrend.reduce((sum, r) => sum + (r.count || 0), 0);
  }, [filteredTrend]);

  const averageDailyCount = useMemo(() => {
    if (!filteredTrend.length) return null;
    return +(monthTotalCount / filteredTrend.length).toFixed(1);
  }, [filteredTrend, monthTotalCount]);

  const averageDailyDetail = useMemo(() => {
    if (!filteredTrend.length) return null;
    if (month !== "ALL") {
      const d = dayjs(`${month}-01`);
      const prefix = d.isValid() ? d.format("YYYY 年 MM 月") : month;
      return `${prefix}平均（${filteredTrend.length} 日）`;
    }
    if (year !== "ALL") {
      return `${year} 年平均（${filteredTrend.length} 日）`;
    }
    return `全部年份平均（${filteredTrend.length} 日）`;
  }, [filteredTrend, month, year]);

  const avgDurationStat = useMemo(() => {
    if (!avgDurationRows.length) return { value: null as number | null, detail: null as string | null };

    const durationByMonth = new Map<string, number>();
    const monthsByYearFromDuration = new Map<string, string[]>();
    for (const item of avgDurationRows) {
      durationByMonth.set(item.month, item.value);
      if (!monthsByYearFromDuration.has(item.year)) monthsByYearFromDuration.set(item.year, []);
      monthsByYearFromDuration.get(item.year)!.push(item.month);
    }
    for (const months of monthsByYearFromDuration.values()) {
      months.sort();
    }

    const weightByMonth = new Map<string, number>();
    for (const m of monthlyTotals) {
      weightByMonth.set(m.month, Math.max(0, m.total || 0));
    }

    const computeAverage = (months: string[]) => {
      if (!months.length) return null;
      let weightedSum = 0;
      let weightSum = 0;
      let plainSum = 0;
      let count = 0;
      for (const mk of months) {
        const value = durationByMonth.get(mk);
        if (value == null) continue;
        const weight = weightByMonth.get(mk) ?? 0;
        if (weight > 0) {
          weightedSum += value * weight;
          weightSum += weight;
        }
        plainSum += value;
        count += 1;
      }
      if (!count) return null;
      if (weightSum > 0) {
        return { value: weightedSum / weightSum, months: count, weighted: true };
      }
      return { value: plainSum / count, months: count, weighted: false };
    };

    if (month !== "ALL") {
      const value = durationByMonth.get(month);
      if (value != null) {
        const d = dayjs(`${month}-01`);
        return {
          value,
          detail: d.isValid() ? d.format("YYYY 年 MM 月 平均") : `${month} 平均`,
        };
      }
    }

    if (year !== "ALL") {
      const months = monthsByYearFromDuration.get(year) ?? [];
      const result = computeAverage(months);
      if (result?.value != null) {
        const suffix = result.weighted ? "，依案件量加權" : "";
        return {
          value: result.value,
          detail: `${year} 年平均（${result.months} 月${suffix}）`,
        };
      }
    }

    const allMonths = avgDurationRows.map(item => item.month);
    const overall = computeAverage(allMonths);
    if (overall?.value != null) {
      const suffix = overall.weighted ? "，依案件量加權" : "";
      return {
        value: overall.value,
        detail: `全部年份平均（${overall.months} 月${suffix}）`,
      };
    }

    return { value: null, detail: null };
  }, [avgDurationRows, month, year, monthlyTotals]);

  const totalLabel = useMemo(() => {
    if (month !== "ALL") return "本月總件數";
    if (year !== "ALL") return `${year} 年總件數`;
    return "全部總件數";
  }, [year, month]);

  const rangeLabel = useMemo(() => {
    if (month !== "ALL") {
      const d = dayjs(`${month}-01`);
      return d.isValid() ? d.format("YYYY 年 MM 月") : month;
    }
    if (year !== "ALL") return `${year} 年`;
    return "";
  }, [year, month]);

  const exportKey = useMemo(() => {
    if (month !== "ALL") return month;
    if (year !== "ALL") return year;
    return "all";
  }, [year, month]);

  const dateTickFormatter = useMemo(() => {
    return (value: string) => {
      const d = dayjs(value);
      if (!d.isValid()) return value;
      return month !== "ALL" ? d.format("MM-DD") : d.format("YYYY-MM-DD");
    };
  }, [month]);

  const tooltipLabelFormatter = useMemo(() => {
    return (value: string) => {
      const d = dayjs(value);
      if (!d.isValid()) return value;
      return month !== "ALL" ? d.format("YYYY-MM-DD") : d.format("YYYY-MM-DD");
    };
  }, [month]);

  // 模組 Top5：使用 moduleByMonth（month, module, count）
  const topRows: TopRow[] = useMemo(() => {
    const normalized = (moduleByMonth ?? [])
      .map((r: any) => {
        const rawMonth = r["month"] ?? r["月份"] ?? r["Month"];
        const mk = monthKey(String(rawMonth ?? ""));
        if (!mk || mk.length < 7) return null;
        const module = String(r["module"] ?? r["模組"] ?? r["Module"] ?? "");
        if (!module) return null;
        const count = Number(r["count"] ?? r["件數"] ?? r["Count"] ?? 0);
        return { month: mk, year: yearFromMonth(mk), module, count };
      })
      .filter(Boolean) as Array<{ month: string; year: string; module: string; count: number }>;

    const filtered = normalized.filter(item => {
      if (year !== "ALL" && item.year !== year) return false;
      if (month !== "ALL" && item.month !== month) return false;
      return true;
    });

    const totals = new Map<string, number>();
    for (const r of filtered) {
      totals.set(r.module, (totals.get(r.module) ?? 0) + r.count);
    }

    return Array.from(totals.entries())
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 5);
  }, [moduleByMonth, year, month]);

  const trendRef = useRef<HTMLDivElement | null>(null);
  const topRef = useRef<HTMLDivElement | null>(null);
  const weekdayRef = useRef<HTMLDivElement | null>(null);

  const png = async (ref: React.RefObject<HTMLDivElement | null>, name: string) => {
    if (!ref.current) return;
    const canvas = await html2canvas(ref.current, { backgroundColor: "#fff", scale: 2 });
    canvas.toBlob(b => b && download(name, b));
  };

  const weekdayChartData = useMemo(() => {
    const order = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"];
    const labelByDay = ["週日", "週一", "週二", "週三", "週四", "週五", "週六"];
    const stats = new Map<string, { total: number; days: number }>();

    for (const row of filteredTrend) {
      const d = dayjs(row.date);
      if (!d.isValid()) continue;
      const label = labelByDay[d.day()];
      if (label === "週六") continue; // 週六固定 0
      if (!stats.has(label)) stats.set(label, { total: 0, days: 0 });
      const st = stats.get(label)!;
      st.total += row.count || 0;
      st.days += 1;
    }

    return order.map(name => {
      const st = stats.get(name);
      const average = name === "週六" ? 0 : (st && st.days ? st.total / st.days : 0);
      const total = name === "週六" ? 0 : (st?.total ?? 0);
      return { name, average: +average.toFixed(1), total };
    });
  }, [filteredTrend]);

  const insights = useMemo(() => {
    const list: string[] = [];
    if (!trendAll.length) return list;

    if (month !== "ALL" && selectedMonthly) {
      const { total, days } = selectedMonthly;
      const mom = previousMonthly && previousMonthly.total ? ((total - previousMonthly.total) / previousMonthly.total) * 100 : null;
      const yoy = yoyMonthly && yoyMonthly.total ? ((total - yoyMonthly.total) / yoyMonthly.total) * 100 : null;
      list.push(`本月共處理 ${formatNumber(total)} 件，平均每天 ${formatNumber(Math.round(total / Math.max(days, 1)))} 件`);
      if (mom != null) list.push(`相較上月 ${formatPercent(mom)}`);
      if (yoy != null) list.push(`相較去年同期 ${formatPercent(yoy)}`);
    } else if (year !== "ALL") {
      const avg = filteredTrend.length ? monthTotalCount / filteredTrend.length : 0;
      list.push(`${year} 年累積 ${formatNumber(monthTotalCount)} 件，平均每天 ${formatNumber(Math.round(avg))} 件`);
    } else if (selectedMonthly) {
      const { total, days } = selectedMonthly;
      list.push(`最近月份共處理 ${formatNumber(total)} 件，平均每天 ${formatNumber(Math.round(total / Math.max(days, 1)))} 件`);
    } else {
      list.push(`目前共有 ${formatNumber(trendAll.length)} 日的紀錄，累積 ${formatNumber(monthTotalCount)} 件`);
    }

    if (filteredTrend.length) {
      const peak = filteredTrend.reduce((acc, cur) => (cur.count > (acc?.count ?? -Infinity) ? cur : acc), null as TrendRow | null);
      const low = filteredTrend.reduce((acc, cur) => (cur.count < (acc?.count ?? Infinity) ? cur : acc), null as TrendRow | null);
      if (peak) list.push(`尖峰日 ${dayjs(peak.date).format("YYYY/MM/DD" )} 有 ${formatNumber(peak.count)} 件`);
      if (low && low !== peak) list.push(`最低進件日為 ${dayjs(low.date).format("YYYY/MM/DD")}，僅 ${formatNumber(low.count)} 件`);
    }

    if (topRows.length) {
      const [topModule] = topRows;
      list.push(`模組最多進件：${topModule.name}，${formatNumber(topModule.value)} 件`);
    }

    return list;
  }, [selectedMonthly, previousMonthly, yoyMonthly, filteredTrend, topRows, trendAll.length, monthTotalCount, year]);

  if (loading) return <LoadingState />;

  return (
    <div className="page">
      {/* 固定工具列 */}
      <div className="toolbar">
        <div className="title">IV&V / 客服 數據儀表板</div>
        <div className="filter-group">
          <div className="tool-actions">
            <label className="label">年份</label>
            <select className="select" value={year} onChange={e => setYear(e.target.value)}>
              <option value="ALL">全部年份</option>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <div className={`tool-actions${isYearAll ? " tool-actions--disabled" : ""}`}>
            <label className="label">月份</label>
            {isYearAll ? (
              <div className="select select--placeholder" aria-hidden="true"><span>--</span></div>
            ) : (
              <select
                className="select"
                value={month}
                onChange={e => setMonth(e.target.value)}
              >
                <option value="ALL">全年</option>
                {months.map(m => (
                  <option key={m} value={m}>{dayjs(`${m}-01`).isValid() ? dayjs(`${m}-01`).format("MM 月") : m}</option>
                ))}
              </select>
            )}
          </div>
        </div>
      </div>

      {/* KPI 區：桌面 5 欄、平板 3 欄、手機 1 欄 */}
      <section className="kpi-grid">
        <Kpi
          label={totalLabel}
          value={monthTotalCount}
          detail={rangeLabel || (filteredTrend.length ? `共 ${filteredTrend.length} 日` : null)}
        />
        <Kpi
          label="平均處理時長(分)"
          value={typeof avgDurationStat.value === "number" ? Number(avgDurationStat.value.toFixed(2)) : "-"}
          detail={avgDurationStat.detail}
        />
        <Kpi
          label="平均每日件數"
          value={averageDailyCount ?? "-"}
          detail={averageDailyDetail}
        />
      </section>

      {/* 圖表網格：桌面兩欄 */}
      <section className="cards">
        <div className="card card--insights">
          <div className="card-head">
            <div className="card-title">重點觀察</div>
          </div>
          <div className="insights">
            {insights.length ? (
              <ul className="insights-list">
                {insights.map((text, idx) => <li key={idx}>{text}</li>)}
              </ul>
            ) : (
              <div className="empty">尚無足夠資料</div>
            )}
          </div>
        </div>

        {/* 日趨勢 */}
        <div className="card" ref={trendRef}>
          <div className="card-head">
            <div className="card-title">日趨勢（件數）{rangeLabel ? ` - ${rangeLabel}` : ""}</div>
            <div className="actions">
              <button className="btn" onClick={() => png(trendRef, `trend-${exportKey}.png`)}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`trend-${exportKey}.csv`, new Blob([toCsv(trendRows)], { type: "text/csv;charset=utf-8" }))}
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
                <XAxis dataKey="date" minTickGap={24} tickFormatter={dateTickFormatter} />
                <YAxis allowDecimals={false} />
                <Tooltip labelFormatter={tooltipLabelFormatter} />
                <Legend />
                <Line type="monotone" dataKey="count" name="每日件數" dot={false} strokeWidth={2} />
                <Line type="monotone" dataKey="ma7"   name="MA7" dot={false} strokeWidth={1} strokeDasharray="5 3" />
                <Line type="monotone" dataKey="ma30"  name="MA30" dot={false} strokeWidth={1} strokeDasharray="2 4" />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 週別節奏 */}
        <div className="card" ref={weekdayRef}>
          <div className="card-head">
            <div className="card-title">週期節奏（平均每日件數）</div>
            <div className="actions">
              <button className="btn" onClick={() => png(weekdayRef, `weekday-pattern-${exportKey}.png`)} disabled={!weekdayChartData.length}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`weekday-pattern-${exportKey}.csv`, new Blob([toCsv(weekdayChartData)], { type: "text/csv;charset=utf-8" }))}
                disabled={!weekdayChartData.length}
              >
                匯出 CSV
              </button>
            </div>
          </div>
          <div className="chart">
            <ResponsiveContainer>
              <BarChart data={weekdayChartData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="name" />
                <YAxis allowDecimals />
                <Tooltip formatter={(value: any) => [`${Number(value).toFixed(1)} 件`, "平均每日"]} />
                <Bar dataKey="average" name="平均每日件數" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 模組 Top5（會跟月份變動） */}
        <div className="card" ref={topRef}>
          <div className="card-head">
            <div className="card-title">模組別 Top 5（件數）{rangeLabel ? ` - ${rangeLabel}` : ""}</div>
            <div className="actions">
              <button className="btn" onClick={() => png(topRef, `module-top5-${exportKey}.png`)}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`module-top5-${exportKey}.csv`, new Blob([toCsv(topRows)], { type: "text/csv;charset=utf-8" }))}
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

function Kpi({ label, value, detail }: { label: string; value: any; detail?: ReactNode | null }) {
  const display = typeof value === "number" && !Number.isNaN(value) ? formatNumber(value) : String(value ?? "-");
  return (
    <div className="kpi">
      <div className="kpi-label">{label}</div>
      <div className="kpi-value">{display}</div>
      {detail ? <div className="kpi-detail">{detail}</div> : null}
    </div>
  );
}

function LoadingState() {
  return (
    <div className="page loading-screen">
      <div className="loading-shell" role="status" aria-live="polite">
        <div className="loading-hero" aria-hidden="true">
          <div className="loading-hero__orb" />
          <div className="loading-hero__wave loading-hero__wave--one" />
          <div className="loading-hero__wave loading-hero__wave--two" />
          <div className="loading-hero__spark loading-hero__spark--one" />
          <div className="loading-hero__spark loading-hero__spark--two" />
          <div className="loading-hero__spark loading-hero__spark--three" />
        </div>
        <div className="loading-copy">
          <div className="loading-heading">
            <span className="loading-heading__main">資料載入中</span>
          </div>
          <p className="loading-message">Loading...</p>
        </div>
        <div className="loading-progress">
          <span className="loading-progress__bar" />
        </div>
      </div>
    </div>
  );
}

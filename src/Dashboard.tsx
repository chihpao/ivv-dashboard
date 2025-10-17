import { useEffect, useMemo, useRef, useState, type ReactNode, type CSSProperties } from "react";
import { useSheets } from "./useSheets";
import dayjs from "dayjs";
import html2canvas from "html2canvas";
import {
  LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, ResponsiveContainer,
  BarChart, Bar, Legend, AreaChart, Area, ReferenceLine
} from "recharts";
import "./Dashboard.css";

type TrendRow = { date: string; count: number; ma7?: number | null; ma30?: number | null };
type TopRow   = { name: string; value: number };
type DurationMetric = "count" | "percentage";
type DurationBinMode = "auto" | "fixed";
type DurationGroupBy = "none" | "category" | "module";

type DurationDrawerState =
  | { type: "module"; name: string }
  | {
      type: "duration";
      label: string;
      groupLabel: string | null;
      min: number;
      max: number | null;
      rows: Array<Record<string, any>>;
    };

type DurationRow = {
  minutes: number;
  category: string;
  module: string;
  source: Record<string, any>;
};

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

const CATEGORY_COLOR_PALETTE = ["#0ea5e9", "#22c55e", "#f97316", "#8b5cf6", "#f43f5e", "#14b8a6"];
const DURATION_OTHERS_KEY = "其他";

const FIXED_DURATION_BINS = [
  { min: 0, max: 3 },
  { min: 3, max: 5 },
  { min: 5, max: 10 },
  { min: 10, max: 15 },
  { min: 15, max: 20 },
  { min: 20, max: 30 },
  { min: 30, max: 45 },
  { min: 45, max: 60 },
  { min: 60, max: 90 },
  { min: 90, max: 120 },
  { min: 120, max: 180 },
  { min: 180, max: null },
];

const formatDurationLabel = (min: number, max: number | null) => {
  if (max == null) return `${Math.round(min)}+ 分`;
  const roundedMin = Math.round(min);
  const roundedMax = Math.round(max);
  if (roundedMin === roundedMax) return `${roundedMin} 分`;
  return `${roundedMin}-${roundedMax} 分`;
};

export default function Dashboard() {
  const { trend, moduleByMonth, avgCallDuration, calls, loading } = useSheets();
  const [drawerState, setDrawerState] = useState<DurationDrawerState | null>(null);
  const [year, setYear] = useState<string>("ALL");
  const [month, setMonth] = useState<string>("ALL");
  const [durationMetric, setDurationMetric] = useState<DurationMetric>("count");
  const [durationBinMode, setDurationBinMode] = useState<DurationBinMode>("auto");
  const [durationFocus30, setDurationFocus30] = useState<boolean>(false);
  const [durationGroupBy, setDurationGroupBy] = useState<DurationGroupBy>("none");
  const [durationGroupSelection, setDurationGroupSelection] = useState<string>("ALL");
  const [durationFacetEnabled, setDurationFacetEnabled] = useState<boolean>(false);
  const isYearAll = year === "ALL";
  const toggleActiveStyle: CSSProperties = {
    fontWeight: 600,
    borderColor: "#2563eb",
    color: "#2563eb",
    backgroundColor: "rgba(37, 99, 235, 0.08)",
  };

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

  const filteredCalls = useMemo(() => {
    if (!calls?.length) return [] as Array<Record<string, any>>;
    return (calls as Array<Record<string, any>>).filter(row => {
      const mk = monthKey(String(row["call_month"] ?? row["call_time"] ?? ""));
      if (!mk || mk.length < 7) return false;
      if (year !== "ALL" && yearFromMonth(mk) !== year) return false;
      if (month !== "ALL" && mk !== month) return false;
      return true;
    });
  }, [calls, year, month]);

  const durationBase = useMemo(() => {
    const rows: DurationRow[] = [];
    const missing: Array<Record<string, any>> = [];
    const categoryTotals = new Map<string, number>();
    const moduleTotals = new Map<string, number>();

    const normalizeCategory = (row: Record<string, any>) => {
      const value =
        row["category"] ??
        row["分類"] ??
        row["Category"] ??
        row["呼叫分類"] ??
        row["類別"];
      const name = String(value ?? "").trim();
      return name || "未分類";
    };

    const normalizeModule = (row: Record<string, any>) => {
      const value =
        row["module"] ??
        row["模組"] ??
        row["Module"] ??
        row["呼叫模組"] ??
        row["系統"];
      const name = String(value ?? "").trim();
      return name || "未指派";
    };

    for (const rawRow of filteredCalls) {
      const rawDuration = Number(rawRow["resolve_minute"] ?? rawRow["resolve_minutes"]);
      const category = normalizeCategory(rawRow);
      const moduleName = normalizeModule(rawRow);
      const sourceRow = rawRow as Record<string, any>;

      categoryTotals.set(category, (categoryTotals.get(category) ?? 0) + 1);
      moduleTotals.set(moduleName, (moduleTotals.get(moduleName) ?? 0) + 1);

      if (!Number.isFinite(rawDuration)) {
        missing.push(sourceRow);
        continue;
      }

      const minutes = Math.max(0, rawDuration);
      rows.push({ minutes, category, module: moduleName, source: sourceRow });
    }

    return { rows, missing, categoryTotals, moduleTotals };
  }, [filteredCalls]);

  const durationCategoryOptions = useMemo(() => {
    return Array.from(durationBase.categoryTotals.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([name]) => name);
  }, [durationBase.categoryTotals]);

  const durationModuleOptions = useMemo(() => {
    return Array.from(durationBase.moduleTotals.entries())
      .sort((a, b) => b[1] - a[1])
      .map(([name]) => name);
  }, [durationBase.moduleTotals]);

  useEffect(() => {
    if (durationGroupBy === "none") {
      if (durationGroupSelection !== "ALL") setDurationGroupSelection("ALL");
      if (durationFacetEnabled) setDurationFacetEnabled(false);
      return;
    }
    const options = durationGroupBy === "category" ? durationCategoryOptions : durationModuleOptions;
    if (durationGroupSelection !== "ALL" && !options.includes(durationGroupSelection)) {
      setDurationGroupSelection("ALL");
    }
    if (durationGroupSelection !== "ALL" && durationFacetEnabled) {
      setDurationFacetEnabled(false);
    }
  }, [
    durationGroupBy,
    durationCategoryOptions,
    durationModuleOptions,
    durationGroupSelection,
    durationFacetEnabled,
  ]);

  const durationFilterOptions = durationGroupBy === "category"
    ? durationCategoryOptions
    : durationGroupBy === "module"
    ? durationModuleOptions
    : [];
  const canFacet = durationGroupBy !== "none" && durationGroupSelection === "ALL";

  const {
    data: categoryStackData,
    keys: categoryStackKeys,
    isDaily: categoryStackIsDaily,
  } = useMemo(() => {
    if (!filteredCalls.length) {
      return {
        data: [] as Array<Record<string, number | string>>,
        keys: [] as string[],
        isDaily: month !== "ALL",
      };
    }

    const useDaily = month !== "ALL";
    const aggregate = new Map<string, Map<string, number>>();
    const totals = new Map<string, number>();

    for (const row of filteredCalls) {
      const rawMonth = row["call_month"] ?? row["month"] ?? row["call_time"];
      const mk = monthKey(String(rawMonth ?? ""));
      if (!mk || mk.length < 7) continue;
      const callDate = dayjs(row["call_time"]);
      const bucketKey = useDaily && callDate.isValid() ? callDate.format("YYYY-MM-DD") : mk;
      const category = String(row["category"] ?? row["分類"] ?? "").trim() || "未分類";

      if (!aggregate.has(bucketKey)) aggregate.set(bucketKey, new Map());
      const bucket = aggregate.get(bucketKey)!;
      bucket.set(category, (bucket.get(category) ?? 0) + 1);

      totals.set(category, (totals.get(category) ?? 0) + 1);
    }

    if (!aggregate.size) {
      return {
        data: [] as Array<Record<string, number | string>>,
        keys: [] as string[],
        isDaily: useDaily,
      };
    }

    const topCategories = Array.from(totals.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([key]) => key);

    const sortedKeys = Array.from(aggregate.keys()).sort((a, b) => a.localeCompare(b));

    const dataRows = sortedKeys.map(periodKey => {
      const bucket = aggregate.get(periodKey)!;
      const row: Record<string, number | string> = { period: periodKey };

      for (const [category, count] of bucket.entries()) {
        if (topCategories.includes(category)) {
          row[category] = count;
        }
      }

      return row;
    });

    const keys = [...topCategories];

    return { data: dataRows, keys, isDaily: useDaily };
  }, [filteredCalls, month]);

  const categoryColorMap = useMemo(() => {
    const map = new Map<string, string>();
    categoryStackKeys.forEach((key, idx) => {
      const color = CATEGORY_COLOR_PALETTE[idx % CATEGORY_COLOR_PALETTE.length];
      map.set(key, color);
    });
    return map;
  }, [categoryStackKeys]);

  const categoryXAxisFormatter = useMemo(() => {
    return (value: string) => {
      if (categoryStackIsDaily) {
        const d = dayjs(value);
        return d.isValid() ? d.format("MM-DD") : value;
      }
      const d = dayjs(`${value}-01`);
      return d.isValid() ? d.format("YYYY-MM") : value;
    };
  }, [categoryStackIsDaily]);

  const categoryTooltipLabelFormatter = useMemo(() => {
    return (value: string) => {
      if (categoryStackIsDaily) {
        const d = dayjs(value);
        return d.isValid() ? d.format("YYYY-MM-DD") : value;
      }
      const d = dayjs(`${value}-01`);
      return d.isValid() ? d.format("YYYY-MM") : value;
    };
  }, [categoryStackIsDaily]);

  const durationChart = useMemo(() => {
    const { rows: allRows, missing } = durationBase;
    const groupLabel = durationGroupBy === "category" ? "分類" : durationGroupBy === "module" ? "模組" : null;

    if (!allRows.length) {
      return {
        rows: [] as Array<Record<string, any>>,
        csvRows: [] as Array<Record<string, any>>,
        seriesKeys: [] as string[],
        colorMap: new Map<string, string>(),
        bucketDetails: new Map<string, { label: string; groupLabel: string | null; min: number; max: number | null; rows: Array<Record<string, any>> }>(),
        totalCount: 0,
        mean: null as number | null,
        median: null as number | null,
        meanLabel: null as string | null,
        medianLabel: null as string | null,
        missingCount: missing.length,
        missingRows: missing,
        groupLabel,
        facet: false,
        groupDisplay: new Map<string, string>(),
      };
    }

    const focusLimit = durationFocus30 ? 30 : null;
    const usingCategory = durationGroupBy === "category";
    const usingModule = durationGroupBy === "module";
    const groupField = usingCategory ? "category" : usingModule ? "module" : null;

    let activeRows = allRows;
    if (usingCategory && durationGroupSelection !== "ALL") {
      activeRows = allRows.filter(row => row.category === durationGroupSelection);
    } else if (usingModule && durationGroupSelection !== "ALL") {
      activeRows = allRows.filter(row => row.module === durationGroupSelection);
    }

    const selectedDisplayName =
      usingCategory && durationGroupSelection !== "ALL"
        ? durationGroupSelection
        : usingModule && durationGroupSelection !== "ALL"
        ? durationGroupSelection
        : "全部";

    if (!activeRows.length) {
      return {
        rows: [] as Array<Record<string, any>>,
        csvRows: [] as Array<Record<string, any>>,
        seriesKeys: ["__all"],
        colorMap: new Map<string, string>([["__all", CATEGORY_COLOR_PALETTE[0]]]),
        bucketDetails: new Map<string, { label: string; groupLabel: string | null; min: number; max: number | null; rows: Array<Record<string, any>> }>(),
        totalCount: 0,
        mean: null as number | null,
        median: null as number | null,
        meanLabel: null as string | null,
        medianLabel: null as string | null,
        missingCount: missing.length,
        missingRows: missing,
        groupLabel,
        facet: false,
        groupDisplay: new Map<string, string>([["__all", selectedDisplayName]]),
      };
    }

    const valuesSorted = [...activeRows].map(row => row.minutes).sort((a, b) => a - b);
    const totalCount = valuesSorted.length;
    const mean = valuesSorted.reduce((sum, value) => sum + value, 0) / totalCount;
    const median =
      totalCount % 2 === 0
        ? (valuesSorted[totalCount / 2 - 1] + valuesSorted[totalCount / 2]) / 2
        : valuesSorted[Math.floor(totalCount / 2)];

    const valuesForBins = focusLimit == null
      ? activeRows.map(row => row.minutes)
      : activeRows.filter(row => row.minutes <= focusLimit).map(row => row.minutes);
    const overLimitRows = focusLimit == null
      ? [] as DurationRow[]
      : activeRows.filter(row => row.minutes > focusLimit);

    const buildFixedBins = (limit: number | null) => {
      let bins = FIXED_DURATION_BINS.map(bin => ({
        min: bin.min,
        max: bin.max,
        label: formatDurationLabel(bin.min, bin.max),
      }));
      if (limit != null) {
        bins = bins.filter(bin => (bin.max ?? Infinity) <= limit + 1e-6);
      }
      if (!bins.length) {
        const max = limit != null ? limit : valuesSorted[valuesSorted.length - 1];
        const safeMax = Number.isFinite(max) ? max : 30;
        bins = [{ min: 0, max: safeMax, label: formatDurationLabel(0, safeMax) }];
      }
      return bins;
    };

    const buildAutoBins = (limit: number | null) => {
      if (!valuesForBins.length) {
        const fallback = limit != null ? limit : valuesSorted[valuesSorted.length - 1];
        const safeMax = Number.isFinite(fallback) ? fallback : 30;
        return [{ min: 0, max: safeMax, label: formatDurationLabel(0, safeMax) }];
      }
      const minValue = Math.min(...valuesForBins);
      const maxValue = Math.max(...valuesForBins);
      if (minValue === maxValue) {
        const width = Math.max(1, minValue);
        const lower = Math.max(0, minValue - width / 2);
        return [{ min: lower, max: minValue + width / 2, label: formatDurationLabel(lower, minValue + width / 2) }];
      }
      const span = Math.max(maxValue - minValue, 1);
      const desiredBins = Math.min(12, Math.max(4, Math.ceil(Math.sqrt(valuesForBins.length))));
      const width = span / desiredBins;
      const bins: Array<{ min: number; max: number; label: string }> = [];
      let start = minValue;
      for (let i = 0; i < desiredBins; i += 1) {
        const end = i === desiredBins - 1 ? maxValue : start + width;
        bins.push({ min: start, max: end, label: formatDurationLabel(start, end) });
        start = end;
      }
      return bins;
    };

    const baseBins = durationBinMode === "fixed"
      ? buildFixedBins(focusLimit)
      : buildAutoBins(focusLimit);

    const binSummaries = baseBins.map((bin, index) => ({
      ...bin,
      key: "bin-" + index,
      counts: new Map<string, number>(),
      total: 0,
    }));

    let overLimitIndex: number | null = null;
    if (focusLimit != null && overLimitRows.length) {
      overLimitIndex = binSummaries.length;
      binSummaries.push({
        min: focusLimit,
        max: null,
        label: focusLimit + "+ 分",
        key: "bin-" + binSummaries.length,
        counts: new Map<string, number>(),
        total: 0,
      });
    }

    const getRowGroup = (row: DurationRow) => {
      if (groupField === "category") return row.category;
      if (groupField === "module") return row.module;
      return "__all";
    };

    const facet = durationFacetEnabled && durationGroupBy !== "none" && durationGroupSelection === "ALL" && groupField !== null;
    const groupTotals = new Map<string, number>();
    const colorMap = new Map<string, string>();
    const groupDisplay = new Map<string, string>();

    let groupKeys: string[] = [];
    let topGroupSet: Set<string> | null = null;

    if (facet) {
      const totalsByGroup = new Map<string, number>();
      for (const row of activeRows) {
        const key = getRowGroup(row);
        totalsByGroup.set(key, (totalsByGroup.get(key) ?? 0) + 1);
      }
      const sortedGroups = Array.from(totalsByGroup.entries()).sort((a, b) => b[1] - a[1]);
      const topGroups = sortedGroups.slice(0, 5).map(([key]) => key);
      groupKeys = [...topGroups];
      topGroupSet = new Set(topGroups);
      const othersCount = sortedGroups.slice(5).reduce((sum, [, value]) => sum + value, 0);
      if (othersCount > 0) {
        groupKeys.push(DURATION_OTHERS_KEY);
      }
      for (const key of groupKeys) {
        groupDisplay.set(key, key === DURATION_OTHERS_KEY ? DURATION_OTHERS_KEY : key);
      }
    } else {
      groupKeys = ["__all"];
      groupDisplay.set("__all", selectedDisplayName);
    }

    groupKeys.forEach((key, index) => {
      colorMap.set(key, CATEGORY_COLOR_PALETTE[index % CATEGORY_COLOR_PALETTE.length]);
      groupTotals.set(key, 0);
    });

    const bucketDetails = new Map<string, { label: string; groupLabel: string | null; min: number; max: number | null; rows: Array<Record<string, any>> }>();

    const getGroupKey = (row: DurationRow) => {
      if (!facet) return "__all";
      const value = getRowGroup(row);
      if (topGroupSet && topGroupSet.has(value)) return value;
      return DURATION_OTHERS_KEY;
    };

    const assignRowToBinIndex = (value: number) => {
      if (overLimitIndex != null && focusLimit != null && value > focusLimit) {
        return overLimitIndex;
      }
      for (let i = 0; i < baseBins.length; i += 1) {
        const bin = baseBins[i];
        const max = bin.max ?? Infinity;
        const inclusive = i === baseBins.length - 1;
        const upperBound = inclusive ? max + 1e-6 : max;
        if (value >= bin.min && value < upperBound) {
          return i;
        }
      }
      return baseBins.length - 1;
    };

    for (const row of activeRows) {
      const groupKey = getGroupKey(row);
      const binIndex = assignRowToBinIndex(row.minutes);
      const summary = binSummaries[binIndex];
      summary.counts.set(groupKey, (summary.counts.get(groupKey) ?? 0) + 1);
      summary.total += 1;
      groupTotals.set(groupKey, (groupTotals.get(groupKey) ?? 0) + 1);

      const detailKey = binIndex + "|" + groupKey;
      if (!bucketDetails.has(detailKey)) {
        bucketDetails.set(detailKey, {
          label: summary.label,
          groupLabel: groupKey === "__all" ? null : (groupDisplay.get(groupKey) ?? groupKey),
          min: summary.min,
          max: summary.max ?? null,
          rows: [],
        });
      }
      bucketDetails.get(detailKey)!.rows.push(row.source);
    }

    const overallTotal = activeRows.length;
    const rows: Array<Record<string, any>> = [];
    binSummaries.forEach((bin, index) => {
      if (bin.total === 0 && bin.max !== null && durationBinMode !== "fixed") {
        return;
      }
      const record: Record<string, any> = {
        label: bin.label,
        bucketKey: String(index),
        min: bin.min,
        max: bin.max ?? null,
        total: bin.total,
      };
      const countsRecord: Record<string, number> = {};
      const percentRecord: Record<string, number> = {};
      for (const key of groupKeys) {
        const count = bin.counts.get(key) ?? 0;
        countsRecord[key] = count;
        const denominator = facet ? (groupTotals.get(key) ?? 0) || 1 : overallTotal || 1;
        const percent = denominator ? (count / denominator) * 100 : 0;
        percentRecord[key] = percent;
        record[key] = durationMetric === "count" ? count : Number(percent.toFixed(1));
      }
      record.__counts = countsRecord;
      record.__percents = percentRecord;
      rows.push(record);
    });

    const csvRows = rows.map(row => {
      const base: Record<string, any> = {
        區間: row.label,
        總件數: row.total,
      };
      const counts = row.__counts as Record<string, number>;
      const percents = row.__percents as Record<string, number>;
      for (const key of groupKeys) {
        const display = groupDisplay.get(key) ?? key;
        base[display + "-件數"] = counts[key] ?? 0;
        base[display + "-佔比(%)"] = Number((percents[key] ?? 0).toFixed(1));
      }
      return base;
    });

    if (missing.length) {
      csvRows.push({
        區間: "未填寫",
        總件數: missing.length,
        "全部-件數": missing.length,
        "全部-佔比(%)": Number(((missing.length / (overallTotal + missing.length)) * 100).toFixed(1)),
      });
    }

    const findLabelForValue = (value: number | null) => {
      if (value == null) return null;
      if (focusLimit != null && value > focusLimit && overLimitIndex != null) {
        return focusLimit + "+ 分";
      }
      for (const bin of binSummaries) {
        const max = bin.max ?? Infinity;
        if (value >= bin.min && value <= max + 1e-6) {
          return bin.label;
        }
      }
      return null;
    };

    const meanLabel = findLabelForValue(mean);
    const medianLabel = findLabelForValue(median);

    return {
      rows,
      csvRows,
      seriesKeys: groupKeys,
      colorMap,
      bucketDetails,
      totalCount,
      mean,
      median,
      meanLabel,
      medianLabel,
      missingCount: missing.length,
      missingRows: missing,
      groupLabel,
      facet,
      groupDisplay,
    };
  }, [
    durationBase,
    durationBinMode,
    durationFacetEnabled,
    durationFocus30,
    durationGroupBy,
    durationGroupSelection,
    durationMetric,
  ]);

  const handleDurationBarClick = (entry: any, seriesKey: string) => {
    const payload = entry?.payload;
    if (!payload) return;
    const bucketKey = String(payload.bucketKey ?? "");
    const detailKey = bucketKey + "|" + seriesKey;
    const detail = durationChart.bucketDetails.get(detailKey);
    if (!detail || !detail.rows.length) return;
    setDrawerState({
      type: "duration",
      label: detail.label,
      groupLabel: detail.groupLabel,
      min: detail.min,
      max: detail.max,
      rows: detail.rows,
    });
  };

  const durationMeanDisplay = durationChart.mean != null ? durationChart.mean.toFixed(1) : null;
  const durationMedianDisplay = durationChart.median != null ? durationChart.median.toFixed(1) : null;

  const trendRef = useRef<HTMLDivElement | null>(null);
  const categoryRef = useRef<HTMLDivElement | null>(null);
  const weekdayRef = useRef<HTMLDivElement | null>(null);
  const durationRef = useRef<HTMLDivElement | null>(null);
  const topRef = useRef<HTMLDivElement | null>(null);

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

  const moduleDrilldownRows = useMemo(() => {
    if (drawerState?.type !== "module") return [];
    return filteredCalls.filter((row: any) => {
      const moduleName = String(row["module"] ?? row["模組"] ?? row["Module"] ?? "");
      return moduleName === drawerState.name;
    });
  }, [filteredCalls, drawerState]);

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
      {/* 圖表網格：桌面兩欄 */}

      {/* 圖表網格：桌面兩欄 */}

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
        {/* 日趨勢（件數） */}
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
        {/* 分類堆疊面積圖 */}
        <div className="card" ref={categoryRef}>
          <div className="card-head">
            <div className="card-title">分類堆疊面積圖{rangeLabel ? ` - ${rangeLabel}` : ""}</div>
            <div className="actions">
              <button className="btn" onClick={() => png(categoryRef, `category-stack-${exportKey}.png`)} disabled={!categoryStackData.length}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`category-stack-${exportKey}.csv`, new Blob([toCsv(categoryStackData)], { type: "text/csv;charset=utf-8" }))}
                disabled={!categoryStackData.length}
              >
                匯出 CSV
              </button>
            </div>
          </div>
          <div className="chart">
            {categoryStackData.length && categoryStackKeys.length ? (
              <ResponsiveContainer>
                <AreaChart data={categoryStackData}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="period" tickFormatter={categoryXAxisFormatter} />
                  <YAxis allowDecimals={false} />
                  <Tooltip
                    labelFormatter={categoryTooltipLabelFormatter}
                    formatter={(value: any, name) => [formatNumber(Number(value) || 0), name]}
                  />
                  <Legend />
                  {categoryStackKeys.map(key => {
                    const color = categoryColorMap.get(key) ?? CATEGORY_COLOR_PALETTE[0];
                    return (
                      <Area
                        key={key}
                        type="monotone"
                        dataKey={key}
                        name={key}
                        stackId="categories"
                        stroke={color}
                        fill={color}
                        fillOpacity={0.75}
                        dot={false}
                      />
                    );
                  })}
                </AreaChart>
              </ResponsiveContainer>
            ) : (
              <div className="empty">尚無足夠資料</div>
            )}
          </div>
        </div>
        {/* 週期節奏（平均每日件數） */}
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
        <div className="card" ref={durationRef}>
          <div className="card-head">
            <div className="card-title">處理時間分佈（分鐘）{rangeLabel ? ` - ${rangeLabel}` : ""}</div>
            <div className="actions">
              <button className="btn" onClick={() => png(durationRef, `resolve-distribution-${exportKey}.png`)} disabled={!durationChart.rows.length}>匯出 PNG</button>
              <button
                className="btn"
                onClick={() => download(`resolve-distribution-${exportKey}.csv`, new Blob([toCsv(durationChart.csvRows)], { type: "text/csv;charset=utf-8" }))}
                disabled={!durationChart.rows.length}
              >
                匯出 CSV
              </button>
            </div>
          </div>
          <div className="card-toolbar" style={{ display: "flex", flexWrap: "wrap", gap: "10px", marginBottom: "12px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <span>顯示</span>
              <button
                className="btn"
                style={durationMetric === "count" ? toggleActiveStyle : undefined}
                onClick={() => setDurationMetric("count")}
              >
                件數
              </button>
              <button
                className="btn"
                style={durationMetric === "percentage" ? toggleActiveStyle : undefined}
                onClick={() => setDurationMetric("percentage")}
              >
                佔比
              </button>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <span>分箱</span>
              <button
                className="btn"
                style={durationBinMode === "auto" ? toggleActiveStyle : undefined}
                onClick={() => setDurationBinMode("auto")}
              >
                自動
              </button>
              <button
                className="btn"
                style={durationBinMode === "fixed" ? toggleActiveStyle : undefined}
                onClick={() => setDurationBinMode("fixed")}
              >
                固定
              </button>
            </div>
            <label style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <input
                type="checkbox"
                checked={durationFocus30}
                onChange={(e) => setDurationFocus30(e.target.checked)}
              />
              聚焦 0-30 分
            </label>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <span>篩選</span>
              <select
                className="select"
                value={durationGroupBy}
                onChange={(e) => setDurationGroupBy(e.target.value as DurationGroupBy)}
              >
                <option value="none">全部</option>
                <option value="category">分類</option>
                <option value="module">模組</option>
              </select>
              {durationGroupBy !== "none" && (
                <select
                  className="select"
                  value={durationGroupSelection}
                  onChange={(e) => setDurationGroupSelection(e.target.value)}
                >
                  <option value="ALL">全部</option>
                  {durationFilterOptions.map(option => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              )}
            </div>
            {durationGroupBy !== "none" && (
              <label style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <input
                  type="checkbox"
                  checked={durationFacetEnabled && canFacet}
                  onChange={(e) => setDurationFacetEnabled(e.target.checked && canFacet)}
                  disabled={!canFacet}
                />
                Facet 比較
              </label>
            )}
          </div>
          <div className="chart">
            {durationChart.rows.length ? (
              <ResponsiveContainer>
                <BarChart data={durationChart.rows}>
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis dataKey="label" />
                  <YAxis
                    allowDecimals={durationMetric === "percentage"}
                    tickFormatter={(value) => durationMetric === "percentage" ? `${value}%` : formatNumber(Number(value) || 0)}
                  />
                  <Tooltip
                    labelFormatter={(label) => label}
                    formatter={(_value: any, dataKey: any, props: any) => {
                      const key = String(dataKey);
                      const payload = props?.payload ?? {};
                      const counts = (payload.__counts as Record<string, number> | undefined) ?? {};
                      const percents = (payload.__percents as Record<string, number> | undefined) ?? {};
                      const count = counts[key] ?? 0;
                      const percent = percents[key] ?? 0;
                      const percentDisplay = Number.isFinite(percent) ? Number(percent.toFixed(1)) : 0;
                      const displayName = durationChart.groupDisplay.get(key) ?? key;
                      if (durationMetric === "count") {
                        return [formatNumber(count), displayName];
                      }
                      return [`${percentDisplay}%`, `${displayName}（${formatNumber(count)} 件）`];
                    }}
                  />
                  <Legend />
                  {durationChart.medianLabel && durationMedianDisplay && (
                    <ReferenceLine
                      x={durationChart.medianLabel}
                      stroke="#ef4444"
                      strokeDasharray="6 6"
                      label={{ value: `中位數 ${durationMedianDisplay} 分`, position: "insideTop", fill: "#ef4444", fontSize: 12 }}
                    />
                  )}
                  {durationChart.meanLabel && durationMeanDisplay && (
                    <ReferenceLine
                      x={durationChart.meanLabel}
                      stroke="#0ea5e9"
                      strokeDasharray="4 4"
                      label={{ value: `平均 ${durationMeanDisplay} 分`, position: "insideBottom", fill: "#0ea5e9", fontSize: 12 }}
                    />
                  )}
                  {durationChart.seriesKeys.map(key => (
                    <Bar
                      key={key}
                      dataKey={key}
                      name={durationChart.groupDisplay.get(key) ?? key}
                      fill={durationChart.colorMap.get(key) ?? CATEGORY_COLOR_PALETTE[0]}
                      onClick={(data) => handleDurationBarClick(data, key)}
                      cursor="pointer"
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            ) : (
              <div className="empty">暫無資料</div>
            )}
          </div>
          {durationChart.missingCount > 0 && (
            <div style={{ marginTop: "8px", fontSize: "12px", color: "#6b7280" }}>
              缺少處理時間資料：{formatNumber(durationChart.missingCount)} 筆
            </div>
          )}
        </div>
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
                <Bar
                  dataKey="value"
                  name="案件數"
                  onClick={(data) => {
                    const moduleName = String(data?.name || data?.payload?.name || "");
                    if (!moduleName) return;
                    setDrawerState({ type: "module", name: moduleName });
                  }}
                  style={{ cursor: "pointer" }}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      <footer className="footer">最後更新：{dayjs().format("YYYY-MM-DD HH:mm")}</footer>
      {drawerState && (
        <>
          <div
            className="drawer-overlay active"
            onClick={() => setDrawerState(null)}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => e.key === 'Escape' && setDrawerState(null)}
          />
          <div
            className="drawer open"
            role="dialog"
            aria-modal="true"
            aria-labelledby="drawer-title"
          >
            <div className="drawer-header">
              <h3 id="drawer-title">
                {drawerState.type === "module"
                  ? `${drawerState.name} - 案件明細`
                  : `${drawerState.label}${drawerState.groupLabel ? ` / ${drawerState.groupLabel}` : ""} - 案件明細`}
              </h3>
              <button
                onClick={() => setDrawerState(null)}
                className="close-button"
                aria-label="關閉抽屜"
              >
                ×
              </button>
            </div>
            <div className="drawer-body">
              {drawerState.type === "module" ? (
                <>
                  <p>總件數：{moduleDrilldownRows.length}</p>
                  {moduleDrilldownRows.length === 0 ? (
                    <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>
                      <p>沒有找到 {drawerState.name} 模組的詳細資料</p>
                      <p>請檢查資料來源或選擇其他模組</p>
                    </div>
                  ) : (
                    <div className="table-container">
                      <table>
                        <thead>
                          <tr>
                            <th>日期</th>
                            <th>分類</th>
                            <th>處理時間(分)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {moduleDrilldownRows.slice(0, 100).map((r, i) => (
                            <tr key={i}>
                              <td>{r["call_time"]}</td>
                              <td>{r["category"]}</td>
                              <td>{r["resolve_minute"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  {moduleDrilldownRows.length > 100 && (
                    <p style={{ marginTop: '10px', fontSize: '12px', color: '#666' }}>
                      只顯示前 100 筆，其餘 {moduleDrilldownRows.length - 100} 筆可透過資料匯出取得
                    </p>
                  )}
                </>
              ) : (
                <>
                  <p>區間：{drawerState.label}</p>
                  {drawerState.groupLabel ? <p>維度：{drawerState.groupLabel}</p> : null}
                  <p>筆數：{drawerState.rows.length}</p>
                  {drawerState.rows.length === 0 ? (
                    <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>暫無資料</div>
                  ) : (
                    <div className="table-container">
                      <table>
                        <thead>
                          <tr>
                            <th>日期</th>
                            <th>分類</th>
                            <th>模組</th>
                            <th>處理時間(分)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {drawerState.rows.slice(0, 100).map((r, i) => (
                            <tr key={i}>
                              <td>{r["call_time"]}</td>
                              <td>{r["category"]}</td>
                              <td>{r["module"]}</td>
                              <td>{r["resolve_minute"]}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  {drawerState.rows.length > 100 && (
                    <p style={{ marginTop: '10px', fontSize: '12px', color: '#666' }}>
                      只顯示前 100 筆，其餘 {drawerState.rows.length - 100} 筆可透過資料匯出取得
                    </p>
                  )}
                </>
              )}
            </div>
          </div>
        </>
      )}
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


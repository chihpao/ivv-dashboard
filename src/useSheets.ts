import { useQuery } from "@tanstack/react-query"
import { fetchCsv } from "./data"

const KPI_URL    = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=53717333&single=true&output=csv"
const TREND_URL  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1697285422&single=true&output=csv"
const TOP5_URL   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1042563257&single=true&output=csv"
const MODULE_BY_MONTH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1369816097&single=true&output=csv"

// Optional data sources for new charts (fill in if available)
// SLA distribution (either raw with a "resolve_minute" column or pre-aggregated with two columns [label, count])
const SLA_URL: string = ""
// Category trend over time: first column = date, remaining columns = categories counts
const CATEGORY_TREND_URL: string = ""

export function useSheets() {
  const kpiQ   = useQuery({ queryKey:["kpi"],   queryFn:() => fetchCsv(KPI_URL) })
  const trendQ = useQuery({ queryKey:["trend"], queryFn:() => fetchCsv(TREND_URL) })
  const topQ   = useQuery({ queryKey:["top5"],  queryFn:() => fetchCsv(TOP5_URL) })

  const slaQ = useQuery({
    queryKey: ["sla"],
    queryFn: () => fetchCsv(SLA_URL),
    enabled: Boolean(SLA_URL),
  })

  const categoryTrendQ = useQuery({
    queryKey: ["categoryTrend"],
    queryFn: () => fetchCsv(CATEGORY_TREND_URL),
    enabled: Boolean(CATEGORY_TREND_URL),
  })

  const modMonthQ = useQuery({
    queryKey: ["moduleByMonth"],
    queryFn: () => fetchCsv(MODULE_BY_MONTH_URL)
  })


  
  return {
    kpi:   (kpiQ.data ?? []) as any[],
    trend: (trendQ.data ?? []) as any[],
    top5:  (topQ.data ?? []) as any[],
    moduleByMonth: modMonthQ.data ?? [],
    sla:   (slaQ.data ?? []) as any[],
    categoryTrend: (categoryTrendQ.data ?? []) as any[],
    loading: kpiQ.isLoading || trendQ.isLoading || topQ.isLoading || modMonthQ.isLoading || slaQ.isLoading || categoryTrendQ.isLoading,
    
  }
}

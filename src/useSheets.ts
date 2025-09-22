import { useQuery } from "@tanstack/react-query"
import { fetchCsv } from "./data"

const TREND_URL  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1697285422&single=true&output=csv"
const MODULE_BY_MONTH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1369816097&single=true&output=csv"
const AVG_DURATION_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1897042583&single=true&output=csv"

// Optional data sources for new charts (fill in if available)
// Category trend over time: first column = date, remaining columns = categories counts
const CATEGORY_TREND_URL: string = ""

export function useSheets() {
  const trendQ = useQuery({ queryKey:["trend"], queryFn:() => fetchCsv(TREND_URL) })

  const categoryTrendQ = useQuery({
    queryKey: ["categoryTrend"],
    queryFn: () => fetchCsv(CATEGORY_TREND_URL),
    enabled: Boolean(CATEGORY_TREND_URL),
  })

  const modMonthQ = useQuery({
    queryKey: ["moduleByMonth"],
    queryFn: () => fetchCsv(MODULE_BY_MONTH_URL)
  })

  const avgDurationQ = useQuery({
    queryKey: ["avgCallDuration"],
    queryFn: () => fetchCsv(AVG_DURATION_URL)
  })


  
  return {
    trend: (trendQ.data ?? []) as any[],
    moduleByMonth: modMonthQ.data ?? [],
    categoryTrend: (categoryTrendQ.data ?? []) as any[],
    avgCallDuration: avgDurationQ.data ?? [],
    loading: trendQ.isLoading || modMonthQ.isLoading || categoryTrendQ.isLoading || avgDurationQ.isLoading,
    
  }
}

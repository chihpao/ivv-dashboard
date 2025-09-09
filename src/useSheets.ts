import { useQuery } from "@tanstack/react-query"
import { fetchCsv } from "./data"

const KPI_URL    = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=53717333&single=true&output=csv"
const TREND_URL  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1697285422&single=true&output=csv"
const TOP5_URL   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vS3CFFG7hUU8oLryXhjneEWI1ZbqqDzd6QyppdKkkWLBARdgpVPh4vWezp1fgyiN07Iop7kKm06XEnB/pub?gid=1042563257&single=true&output=csv"

export function useSheets() {
  const kpiQ   = useQuery({ queryKey:["kpi"],   queryFn:() => fetchCsv(KPI_URL) })
  const trendQ = useQuery({ queryKey:["trend"], queryFn:() => fetchCsv(TREND_URL) })
  const topQ   = useQuery({ queryKey:["top5"],  queryFn:() => fetchCsv(TOP5_URL) })

  return {
    kpi:   (kpiQ.data ?? []) as any[],
    trend: (trendQ.data ?? []) as any[],
    top5:  (topQ.data ?? []) as any[],
    loading: kpiQ.isLoading || trendQ.isLoading || topQ.isLoading,
  }
}

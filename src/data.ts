import Papa from "papaparse"
import type { ParseResult } from "papaparse"

interface CsvRow {
  [key: string]: any;
}

export function fetchCsv(url: string): Promise<CsvRow[]> {
  return new Promise((resolve, reject) => {
    Papa.parse<CsvRow>(url, {
      download: true,
      header: true,
      dynamicTyping: true,
      complete: (res: ParseResult<CsvRow>) => {
        // 正規化：去除欄位名稱與字串值的頭尾空白，避免因為表頭多空白而對不上 key
        const normalized = (res.data || []).map((row) => {
          const out: CsvRow = {}
          Object.keys(row || {}).forEach((rawKey) => {
            const k = typeof rawKey === 'string' ? rawKey.trim() : String(rawKey)
            const v = (row as any)[rawKey]
            out[k] = typeof v === 'string' ? v.trim() : v
          })
          return out
        })
        resolve(normalized)
      },
      error: (error: Error) => reject(error),
    })
  })
}

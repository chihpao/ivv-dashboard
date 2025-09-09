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
      complete: (res: ParseResult<CsvRow>) => resolve(res.data),
      error: (error: Error) => reject(error),
    })
  })
}

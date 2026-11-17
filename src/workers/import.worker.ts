import * as XLSX from "xlsx";

type ImportRequest = { files: File[] };
type ImportMessage =
  | { type: "progress"; fileName: string; fileIndex: number; fileCount: number; rowsRead: number; totalRows: number }
  | { type: "batch"; records: Array<{ id: string; source: string; rowNumber: number; data: Record<string, string> }> }
  | { type: "done"; totalImported: number }
  | { type: "error"; message: string };

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

self.onmessage = (event: MessageEvent<ImportRequest>) => {
  void (async () => {
    let totalImported = 0;
    try {
      const { files } = event.data;
      for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
        const file = files[fileIndex];
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true, dense: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;
        if (!sheet) continue;
        const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false, blankrows: false });
        if (!matrix.length) continue;
        const headers = (matrix[0] ?? []).map((value, index) => clean(value) || `Unnamed column ${index + 1}`);
        const totalRows = Math.max(0, matrix.length - 1);
        const batch: Array<{ id: string; source: string; rowNumber: number; data: Record<string, string> }> = [];
        for (let rowIndex = 1; rowIndex < matrix.length; rowIndex++) {
          const row = matrix[rowIndex] ?? [];
          const data: Record<string, string> = {};
          let hasValue = false;
          for (let col = 0; col < headers.length; col++) {
            const value = clean(row[col]);
            if (value) hasValue = true;
            data[headers[col]] = value;
          }
          if (!hasValue) continue;
          batch.push({ id: `${file.name}:${rowIndex + 1}:${crypto.randomUUID()}`, source: file.name, rowNumber: rowIndex + 1, data });
          if (batch.length >= 1000) {
            self.postMessage({ type: "batch", records: batch.splice(0, batch.length) } satisfies ImportMessage);
          }
          if (rowIndex % 500 === 0 || rowIndex === matrix.length - 1) {
            self.postMessage({ type: "progress", fileName: file.name, fileIndex, fileCount: files.length, rowsRead: rowIndex, totalRows } satisfies ImportMessage);
          }
        }
        if (batch.length) self.postMessage({ type: "batch", records: batch } satisfies ImportMessage);
        totalImported += totalRows;
      }
      self.postMessage({ type: "done", totalImported } satisfies ImportMessage);
    } catch (error) {
      self.postMessage({ type: "error", message: error instanceof Error ? error.message : "Could not parse the selected spreadsheets." } satisfies ImportMessage);
    }
  })();
};

export {};

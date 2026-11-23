import * as XLSX from "xlsx";

type ImportRequest = { files: File[] };
type ImportedRecord = { id: string; source: string; rowNumber: number; data: Record<string, string> };
type ImportMessage =
  | { type: "progress"; fileName: string; fileIndex: number; fileCount: number; rowsRead: number; totalRows: number }
  | { type: "batch"; records: ImportedRecord[] }
  | { type: "done"; totalImported: number }
  | { type: "error"; message: string };

function clean(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value).trim().replace(/\s+/g, " ");
}

// Read worksheet rows directly from SheetJS' dense worksheet. Avoid
// sheet_to_json(header:1), which creates a second full in-memory matrix and
// can spike memory on large alumni workbooks.
function getCellValue(sheet: XLSX.WorkSheet, row: number, column: number): unknown {
  const dense = sheet as unknown as Array<Array<XLSX.CellObject | undefined>> & { [key: string]: unknown };
  const denseCell = Array.isArray(dense[row]) ? dense[row]?.[column] : undefined;
  const cell = denseCell ?? sheet[XLSX.utils.encode_cell({ r: row, c: column })] as XLSX.CellObject | undefined;
  if (!cell) return "";
  return cell.w ?? cell.v ?? "";
}

self.onmessage = (event: MessageEvent<ImportRequest>) => {
  void (async () => {
    let totalImported = 0;
    try {
      const files = Array.from(event.data.files ?? []);
      for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
        const file = files[fileIndex];
        if (!file || file.size === 0) continue;
        self.postMessage({ type: "progress", fileName: file.name, fileIndex, fileCount: files.length, rowsRead: 0, totalRows: 0 } satisfies ImportMessage);

        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true, dense: true, raw: false });
        const sheetName = workbook.SheetNames[0];
        const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;
        if (!sheet || !sheet["!ref"]) continue;

        const range = XLSX.utils.decode_range(sheet["!ref"]);
        const headerRow = range.s.r;
        const headers: string[] = [];
        for (let column = range.s.c; column <= range.e.c; column++) {
          headers.push(clean(getCellValue(sheet, headerRow, column)) || `Unnamed column ${column - range.s.c + 1}`);
        }

        const totalRows = Math.max(0, range.e.r - headerRow);
        let batch: ImportedRecord[] = [];
        for (let rowIndex = headerRow + 1; rowIndex <= range.e.r; rowIndex++) {
          const data: Record<string, string> = {};
          let hasValue = false;
          for (let column = range.s.c; column <= range.e.c; column++) {
            const value = clean(getCellValue(sheet, rowIndex, column));
            if (value) hasValue = true;
            data[headers[column - range.s.c]] = value;
          }
          if (hasValue) {
            batch.push({
              id: `${file.name}:${rowIndex + 1}:${fileIndex}:${crypto.randomUUID()}`,
              source: file.name,
              rowNumber: rowIndex + 1,
              data,
            });
          }

          if (batch.length >= 500) {
            self.postMessage({ type: "batch", records: batch } satisfies ImportMessage);
            totalImported += batch.length;
            batch = [];
          }
          if ((rowIndex - headerRow) % 250 === 0 || rowIndex === range.e.r) {
            self.postMessage({ type: "progress", fileName: file.name, fileIndex, fileCount: files.length, rowsRead: rowIndex - headerRow, totalRows } satisfies ImportMessage);
          }
        }
        if (batch.length) {
          self.postMessage({ type: "batch", records: batch } satisfies ImportMessage);
          totalImported += batch.length;
        }
        // Release references before parsing the next selected workbook.
        delete workbook.Sheets[sheetName];
      }
      self.postMessage({ type: "done", totalImported } satisfies ImportMessage);
    } catch (error) {
      self.postMessage({ type: "error", message: error instanceof Error ? error.message : "Could not parse the selected spreadsheets." } satisfies ImportMessage);
    }
  })();
};

export {};

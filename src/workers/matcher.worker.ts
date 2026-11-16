import { compareRecords, type StudentRecord } from "@/lib/matcher";

self.onmessage = (event: MessageEvent<{ records: StudentRecord[]; threshold: number; selectedFields?: string[] }>) => {
  try {
    const { records, threshold, selectedFields } = event.data;
    const matches = compareRecords(records, threshold, selectedFields);
    self.postMessage({ matches });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Unknown matching error" });
  }
};

export {};

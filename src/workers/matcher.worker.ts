import { compareRecords, type StudentRecord } from "@/lib/matcher";

self.onmessage = (event: MessageEvent<{ records: StudentRecord[]; threshold: number }>) => {
  try {
    const { records, threshold } = event.data;
    const matches = compareRecords(records, threshold);
    self.postMessage({ matches });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Unknown matching error" });
  }
};

export {};

import { compareRecords, type StudentRecord, type NormalizationRule } from "@/lib/matcher";

self.onmessage = (event: MessageEvent<{ records: StudentRecord[]; threshold: number; selectedFields?: string[]; normalizationRules?: NormalizationRule[] }>) => {
  try {
    const { records, threshold, selectedFields, normalizationRules } = event.data;
    const matches = compareRecords(records, threshold, selectedFields, normalizationRules);
    self.postMessage({ matches });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : "Unknown matching error" });
  }
};

export {};

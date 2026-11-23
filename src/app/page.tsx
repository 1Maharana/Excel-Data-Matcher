"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import RelationshipScene from "./RelationshipScene";
import {
  canonicalHeader,
  compareRecords,
  MATCH_FIELDS,
  comparisonValue,
  normalizeValue,
  standardFields,
  type StudentRecord
} from "@/lib/matcher";

type Tab = "records" | "matches" | "missing" | "unmatched" | "visualization";
type ReviewDecision = "Confirmed same person" | "Not a match" | "Kept unchanged";
type EffectiveDecision = ReviewDecision | "Auto-confirmed same person";

const buttonBase = "rounded-xl px-3.5 py-2 text-sm font-semibold transition";
const buttonPrimary = `${buttonBase} bg-indigo-600 text-white hover:bg-indigo-700`;
const buttonSecondary = `${buttonBase} border border-slate-200 bg-white text-slate-700 hover:bg-slate-50`;

export default function Home() {
  const [records, setRecords] = useState<StudentRecord[]>([]);
  const [query, setQuery] = useState("");
  const [sourceFilter, setSourceFilter] = useState("All files");
  const [missingOnly, setMissingOnly] = useState(false);
  const [tab, setTab] = useState<Tab>("records");
  const [threshold, setThreshold] = useState(60);
  const [selectedCompareFields, setSelectedCompareFields] = useState<string[]>(MATCH_FIELDS.map((field) => field.name));
  const [showExportSettings, setShowExportSettings] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [importProgress, setImportProgress] = useState(0);
  const [importStatus, setImportStatus] = useState("");
  const [matches, setMatches] = useState<ReturnType<typeof compareRecords>>([]);
  const [matching, setMatching] = useState(false);
  const [selectedMatch, setSelectedMatch] = useState<string | null>(null);
  const [decisions, setDecisions] = useState<Record<string, ReviewDecision>>({});
  const [selectedMatchIds, setSelectedMatchIds] = useState<string[]>([]);
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [showMapping, setShowMapping] = useState(false);
  const [collapseRepeatedSubmissions, setCollapseRepeatedSubmissions] = useState(false);
  const [autoConfirmEnabled, setAutoConfirmEnabled] = useState(false);
  const [masterSource, setMasterSource] = useState("");
  const [selectedD2Source, setSelectedD2Source] = useState("");
  const [exportComposition, setExportComposition] = useState<"union" | "enriched" | "original">("enriched");
  const [selectedExportFields, setSelectedExportFields] = useState<string[]>([]);
  const [exportFieldsInitialized, setExportFieldsInitialized] = useState(false);
  const [includeOriginalValues, setIncludeOriginalValues] = useState(false);
  const [includeEnrichmentStatus, setIncludeEnrichmentStatus] = useState(true);
  const [includeReviewSheets, setIncludeReviewSheets] = useState(true);
  const [includeRejectedSheet, setIncludeRejectedSheet] = useState(true);
  const [includeMatchedSheet, setIncludeMatchedSheet] = useState(true);

  const sourceColumns = useMemo(() => Array.from(new Set(records.flatMap((r) => Object.keys(r.data)))), [records]);
  const sources = useMemo(() => Array.from(new Set(records.map((r) => r.source))), [records]);
  useEffect(() => {
    if (sources.length && (!masterSource || !sources.includes(masterSource))) setMasterSource(sources[0]);
    const availableD2 = sources.find((source) => source !== masterSource);
    if (sources.length > 1 && (!selectedD2Source || !sources.includes(selectedD2Source) || selectedD2Source === masterSource)) setSelectedD2Source(availableD2 ?? "");
    if (sources.length < 2) setSelectedD2Source("");
  }, [sources, masterSource, selectedD2Source]);
  const mappedRecords = useMemo(() => records.map((r) => {
    const data: Record<string, string> = {};
    for (const [key, value] of Object.entries(r.data)) {
      const mappedKey = mapping[key] || canonicalHeader(key);
      if (!(mappedKey in data) || (!data[mappedKey] && value)) data[mappedKey] = value;
    }
    return { ...r, data };
  }), [records, mapping]);
  // Display columns must come from normalized data keys, not original source headers.
  const columns = useMemo(() => Array.from(new Set(mappedRecords.flatMap((r) => Object.keys(r.data)))), [mappedRecords]);
  useEffect(() => {
    if (columns.length && !exportFieldsInitialized) {
      setSelectedExportFields(columns);
      setExportFieldsInitialized(true);
    }
  }, [columns, exportFieldsInitialized]);

  // Optional per-file deduplication: repeated submissions with the same normalized
  // name + graduation year + department are represented by the most complete row
  // for matching. Original rows remain untouched and available for exports.
  const comparisonGroups = useMemo(() => {
    if (!collapseRepeatedSubmissions) {
      return { records: mappedRecords, membersByRepresentative: new Map<string, StudentRecord[]>() };
    }
    const groups = new Map<string, StudentRecord[]>();
    const ungrouped: StudentRecord[] = [];
    const compactKey = (value: string) => normalizeValue(value).toLowerCase().normalize("NFKD")
      .replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
    for (const record of mappedRecords) {
      const name = compactKey(record.data.Name ?? record.data["Full Name"] ?? "");
      const yearRaw = normalizeValue(record.data["Graduation Year"] ?? "");
      const year = yearRaw.match(/\b(19|20)\d{2}\b/)?.[0] ?? compactKey(yearRaw);
      // Use the same approved department aliases as the matching engine so
      // legacy/current department labels group consistently within each file.
      const department = comparisonValue("Department", record.data.Department ?? "");
      // Avoid collapsing same-name rows when there is no supporting batch/department evidence.
      if (!name || (!year && !department)) { ungrouped.push(record); continue; }
      const key = `${record.source}::${name}::${year}::${department}`;
      const bucket = groups.get(key);
      if (bucket) bucket.push(record); else groups.set(key, [record]);
    }
    const membersByRepresentative = new Map<string, StudentRecord[]>();
    const representatives: StudentRecord[] = [...ungrouped];
    for (const bucket of groups.values()) {
      if (bucket.length === 1) { representatives.push(bucket[0]); continue; }
      const completeness = (r: StudentRecord) => Object.values(r.data).filter((v) => normalizeValue(v)).length;
      const ranked = [...bucket].sort((a, b) => completeness(b) - completeness(a) || a.rowNumber - b.rowNumber);
      const representative = ranked[0];
      const mergedData = { ...representative.data };
      // Fill missing fields from other submissions; never overwrite a non-empty representative value.
      for (const field of new Set(bucket.flatMap((r) => Object.keys(r.data)))) {
        if (!normalizeValue(mergedData[field])) {
          const donor = ranked.find((r) => normalizeValue(r.data[field]));
          if (donor) mergedData[field] = donor.data[field];
        }
      }
      representatives.push({ ...representative, data: mergedData });
      membersByRepresentative.set(representative.id, bucket);
    }
    return { records: representatives, membersByRepresentative };
  }, [mappedRecords, collapseRepeatedSubmissions]);

  const comparisonRecords = useMemo(() => {
    if (!masterSource || !selectedD2Source || masterSource === selectedD2Source) return comparisonGroups.records;
    return comparisonGroups.records.filter((record) => record.source === masterSource || record.source === selectedD2Source);
  }, [comparisonGroups.records, masterSource, selectedD2Source]);

  // Run the expensive matching pass off the main UI thread. This keeps search,
  // scrolling, and the rest of the interface responsive on large workbooks.
  useEffect(() => {
    if (comparisonRecords.length < 2) {
      setMatches([]);
      setMatching(false);
      return;
    }

    setMatching(true);
    const worker = new Worker(new URL("../workers/matcher.worker.ts", import.meta.url), { type: "module" });
    worker.onmessage = (event: MessageEvent<{ matches?: ReturnType<typeof compareRecords>; error?: string }>) => {
      if (event.data.error) {
        setError(event.data.error);
      } else {
        setMatches(event.data.matches ?? []);
      }
      setMatching(false);
      worker.terminate();
    };
    worker.onerror = () => {
      setError("Matching failed in the background worker. Try fewer files or a smaller dataset.");
      setMatching(false);
      worker.terminate();
    };
    worker.postMessage({ records: comparisonRecords, threshold, selectedFields: selectedCompareFields });
    return () => worker.terminate();
  }, [comparisonRecords, threshold, selectedCompareFields, sources.length]);
  // When consolidation is enabled, every record-facing section uses the same per-file
  // representative dataset as matching. This keeps All Records, Missing, Unmatched,
  // filtered exports, and Union Master consistent instead of showing duplicate source rows.
  const activeRecords = collapseRepeatedSubmissions ? comparisonGroups.records : mappedRecords;

  const filteredRecords = useMemo(() => activeRecords.filter((r) => {
    const text = [r.source, ...Object.values(r.data)].join(" ").toLowerCase();
    if (!text.includes(query.toLowerCase())) return false;
    if (sourceFilter !== "All files" && r.source !== sourceFilter) return false;
    if (missingOnly && !Object.values(r.data).some((v) => !normalizeValue(v))) return false;
    return true;
  }), [activeRecords, query, sourceFilter, missingOnly]);

  const effectiveDecision = (match: (typeof matches)[number]): EffectiveDecision | undefined => {
    if (decisions[match.id]) return decisions[match.id];
    if (autoConfirmEnabled && match.autoConfirmEligible) return "Auto-confirmed same person";
    return undefined;
  };

  const pendingMatches = useMemo(() => matches.filter((m) => !decisions[m.id] && !(autoConfirmEnabled && m.autoConfirmEligible)).filter((m) => {
    const text = [m.left.source, m.right.source, ...Object.values(m.left.data), ...Object.values(m.right.data)].join(" ").toLowerCase();
    return text.includes(query.toLowerCase()) &&
      (sourceFilter === "All files" || m.left.source === sourceFilter || m.right.source === sourceFilter);
  }), [matches, decisions, autoConfirmEnabled, query, sourceFilter]);

  const approvedMatches = useMemo(() => matches.filter((m) => {
    const decision = effectiveDecision(m);
    return decision === "Confirmed same person" || decision === "Auto-confirmed same person";
  }), [matches, decisions, autoConfirmEnabled]);

  const selectedPendingMatches = useMemo(() => pendingMatches.filter((m) => selectedMatchIds.includes(m.id)), [pendingMatches, selectedMatchIds]);
  const allPendingSelected = pendingMatches.length > 0 && pendingMatches.every((m) => selectedMatchIds.includes(m.id));

  function toggleMatchSelection(id: string, checked: boolean) {
    setSelectedMatchIds((current) => checked ? (current.includes(id) ? current : [...current, id]) : current.filter((item) => item !== id));
  }

  function selectAllPendingMatches() {
    setSelectedMatchIds(pendingMatches.map((match) => match.id));
  }

  function rejectSelectedMatches() {
    const idsToReject = selectedPendingMatches.map((match) => match.id);
    if (!idsToReject.length) return;
    const crossFileCount = selectedPendingMatches.filter((match) => match.left.source !== match.right.source).length;
    const withinFileCount = idsToReject.length - crossFileCount;
    const message = `Mark ${idsToReject.length.toLocaleString()} selected candidate pairs as NOT a match?\n\n${crossFileCount.toLocaleString()} cross-file pairs will be kept separate.\n${withinFileCount.toLocaleString()} within-file pairs will remain separate submissions.\n\nNo original rows will be changed or deleted.`;
    if (typeof window !== "undefined" && !window.confirm(message)) return;
    setDecisions((current) => { const next = { ...current }; for (const id of idsToReject) next[id] = "Not a match"; return next; });
    setSelectedMatchIds([]);
  }

  function approveSelectedMatches() {
    const idsToApprove = selectedPendingMatches.map((match) => match.id);
    if (!idsToApprove.length) return;
    const crossFileCount = selectedPendingMatches.filter((match) => match.left.source !== match.right.source).length;
    const withinFileCount = idsToApprove.length - crossFileCount;
    const message = `Approve ${idsToApprove.length.toLocaleString()} selected candidate pairs?\n\n${crossFileCount.toLocaleString()} cross-file matches will become eligible for D1 approved enrichment.\n${withinFileCount.toLocaleString()} within-file pairs will be marked as confirmed duplicates for source-file consolidation.\n\nNo original rows will be deleted.`;
    if (typeof window !== "undefined" && !window.confirm(message)) return;
    setDecisions((current) => {
      const next = { ...current };
      for (const id of idsToApprove) next[id] = "Confirmed same person";
      return next;
    });
    setSelectedMatchIds([]);
  }

  const missingRecords = useMemo(() => filteredRecords.filter((r) =>
    Object.values(r.data).some((value) => !normalizeValue(value))
  ), [filteredRecords]);

  const matchedRecordIds = useMemo(() => {
    const ids = new Set<string>();
    for (const match of matches) {
      for (const representativeId of [match.left.id, match.right.id]) {
        ids.add(representativeId);
        for (const member of comparisonGroups.membersByRepresentative.get(representativeId) ?? []) ids.add(member.id);
      }
    }
    return ids;
  }, [matches, comparisonGroups]);
  const unmatchedRecords = useMemo(() => filteredRecords.filter((r) => !matchedRecordIds.has(r.id)), [filteredRecords, matchedRecordIds]);
  const activeMatch = matches.find((m) => m.id === selectedMatch);

  async function importFiles(fileList: FileList | null) {
    if (!fileList?.length) return;
    setLoading(true);
    setError("");
    setImportProgress(0);
    setImportStatus("Preparing spreadsheets…");
    const files = Array.from(fileList);
    const invalid = files.find((file) => !["xlsx", "xls", "csv"].includes(file.name.split(".").pop()?.toLowerCase() ?? ""));
    if (invalid) {
      setError(`${invalid.name} is not an Excel or CSV file.`);
      setLoading(false);
      setImportStatus("");
      return;
    }

    const imported: StudentRecord[] = [];
    const worker = new Worker(new URL("../workers/import.worker.ts", import.meta.url), { type: "module" });
    let settled = false;
    const finish = (message?: string) => {
      if (settled) return;
      settled = true;
      worker.terminate();
      if (message) setError(message);
      setLoading(false);
      setImportStatus("");
    };
    worker.onmessage = (event: MessageEvent<{
      type: "progress" | "batch" | "done" | "error";
      fileName?: string; fileIndex?: number; fileCount?: number; rowsRead?: number; totalRows?: number;
      records?: StudentRecord[]; totalImported?: number; message?: string;
    }>) => {
      const data = event.data;
      if (data.type === "batch" && data.records) {
        imported.push(...data.records);
      } else if (data.type === "progress") {
        const fileFraction = data.totalRows ? data.rowsRead! / data.totalRows : 1;
        const overall = ((data.fileIndex! + fileFraction) / data.fileCount!) * 100;
        setImportProgress(Math.min(99, Math.round(overall)));
        setImportStatus(`Reading ${data.fileName} · ${data.rowsRead?.toLocaleString()} of ${data.totalRows?.toLocaleString()} rows`);
      } else if (data.type === "done") {
        if (!imported.length) {
          finish("No data rows found. Check that the first worksheet has a header row and records.");
          return;
        }
        setRecords((old) => [...old, ...imported]);
        setTab("records");
        setImportProgress(100);
        finish();
      } else if (data.type === "error") {
        finish(data.message || "Could not read these files.");
      }
    };
    worker.onerror = () => finish("Spreadsheet import failed in the background worker. Try saving the workbook as .xlsx or .csv and importing again.");
    try {
      // File/Blob objects are structured-cloned efficiently; the worker reads
      // one workbook at a time so multiple large uploads do not duplicate buffers.
      worker.postMessage({ files });
    } catch (error) {
      finish(error instanceof Error ? error.message : "Could not load the selected files.");
    }
  }

  function exportRows(rows: Record<string, unknown>[], filename: string) {
    if (!rows.length) { setError("There are no rows to export."); return; }
    const sheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Results");
    XLSX.writeFile(workbook, filename);
    setError("");
  }

  function exportRecords(rows: StudentRecord[], filename: string) {
    const exportData = rows.map((r) => {
      const members = comparisonGroups.membersByRepresentative.get(r.id) ?? [r];
      return {
        "Source File": r.source,
        "Representative Source Row": r.rowNumber,
        "Repeated Submissions Grouped": members.length > 1 ? "Yes" : "No",
        "Submission Count": members.length,
        "Source Rows In Group": members.map((member) => `${member.source}: ${member.rowNumber}`).join("; "),
        ...r.data
      };
    });
    exportRows(exportData, filename);
  }

  function exportMatches() {
    exportRows(matches.map((m) => ({
      "Review Decision": effectiveDecision(m) || "Pending review",
      "Auto-confirm eligible": m.autoConfirmEligible ? "Yes" : "No",
      Status: m.status,
      "Similarity Score": m.score,
      Reasons: m.reasons.join("; "),
      "File A": m.left.source,
      "Row A": m.left.rowNumber,
      ...Object.fromEntries(Object.entries(m.left.data).map(([k, v]) => [`A - ${k}`, v])),
      "File B": m.right.source,
      "Row B": m.right.rowNumber,
      ...Object.fromEntries(Object.entries(m.right.data).map(([k, v]) => [`B - ${k}`, v]))
    })), "match-review.xlsx");
  }

  function exportUnion() {
    if (!mappedRecords.length) {
      setError("Import at least one file before exporting the union master.");
      return;
    }

    // Only merge records that a reviewer explicitly confirmed as the same person.
    // Unreviewed/possible matches remain separate to avoid accidentally deleting alumni.
    const parent = new Map<string, string>();
    const find = (id: string): string => {
      const current = parent.get(id);
      if (!current) { parent.set(id, id); return id; }
      if (current === id) return id;
      const root = find(current);
      parent.set(id, root);
      return root;
    };
    const unite = (a: string, b: string) => {
      const rootA = find(a);
      const rootB = find(b);
      if (rootA !== rootB) parent.set(rootB, rootA);
    };

    // When the toggle is enabled, consolidate the repeated submissions that were
    // represented by one best-information record for comparison.
    if (collapseRepeatedSubmissions) {
      for (const members of comparisonGroups.membersByRepresentative.values()) {
        const first = members[0];
        for (const member of members.slice(1)) unite(first.id, member.id);
      }
    }
    const confirmedMatches = matches.filter((match) => {
      const decision = effectiveDecision(match);
      return decision === "Confirmed same person" || decision === "Auto-confirmed same person";
    });
    for (const match of confirmedMatches) {
      const leftMembers = comparisonGroups.membersByRepresentative.get(match.left.id) ?? [match.left];
      const rightMembers = comparisonGroups.membersByRepresentative.get(match.right.id) ?? [match.right];
      for (const left of leftMembers) for (const right of rightMembers) unite(left.id, right.id);
    }

    const groups = new Map<string, StudentRecord[]>();
    for (const record of mappedRecords) {
      const root = find(record.id);
      const group = groups.get(root);
      if (group) group.push(record);
      else groups.set(root, [record]);
    }

    const rows = Array.from(groups.values()).map((group) => {
      // Prefer the most complete source row as the primary version, then fill its
      // blanks from other confirmed submissions. Conflicting non-empty values are
      // retained in dedicated columns rather than silently discarded.
      const completeness = (r: StudentRecord) => Object.values(r.data).filter((v) => normalizeValue(v)).length;
      const rankedGroup = [...group].sort((a, b) => completeness(b) - completeness(a) || a.rowNumber - b.rowNumber);
      const merged: Record<string, unknown> = {
        "Source Files": Array.from(new Set(rankedGroup.map((r) => r.source))).join("; "),
        "Source Rows": rankedGroup.map((r) => `${r.source}: ${r.rowNumber}`).join("; "),
        "Record Count Merged": rankedGroup.length,
        "Primary Source File": rankedGroup[0]?.source ?? "",
        "Primary Source Row": rankedGroup[0]?.rowNumber ?? ""
      };
      const allFields = Array.from(new Set(rankedGroup.flatMap((r) => Object.keys(r.data))));
      const conflicts: string[] = [];

      for (const field of allFields) {
        const valuesByPriority = rankedGroup.map((r) => normalizeValue(r.data[field])).filter(Boolean);
        const distinctValues = Array.from(new Set(valuesByPriority));
        merged[field] = valuesByPriority[0] ?? "";
        if (distinctValues.length > 1) {
          conflicts.push(`${field}: ${distinctValues.join(" | ")}`);
          merged[`${field} (All source values)`] = distinctValues.join(" | ");
        }
      }
      merged["Field Conflicts"] = conflicts.join("; ");
      merged["Review Status"] = rankedGroup.length > 1 ? "Merged confirmed match" : "Single source record";
      return merged;
    });

    exportRows(rows, "union-master.xlsx");
  }

  function exportConfiguredMaster() {
    if (!masterSource) { setError("Choose the master source file first."); return; }
    if (!selectedD2Source || selectedD2Source === masterSource) { setError("Choose a different D2 dataset."); return; }
    const incomingSources = [selectedD2Source];
    const masterRecords = mappedRecords.filter((record) => record.source === masterSource);
    if (!masterRecords.length) { setError("No records were found in the selected master source."); return; }
    if (!selectedExportFields.length) { setError("Select at least one field to export."); return; }

    const confirmed = matches.filter((match) => {
      const decision = effectiveDecision(match);
      return decision === "Confirmed same person" || decision === "Auto-confirmed same person";
    });
    const pending = matches.filter((match) => !decisions[match.id] && !(autoConfirmEnabled && match.autoConfirmEligible));
    const rejected = matches.filter((match) => decisions[match.id] === "Not a match");

    // Map confirmed cross-file matches from each master record to its incoming record(s).
    const enrichments = new Map<string, StudentRecord[]>();
    for (const match of confirmed) {
      const masterSide = match.left.source === masterSource ? match.left : match.right.source === masterSource ? match.right : null;
      const incomingSide = match.left.source === masterSource ? match.right : match.right.source === masterSource ? match.left : null;
      if (!masterSide || !incomingSide || !incomingSources.includes(incomingSide.source)) continue;
      const list = enrichments.get(masterSide.id) ?? [];
      list.push(incomingSide);
      enrichments.set(masterSide.id, list);
    }

    const selectedFields = selectedExportFields;
    const masterRows = masterRecords.map((master) => {
      const donors = enrichments.get(master.id) ?? [];
      const output: Record<string, unknown> = {
        "Master Source File": master.source,
        "Master Source Row": master.rowNumber
      };
      for (const field of selectedFields) {
        const original = normalizeValue(master.data[field] ?? "");
        const donorValues = donors.map((donor) => normalizeValue(donor.data[field] ?? "")).filter(Boolean);
        const enrichedValue = original || donorValues[0] || "";
        output[field] = enrichedValue;
        if (includeOriginalValues) output[`${field} (Original D1)`] = original;
      }
      if (includeEnrichmentStatus) {
        const enrichedFields = selectedFields.filter((field) => !normalizeValue(master.data[field] ?? "") && donors.some((donor) => normalizeValue(donor.data[field] ?? "")));
        output["Enrichment Status"] = enrichedFields.length ? "Enriched from approved match" : donors.length ? "Matched; no fields enriched" : "No approved D2 match";
        output["Enriched Fields"] = enrichedFields.join("; ");
        output["Approved D2 Sources"] = Array.from(new Set(donors.map((donor) => donor.source))).join("; ");
      }
      return output;
    });

    let finalRows = masterRows;
    if (exportComposition === "union") {
      const candidateIncomingIds = new Set<string>();
      for (const match of matches) {
        const incomingSide = match.left.source === masterSource ? match.right : match.right.source === masterSource ? match.left : null;
        if (incomingSide && incomingSources.includes(incomingSide.source)) candidateIncomingIds.add(incomingSide.id);
      }
      const confirmedIncomingIds = new Set<string>();
      for (const match of confirmed) {
        const incomingSide = match.left.source === masterSource ? match.right : match.right.source === masterSource ? match.left : null;
        if (incomingSide && incomingSources.includes(incomingSide.source)) confirmedIncomingIds.add(incomingSide.id);
      }
      const incomingOnly = mappedRecords.filter((record) => incomingSources.includes(record.source) && !confirmedIncomingIds.has(record.id) && !pending.some((match) => match.left.id === record.id || match.right.id === record.id));
      const unmatchedIncoming = incomingOnly.filter((record) => !candidateIncomingIds.has(record.id) || rejected.some((match) => match.left.id === record.id || match.right.id === record.id));
      const appended = unmatchedIncoming.map((record) => {
        const row: Record<string, unknown> = { "Record Origin": "Unmatched D2", "Source File": record.source, "Source Row": record.rowNumber };
        for (const field of selectedFields) row[field] = normalizeValue(record.data[field] ?? "");
        if (includeEnrichmentStatus) { row["Enrichment Status"] = "New D2-only record"; row["Enriched Fields"] = ""; }
        return row;
      });
      finalRows = [...masterRows, ...appended];
    }
    if (exportComposition === "original") {
      finalRows = masterRecords.map((record) => {
        const row: Record<string, unknown> = { "Master Source File": record.source, "Master Source Row": record.rowNumber };
        for (const field of selectedFields) row[field] = normalizeValue(record.data[field] ?? "");
        if (includeEnrichmentStatus) { row["Enrichment Status"] = "Original D1"; row["Enriched Fields"] = ""; }
        return row;
      });
    }

    const workbook = XLSX.utils.book_new();
    const appendSheet = (rows: Record<string, unknown>[], name: string) => {
      const safeRows = rows.length ? rows : [{ Note: "No records in this category." }];
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(safeRows), name.slice(0, 31));
    };
    appendSheet(finalRows, "Final_Master");
    if (includeMatchedSheet) {
      appendSheet(confirmed.map((match) => ({ "Review Decision": effectiveDecision(match), "Similarity Score": match.score, Reasons: match.reasons.join("; "), "Master File": masterSource, "Master Row": match.left.source === masterSource ? match.left.rowNumber : match.right.rowNumber, "Incoming File": match.left.source === masterSource ? match.right.source : match.left.source, "Incoming Row": match.left.source === masterSource ? match.right.rowNumber : match.left.rowNumber })), "Matched_Records");
    }
    if (includeReviewSheets) {
      appendSheet(pending.map((match) => ({ Status: "Pending review", "Similarity Score": match.score, Reasons: match.reasons.join("; "), "File A": match.left.source, "Row A": match.left.rowNumber, "Name A": match.left.data.Name ?? match.left.data["Full Name"] ?? "", "File B": match.right.source, "Row B": match.right.rowNumber, "Name B": match.right.data.Name ?? match.right.data["Full Name"] ?? "" })), "Pending_Review");
    }
    if (includeRejectedSheet) {
      appendSheet(rejected.map((match) => ({ Status: "Rejected match", "Similarity Score": match.score, Reasons: match.reasons.join("; "), "File A": match.left.source, "Row A": match.left.rowNumber, "Name A": match.left.data.Name ?? match.left.data["Full Name"] ?? "", "File B": match.right.source, "Row B": match.right.rowNumber, "Name B": match.right.data.Name ?? match.right.data["Full Name"] ?? "" })), "Rejected_Matches");
    }
    XLSX.writeFile(workbook, "configured-master-export.xlsx");
    setError("");
  }

  function resetWorkspace() {
    setRecords([]); setQuery(""); setSourceFilter("All files"); setMissingOnly(false);
    setDecisions({}); setSelectedMatchIds([]); setSelectedMatch(null); setMapping({}); setError(""); setTab("records"); setCollapseRepeatedSubmissions(false); setAutoConfirmEnabled(false);
    setMasterSource(""); setSelectedD2Source(""); setExportComposition("enriched"); setSelectedExportFields([]); setExportFieldsInitialized(false); setShowExportSettings(false); setSelectedCompareFields(MATCH_FIELDS.map((field) => field.name));
  }

  const tabItems: { id: Tab; label: string; count: number }[] = [
    { id: "records", label: "All records", count: filteredRecords.length },
    { id: "matches", label: "Match review", count: pendingMatches.length },
    { id: "missing", label: "Missing information", count: missingRecords.length },
    { id: "unmatched", label: "Unmatched", count: unmatchedRecords.length },
    { id: "visualization", label: "3D Visualization", count: matches.filter((m) => m.left.source !== m.right.source).length }
  ];

  if (tab === "visualization") {
    const pairMatches = matches.filter((m) => m.left.source === masterSource && m.right.source === selectedD2Source || m.left.source === selectedD2Source && m.right.source === masterSource);
    const pairApproved = pairMatches.filter((m) => { const d = effectiveDecision(m); return d === "Confirmed same person" || d === "Auto-confirmed same person"; }).length;
    const pairRejected = pairMatches.filter((m) => effectiveDecision(m) === "Not a match").length;
    const pairPending = pairMatches.length - pairApproved - pairRejected;
    const d1Count = mappedRecords.filter((r) => r.source === masterSource).length;
    const d2Count = mappedRecords.filter((r) => r.source === selectedD2Source).length;
    const linkedD1 = new Set(pairMatches.filter((m) => { const d = effectiveDecision(m); return d === "Confirmed same person" || d === "Auto-confirmed same person"; }).map((m) => m.left.source === masterSource ? m.left.id : m.right.id)).size;
    const linkedD2 = new Set(pairMatches.filter((m) => { const d = effectiveDecision(m); return d === "Confirmed same person" || d === "Auto-confirmed same person"; }).map((m) => m.left.source === selectedD2Source ? m.left.id : m.right.id)).size;
    return <VisualizationWorkspace masterSource={masterSource} d2Source={selectedD2Source} sources={sources} onD1Change={setMasterSource} onD2Change={setSelectedD2Source} onBack={() => setTab("matches")} d1Count={d1Count} d2Count={d2Count} candidateCount={pairMatches.length} confirmedCount={pairApproved} pendingCount={pairPending} rejectedCount={pairRejected} unmatchedD1={Math.max(0, d1Count - linkedD1)} unmatchedD2={Math.max(0, d2Count - linkedD2)} relationships={pairMatches.map((match) => { const d1IsLeft = match.left.source === masterSource; return { ...match, left: d1IsLeft ? match.left : match.right, right: d1IsLeft ? match.right : match.left, status: effectiveDecision(match) ?? match.status }; })} />;
  }

  return (
    <main className="min-h-screen px-4 py-7 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-[1500px]">
        <header className="mb-7 flex flex-col justify-between gap-4 md:flex-row md:items-center">
          <div>
            <div className="mb-2 flex items-center gap-2 text-xs font-bold uppercase tracking-[.18em] text-indigo-600">
              <span className="inline-block h-2 w-2 rounded-full bg-emerald-500" /> Local-first workspace
            </div>
            <h1 className="text-3xl font-bold tracking-tight sm:text-4xl">Excel Data Matcher</h1>
            <p className="mt-2 max-w-2xl text-sm leading-6 text-slate-500">
              Consolidate spreadsheets, find duplicate candidates, review differences, and export clean reports.
            </p>
          </div>
          <button onClick={resetWorkspace} className={buttonSecondary}>Clear workspace</button>
        </header>

        <section className="mb-5 rounded-2xl border border-dashed border-indigo-300 bg-white p-5 shadow-sm sm:p-6">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <h2 className="font-bold">Import spreadsheets</h2>
              <p className="mt-1 text-sm text-slate-500">Upload one file to inspect its fields and missing information, or upload multiple files to compare records. The first worksheet is imported.</p>
            </div>
            <label className={`${buttonPrimary} inline-flex items-center justify-center gap-2`}>
              <span>{loading ? "Importing…" : "+ Choose files"}</span>
              <input type="file" accept=".xlsx,.xls,.csv" multiple disabled={loading} className="hidden"
                onChange={(e) => { void importFiles(e.target.files); e.currentTarget.value = ""; }} />
            </label>
          </div>
          <p className="mt-3 text-xs text-slate-400">Files are parsed in a background worker in 1,000-row batches. Matching also runs off the main thread, so large imports do not freeze the page.</p>
          {loading && <div className="mt-3 rounded-xl border border-indigo-100 bg-indigo-50/70 p-3" role="status" aria-live="polite">
            <div className="mb-2 flex items-center justify-between gap-3 text-xs font-semibold text-indigo-900"><span>{importStatus || "Importing records…"}</span><span>{importProgress}%</span></div>
            <div className="h-2 overflow-hidden rounded-full bg-indigo-100"><div className="h-full rounded-full bg-indigo-600 transition-[width] duration-300" style={{ width: `${importProgress}%` }} /></div>
          </div>}
          <label className="mt-4 flex cursor-pointer items-start gap-3 rounded-xl border border-indigo-100 bg-indigo-50/60 p-3">
            <input type="checkbox" checked={collapseRepeatedSubmissions} onChange={(e) => setCollapseRepeatedSubmissions(e.target.checked)} className="mt-1 accent-indigo-600" />
            <span>
              <span className="block text-sm font-semibold text-slate-800">Compare using the most complete submission per person, per file</span>
              <span className="mt-1 block text-xs leading-5 text-slate-600">When enabled, likely repeated rows within each file are grouped using name, graduation year, and department. The most complete row represents the group, and its blank fields are filled from the other submissions. All Records, Missing Information, Unmatched, filtered Excel/CSV exports, matching, and Union Master use this consolidated view. Exports include the original source rows and submission count. Turn it off to inspect every original row. This is a heuristic—review conflicts before relying on it.</span>
            </span>
          </label>
          {error && <p role="alert" className="mt-4 rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-700">{error}</p>}
        </section>

        <section className="mb-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
          <Stat label={collapseRepeatedSubmissions ? "Records after consolidation" : "Imported records"} value={collapseRepeatedSubmissions ? comparisonGroups.records.length : records.length} note={collapseRepeatedSubmissions ? `${records.length.toLocaleString()} original rows · duplicates grouped within each file` : "Original rows across all files"} />
          <Stat label="Source files" value={sources.length} note="Distinct uploaded files" />
          <Stat label="Match candidates" value={matches.length} note={matching ? "Matching in background…" : "Pairs to assess"} />
          <Stat label="Reviewed / auto-confirmed" value={Object.keys(decisions).length + (autoConfirmEnabled ? matches.filter((m) => m.autoConfirmEligible && !decisions[m.id]).length : 0)} note="Manual decisions + eligible automatic confirmations" />
        </section>

        {sources.length > 1 && <section className="mb-5 rounded-2xl border border-indigo-100 bg-white p-4 shadow-sm">
          <div className="mb-3"><h2 className="font-bold">Choose datasets to compare</h2><p className="mt-1 text-sm text-slate-500">Select the master dataset (D1) and the incoming dataset (D2). Match review and the relationship visualisation will focus on this pair.</p></div>
          <div className="grid gap-3 sm:grid-cols-2">
            <label className="text-sm font-medium">D1 · Master dataset<select value={masterSource} onChange={(e) => setMasterSource(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-200 bg-white px-3 py-2">{sources.map((source) => <option key={source} value={source}>{source}</option>)}</select></label>
            <label className="text-sm font-medium">D2 · Incoming dataset<select value={selectedD2Source} onChange={(e) => setSelectedD2Source(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-200 bg-white px-3 py-2">{sources.filter((source) => source !== masterSource).map((source) => <option key={source} value={source}>{source}</option>)}</select></label>
          </div>
          <p className="mt-2 text-xs text-slate-500">Currently comparing {comparisonRecords.length.toLocaleString()} records across the selected pair. Other imported files remain available in All Records but are not included in this comparison pass.</p>
        </section>}

        {records.length > 0 && (
          <section className="mb-4 rounded-xl border border-indigo-100 bg-indigo-50/70 px-4 py-3 text-sm text-indigo-950">
            <strong>{sources.length === 1 ? "Single-file inspection" : "Multi-file comparison"}</strong>
            <span className="ml-2 text-indigo-800">{sources.length === 1 ? "Review all imported columns, find blank fields, and export the data. Upload another file when you are ready to compare datasets." : `Comparing ${sources.length} files. Review candidate matches and unmatched records.`}</span>
          </section>
        )}

        <section className="mb-4 flex flex-wrap gap-2">
          {tabItems.map((item) => (
            <button key={item.id} onClick={() => setTab(item.id)}
              className={`${buttonBase} ${tab === item.id ? "bg-slate-900 text-white" : "border border-slate-200 bg-white text-slate-600 hover:bg-slate-100"}`}>
              {item.label} <span className="ml-1 opacity-70">{item.count}</span>
            </button>
          ))}
          <div className="flex-1" />
          <button className={buttonSecondary} onClick={() => setShowMapping((v) => !v)}>
            {showMapping ? "Hide column mapping" : "Map columns"}
          </button>
          <button className={buttonPrimary} onClick={() => setShowExportSettings(true)}>Export</button>
        </section>

        {showExportSettings && (
          <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto bg-slate-950/50 p-4 sm:p-8" onMouseDown={(e) => { if (e.target === e.currentTarget) setShowExportSettings(false); }}>
          <section role="dialog" aria-modal="true" aria-labelledby="export-settings-title" className="my-auto max-h-[92vh] w-full max-w-4xl overflow-y-auto rounded-2xl border border-slate-200 bg-white p-5 shadow-2xl sm:p-6">
            <div className="mb-4 flex items-center justify-between gap-3"><div><h2 id="export-settings-title" className="text-lg font-bold">Export settings</h2><p className="mt-1 text-sm text-slate-500">Choose what you want to export and configure the workbook.</p></div><button className={buttonSecondary} onClick={() => setShowExportSettings(false)}>Close</button></div>
            <div className="mb-5 flex flex-wrap gap-2 border-b border-slate-100 pb-5">
              <button className={buttonSecondary} onClick={() => exportRecords(filteredRecords, "filtered-records.xlsx")}>Export filtered records (Excel)</button>
              <button className={buttonSecondary} onClick={() => exportRows(filteredRecords.map((r) => { const members = comparisonGroups.membersByRepresentative.get(r.id) ?? [r]; return { "Source File": r.source, "Representative Source Row": r.rowNumber, "Repeated Submissions Grouped": members.length > 1 ? "Yes" : "No", "Submission Count": members.length, "Source Rows In Group": members.map((member) => `${member.source}: ${member.rowNumber}`).join("; "), ...r.data }; }), "filtered-records.csv")}>Export filtered records (CSV)</button>
              <button className={buttonSecondary} onClick={exportMatches}>Export matches</button>
              <button className={buttonSecondary} onClick={() => exportRecords(unmatchedRecords, "unmatched-records.xlsx")}>Export unmatched</button>
            </div>
            {records.length > 0 && sources.length > 1 && <>

            <div className="flex flex-wrap items-center justify-between gap-3">
              <div><h2 className="text-lg font-bold">Export Master Configuration</h2><p className="mt-1 text-sm text-slate-500">Choose which datasets, fields and review worksheets to include in your export.</p></div>
              <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-semibold text-emerald-700">Master-based</span>
            </div>
            <div className="mt-4 grid gap-4 md:grid-cols-2">
              <label className="text-sm font-medium">Master dataset (D1)
                <select value={masterSource} onChange={(e) => setMasterSource(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-200 bg-white px-3 py-2">
                  {sources.map((source) => <option key={source} value={source}>{source}</option>)}
                </select>
              </label>
              <label className="text-sm font-medium">Incoming dataset (D2)<select value={selectedD2Source} onChange={(e) => setSelectedD2Source(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-200 bg-white px-3 py-2">{sources.filter((source) => source !== masterSource).map((source) => <option key={source} value={source}>{source}</option>)}</select></label>
            </div>
            <div className="mt-5">
              <h3 className="font-semibold">Choose dataset composition</h3>
              <div className="mt-2 grid gap-2 lg:grid-cols-3">
                <label className={`cursor-pointer rounded-xl border p-3 ${exportComposition === "union" ? "border-indigo-400 bg-indigo-50" : "border-slate-200"}`}><input type="radio" name="exportComposition" checked={exportComposition === "union"} onChange={() => setExportComposition("union")} className="mr-2 accent-indigo-600"/><span className="font-semibold">D1 + Unmatched D2 + Approved Enrichment</span><span className="mt-1 block text-xs text-slate-500">Adds D2-only individuals that are not pending a possible match.</span></label>
                <label className={`cursor-pointer rounded-xl border p-3 ${exportComposition === "enriched" ? "border-indigo-400 bg-indigo-50" : "border-slate-200"}`}><input type="radio" name="exportComposition" checked={exportComposition === "enriched"} onChange={() => setExportComposition("enriched")} className="mr-2 accent-indigo-600"/><span className="font-semibold">D1 + Approved Enrichment</span><span className="mt-1 block text-xs text-slate-500">Keeps D1's record count; excludes all unmatched D2 individuals.</span></label>
                <label className={`cursor-pointer rounded-xl border p-3 ${exportComposition === "original" ? "border-indigo-400 bg-indigo-50" : "border-slate-200"}`}><input type="radio" name="exportComposition" checked={exportComposition === "original"} onChange={() => setExportComposition("original")} className="mr-2 accent-indigo-600"/><span className="font-semibold">D1 Only</span><span className="mt-1 block text-xs text-slate-500">Exports original master values without D2 enrichment.</span></label>
              </div>
            </div>
            <div className="mt-5 border-t border-slate-100 pt-4">
              <div className="flex flex-wrap items-center justify-between gap-2"><h3 className="font-semibold">Select fields</h3><div className="flex gap-3 text-xs"><button onClick={() => setSelectedExportFields(columns)} className="font-semibold text-indigo-700">Select all</button><button onClick={() => setSelectedExportFields([])} className="font-semibold text-slate-500">Clear</button></div></div>
              <p className="mt-1 text-xs text-slate-500">{selectedExportFields.length} of {columns.length} fields selected</p>
              <div className="mt-3 grid gap-2 sm:grid-cols-2 lg:grid-cols-3">{columns.map((field) => <label key={field} className="flex items-center gap-2 text-sm"><input type="checkbox" checked={selectedExportFields.includes(field)} onChange={(e) => setSelectedExportFields((old) => e.target.checked ? [...old, field] : old.filter((item) => item !== field))} className="accent-indigo-600"/>{field}</label>)}</div>
            </div>
            <div className="mt-5 border-t border-slate-100 pt-4">
              <h3 className="font-semibold">Additional export options</h3>
              <div className="mt-3 grid gap-2 sm:grid-cols-2">
                <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includeOriginalValues} onChange={(e) => setIncludeOriginalValues(e.target.checked)} className="accent-indigo-600"/>Include original D1 values in separate columns</label>
                <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includeEnrichmentStatus} onChange={(e) => setIncludeEnrichmentStatus(e.target.checked)} className="accent-indigo-600"/>Include enrichment and verification status</label>
                <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includeMatchedSheet} onChange={(e) => setIncludeMatchedSheet(e.target.checked)} className="accent-indigo-600"/>Include confirmed matches worksheet</label>
                <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includeReviewSheets} onChange={(e) => setIncludeReviewSheets(e.target.checked)} className="accent-indigo-600"/>Include pending Match Review worksheet</label>
                <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includeRejectedSheet} onChange={(e) => setIncludeRejectedSheet(e.target.checked)} className="accent-indigo-600"/>Include rejected matches worksheet</label>
              </div>
            </div>
            <div className="mt-5 rounded-xl bg-slate-50 p-4">
              <h3 className="font-semibold">Export preview</h3>
              <div className="mt-2 space-y-2 text-sm"><div className="flex justify-between"><span className="text-slate-500">Master records</span><strong>{mappedRecords.filter((record) => record.source === masterSource).length.toLocaleString()}</strong></div><div className="flex justify-between"><span className="text-slate-500">Approved cross-file matches</span><strong>{matches.filter((match) => (effectiveDecision(match) === "Confirmed same person" || effectiveDecision(match) === "Auto-confirmed same person") && (match.left.source === masterSource || match.right.source === masterSource)).length.toLocaleString()}</strong></div><div className="flex justify-between"><span className="text-slate-500">Pending candidates in review sheet</span><strong>{matches.filter((match) => !decisions[match.id] && !(autoConfirmEnabled && match.autoConfirmEligible)).length.toLocaleString()}</strong></div><div className="flex justify-between"><span className="text-slate-500">Unmatched D2 included in master</span><strong>{exportComposition === "union" ? "Yes" : "No"}</strong></div><div className="flex justify-between"><span className="text-slate-500">Selected fields</span><strong>{selectedExportFields.length}</strong></div></div>
            </div>
            <button disabled={!selectedExportFields.length || !masterSource} onClick={() => { exportConfiguredMaster(); setShowExportSettings(false); }} className={`${buttonPrimary} mt-4 w-full disabled:cursor-not-allowed disabled:opacity-50`}>Export configured Excel workbook</button>
            <p className="mt-2 text-xs text-slate-500">Pending matches are exported to a separate worksheet when selected; they never enrich Final_Master. Only confirmed or eligible auto-confirmed matches enrich D1.</p>
            </>}
          </section>
          </div>
        )}

        {showMapping && (
          <section className="mb-4 rounded-2xl border bg-white p-5">
            <div className="mb-3">
              <h2 className="font-bold">Column mapping</h2>
              <p className="mt-1 text-sm text-slate-500">Map source headers to common fields. Changes update matching and displayed data.</p>
            </div>
            {columns.length === 0 ? <p className="text-sm text-slate-500">Import a file first.</p> : (
              <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
                {sourceColumns.map((col) => (
                  <label key={col} className="block text-sm">
                    <span className="mb-1 block font-medium text-slate-700">{col}</span>
                    <select value={mapping[col] ?? canonicalHeader(col)} onChange={(e) => setMapping((old) => ({ ...old, [col]: e.target.value }))}
                      className="w-full rounded-lg border border-slate-200 bg-white px-3 py-2">
                      <option value={col}>{`Keep as "${col}"`}</option>
                      {standardFields.map((field) => <option key={field} value={field}>{field}</option>)}
                    </select>
                  </label>
                ))}
              </div>
            )}
            <button className={`${buttonSecondary} mt-4`} onClick={() => setMapping({})}>Reset mapping</button>
          </section>
        )}

        <section className="mb-4 grid gap-3 rounded-2xl border bg-white p-4 sm:grid-cols-[minmax(240px,1fr)_220px_auto] sm:items-center">
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search names, departments, IDs…"
            className="w-full rounded-xl border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100" />
          <select value={sourceFilter} onChange={(e) => setSourceFilter(e.target.value)}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm">
            <option>All files</option>{sources.map((source) => <option key={source}>{source}</option>)}
          </select>
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={missingOnly} onChange={(e) => setMissingOnly(e.target.checked)} className="accent-indigo-600" />
            Has missing fields
          </label>
        </section>

        {tab === "matches" && (
          <section className="mb-4 rounded-2xl border bg-white p-5">
            <div className="mb-5 rounded-xl border border-slate-200 bg-slate-50 p-4">
              <div className="flex flex-wrap items-center justify-between gap-2"><div><h3 className="font-semibold">Fields to compare</h3><p className="mt-1 text-xs text-slate-500">Only checked fields are used by the existing matching algorithm.</p></div><div className="flex gap-3 text-xs"><button onClick={() => setSelectedCompareFields(MATCH_FIELDS.map((field) => field.name))} className="font-semibold text-indigo-700">Select all</button><button onClick={() => setSelectedCompareFields([])} className="font-semibold text-slate-500">Clear</button></div></div>
              <div className="mt-3 grid gap-2 sm:grid-cols-2 lg:grid-cols-4">{MATCH_FIELDS.map((field) => <label key={field.name} className="flex items-center gap-2 text-sm"><input type="checkbox" checked={selectedCompareFields.includes(field.name)} onChange={(e) => setSelectedCompareFields((old) => e.target.checked ? [...old, field.name] : old.filter((name) => name !== field.name))} className="accent-indigo-600" />{field.name}{field.identifier && <span className="text-[10px] text-slate-400">ID</span>}</label>)}</div>
              <p className="mt-2 text-xs text-slate-500">{selectedCompareFields.length} of {MATCH_FIELDS.length} comparison fields selected. Sl. No. is not a comparison field.</p>
              {!selectedCompareFields.length && <p className="mt-2 text-xs font-semibold text-amber-700">Select at least one field to run comparison.</p>}
            </div>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div><h2 className="font-bold">Match sensitivity</h2><p className="mt-1 text-sm text-slate-500">Raise the threshold to show stronger candidates only.</p></div>
              <span className="rounded-lg bg-indigo-50 px-3 py-1.5 text-sm font-bold text-indigo-700">{threshold}%</span>
            </div>
            <input type="range" min={30} max={95} step={5} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))}
              className="mt-4 w-full accent-indigo-600" />
            <div className="mt-1 flex justify-between text-xs text-slate-400"><span>More candidates</span><span>Stricter matches</span></div>
            <div className="mt-4 rounded-xl border border-emerald-200 bg-emerald-50/70 p-3">
              <label className="flex cursor-pointer items-start gap-3">
                <input type="checkbox" checked={autoConfirmEnabled} onChange={(e) => setAutoConfirmEnabled(e.target.checked)} className="mt-1 accent-emerald-600" />
                <span>
                  <span className="block text-sm font-semibold text-emerald-950">Automatically confirm unique high-confidence matches</span>
                  <span className="mt-1 block text-xs leading-5 text-emerald-800">Only cross-file pairs with exact normalised Name + Roll No + Graduation Year, no conflicting email/mobile values, and a unique one-to-one candidate are eligible. All other candidates remain in Match Review. This is optional and OFF by default.</span>
                  <span className="mt-1 block text-xs font-semibold text-emerald-900">{matches.filter((m) => m.autoConfirmEligible).length.toLocaleString()} candidates currently eligible.</span>
                </span>
              </label>
            </div>
          </section>
        )}

        {tab === "records" && <RecordTable records={filteredRecords} columns={columns} pageSize={100} />}
        {tab === "matches" && (
          <section className="space-y-3">
            <div className="rounded-2xl border border-indigo-100 bg-white p-4 shadow-sm">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="font-bold">Bulk approve match candidates</h2>
                  <p className="mt-1 text-sm text-slate-500">Select candidates across all result pages, then approve them together. Cross-file approvals feed Approved Enrichment; within-file approvals confirm repeated submissions in that source.</p>
                </div>
                <span className="rounded-full bg-indigo-50 px-3 py-1.5 text-sm font-semibold text-indigo-700">{selectedPendingMatches.length.toLocaleString()} selected</span>
              </div>
              <div className="mt-3 flex flex-wrap items-center gap-3 border-t border-slate-100 pt-3">
                <label className="flex cursor-pointer items-center gap-2 text-sm font-medium text-slate-700">
                  <input type="checkbox" checked={allPendingSelected} onChange={(e) => e.target.checked ? selectAllPendingMatches() : setSelectedMatchIds([])} className="h-4 w-4 accent-indigo-600" />
                  Select all {pendingMatches.length.toLocaleString()} pending candidates (all pages)
                </label>
                <button className={buttonSecondary} onClick={() => setSelectedMatchIds([])} disabled={!selectedMatchIds.length}>Clear selection</button>
                <button className={`${buttonBase} bg-emerald-600 text-white hover:bg-emerald-700 disabled:cursor-not-allowed disabled:opacity-40`} onClick={approveSelectedMatches} disabled={!selectedPendingMatches.length || matching}>Bulk approve ({selectedPendingMatches.length.toLocaleString()})</button>
                <button className={`${buttonBase} bg-rose-600 text-white hover:bg-rose-700 disabled:cursor-not-allowed disabled:opacity-40`} onClick={rejectSelectedMatches} disabled={!selectedPendingMatches.length || matching}>Bulk not a match ({selectedPendingMatches.length.toLocaleString()})</button>
              </div>
              <div className="mt-2 flex flex-wrap gap-x-5 gap-y-1 text-xs text-slate-500">
                <span>{selectedPendingMatches.filter((m) => m.left.source !== m.right.source).length.toLocaleString()} cross-file → Approved Enrichment</span>
                <span>{selectedPendingMatches.filter((m) => m.left.source === m.right.source).length.toLocaleString()} within-file → duplicate consolidation</span>
                <span>{approvedMatches.length.toLocaleString()} already approved/auto-confirmed</span>
              </div>
            </div>
            {matching && <div className="rounded-xl border border-indigo-100 bg-indigo-50 p-4 text-sm text-indigo-800">Checking {comparisonRecords.length.toLocaleString()} comparison records for within-file duplicates and cross-file matches in the background. You can keep using the interface; results appear when the pass completes.</div>}
            {!matching && pendingMatches.length > 200 && <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">Showing the first 200 of {pendingMatches.length.toLocaleString()} pending candidates to keep the page responsive. Export matches to review the full candidate report, or narrow the list with search and filters.</div>}
            {pendingMatches.length === 0 ? <Empty title="No pending candidates" detail={sources.length < 1 ? "Import an Excel or CSV file to find possible duplicates." : "No candidates found. Check the column mapping or lower the threshold. Matching checks both duplicates within a file and common records across files."} /> :
              pendingMatches.map((m) => <article key={m.id} className="rounded-2xl border bg-white p-5 shadow-sm">
                <div className="mb-3 flex items-center justify-between border-b border-slate-100 pb-3">
                  <label className="flex cursor-pointer items-center gap-2 text-sm font-medium text-slate-600"><input type="checkbox" checked={selectedMatchIds.includes(m.id)} onChange={(e) => toggleMatchSelection(m.id, e.target.checked)} className="h-4 w-4 accent-indigo-600" />Select candidate</label>
                  <span className="text-xs text-slate-400">{m.left.source === m.right.source ? "Within-file candidate" : "Cross-file candidate"}</span>
                </div>
                <div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-start">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="rounded-full bg-indigo-50 px-3 py-1 text-sm font-bold text-indigo-700">{m.score}% score</span>
                      <span className={`rounded-full px-3 py-1 text-xs font-semibold ${m.status === "High confidence" ? "bg-emerald-50 text-emerald-700" : m.status === "Likely match" ? "bg-blue-50 text-blue-700" : "bg-amber-50 text-amber-700"}`}>{m.status}</span>
                      {autoConfirmEnabled && m.autoConfirmEligible && <span className="rounded-full bg-emerald-100 px-2.5 py-1 text-xs font-semibold text-emerald-800">Auto-confirmed</span>}
                    </div>
                    <h3 className="mt-3 text-lg font-bold">{m.left.data.Name || "Record A"} <span className="font-normal text-slate-400">↔</span> {m.right.data.Name || "Record B"}</h3>
                    <p className="mt-1 break-words text-sm text-slate-500">{m.left.source} · row {m.left.rowNumber} <span className="mx-1">/</span> {m.right.source} · row {m.right.rowNumber}</p>
                    <div className="mt-3 flex flex-wrap gap-2">{m.reasons.map((reason) => <span key={reason} className="rounded-lg bg-slate-100 px-2.5 py-1 text-xs text-slate-600">{reason}</span>)}</div>
                  </div>
                  <button className={buttonSecondary} onClick={() => setSelectedMatch(m.id)}>Compare records</button>
                </div>
                <div className="mt-4 flex flex-wrap gap-2 border-t border-slate-100 pt-4">
                  <button className={`${buttonBase} bg-emerald-50 text-emerald-800 hover:bg-emerald-100`} onClick={() => setDecisions((d) => ({ ...d, [m.id]: "Confirmed same person" }))}>Confirm match</button>
                  <button className={`${buttonBase} bg-rose-50 text-rose-800 hover:bg-rose-100`} onClick={() => setDecisions((d) => ({ ...d, [m.id]: "Not a match" }))}>Not a match</button>
                  <button className={`${buttonBase} border border-slate-200 bg-slate-50 text-slate-700 hover:bg-slate-100`} onClick={() => setDecisions((d) => ({ ...d, [m.id]: "Kept unchanged" }))}>Keep it same</button>
                </div>
              </article>)}
          </section>
        )}
        {tab === "missing" && <RecordTable records={missingRecords} columns={columns} pageSize={100} />}
        {tab === "unmatched" && <RecordTable records={unmatchedRecords} columns={columns} pageSize={100} />}

        {activeMatch && <CompareModal
          match={activeMatch}
          onClose={() => setSelectedMatch(null)}
          onSave={(leftData, rightData) => {
            setRecords((current) => current.map((record) => {
              if (record.id !== activeMatch.left.id && record.id !== activeMatch.right.id) return record;
              const edited = record.id === activeMatch.left.id ? leftData : rightData;
              const nextData = { ...record.data };
              for (const [field, value] of Object.entries(edited)) {
                const originalKey = Object.keys(nextData).find((key) => (mapping[key] || canonicalHeader(key)) === field);
                if (originalKey) nextData[originalKey] = value;
                else nextData[field] = value;
              }
              return { ...record, data: nextData };
            }));
            setSelectedMatch(null);
          }}
        />}
        <footer className="mt-8 flex flex-col gap-1 border-t border-slate-200 pt-5 text-xs leading-5 text-slate-400 sm:flex-row sm:justify-between">
          <span>Excel Data Matcher · Standalone local-first tool</span>
          <span>Similarity scores are heuristics. Automatic confirmation is restricted to unique exact-identifier matches.</span>
        </footer>
      </div>
    </main>
  );
}

function Stat({ label, value, note }: { label: string; value: number; note: string }) {
  return <div className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
    <p className="text-sm font-medium text-slate-500">{label}</p><p className="mt-2 text-3xl font-bold tracking-tight">{value.toLocaleString()}</p><p className="mt-1 text-xs text-slate-400">{note}</p>
  </div>;
}

function Empty({ title, detail }: { title: string; detail: string }) {
  return <div className="rounded-2xl border bg-white px-5 py-14 text-center"><h2 className="font-bold">{title}</h2><p className="mx-auto mt-2 max-w-lg text-sm text-slate-500">{detail}</p></div>;
}


function VisualizationWorkspace({ masterSource, d2Source, sources, onD1Change, onD2Change, onBack, d1Count, d2Count, candidateCount, confirmedCount, pendingCount, rejectedCount, unmatchedD1, unmatchedD2, relationships }: {
  masterSource: string; d2Source: string; sources: string[]; onD1Change: (value: string) => void; onD2Change: (value: string) => void; onBack: () => void;
  d1Count: number; d2Count: number; candidateCount: number; confirmedCount: number; pendingCount: number; rejectedCount: number; unmatchedD1: number; unmatchedD2: number;
  relationships: Array<Omit<ReturnType<typeof compareRecords>[number], "status"> & { status: string }>;
}) {
  const [mode, setMode] = useState<"all" | "confirmed" | "pending" | "rejected">("all");
  const [inspected, setInspected] = useState<(typeof relationships)[number] | null>(null);
  const visible = relationships.filter((r) => mode === "all" || (mode === "confirmed" ? r.status === "Confirmed same person" || r.status === "Auto-confirmed same person" : mode === "rejected" ? r.status === "Not a match" : r.status !== "Confirmed same person" && r.status !== "Auto-confirmed same person" && r.status !== "Not a match"));
  const statusColor = (status: string) => status === "Confirmed same person" || status === "Auto-confirmed same person" ? "#baffd5" : status === "Not a match" ? "#ff7185" : "#8dd6ae";
  const shortName = (r: (typeof relationships)[number], side: "left" | "right") => String(r[side].data.Name ?? r[side].data["Full Name"] ?? r[side].data.name ?? "Unnamed record");
  return <main className="fixed inset-0 z-50 flex h-[100dvh] min-h-0 flex-col overflow-hidden bg-[#07110d] text-white">
    <header className="z-20 flex shrink-0 flex-wrap items-center justify-between gap-3 border-b border-emerald-100/10 bg-[#07110d]/90 px-4 py-3 backdrop-blur-xl sm:px-7">
      <div className="flex items-center gap-3"><button onClick={onBack} className="rounded-lg border border-emerald-100/15 bg-white/[.035] px-3 py-2 text-xs font-semibold text-white/80 transition hover:bg-white/10">← Workspace</button><div><div className="text-[9px] font-semibold uppercase tracking-[.25em] text-emerald-300/80">Three-dimensional data observatory</div><h1 className="text-lg font-light tracking-[.12em]">D1 <span className="text-emerald-300">↔</span> D2 RELATIONSHIP FIELD</h1></div></div>
      <div className="flex flex-wrap items-center gap-2 rounded-xl border border-emerald-100/10 bg-black/30 p-2"><label className="flex items-center gap-2 text-[10px] uppercase tracking-wider text-white/45">D1<select value={masterSource} onChange={e=>{onD1Change(e.target.value);setInspected(null)}} className="max-w-40 rounded-md border border-white/10 bg-[#101b15] px-2 py-1.5 text-xs normal-case tracking-normal text-white">{sources.map(s=><option key={s} value={s}>{s}</option>)}</select></label><label className="flex items-center gap-2 text-[10px] uppercase tracking-wider text-white/45">D2<select value={d2Source} onChange={e=>{onD2Change(e.target.value);setInspected(null)}} className="max-w-40 rounded-md border border-white/10 bg-[#101b15] px-2 py-1.5 text-xs normal-case tracking-normal text-white">{sources.filter(s=>s!==masterSource).map(s=><option key={s} value={s}>{s}</option>)}</select></label></div>
    </header>
    <section className="relative min-h-0 flex-1 overflow-hidden bg-[radial-gradient(ellipse_at_50%_48%,#14251b_0%,#0a1510_48%,#050b08_100%)]">
      <div className="pointer-events-none absolute inset-0 z-[1] opacity-[.09]" style={{backgroundImage:"linear-gradient(rgba(184,232,201,.18) 1px,transparent 1px),linear-gradient(90deg,rgba(184,232,201,.18) 1px,transparent 1px)",backgroundSize:"64px 64px"}} />
      <RelationshipScene relationships={relationships} mode={mode} onInspect={setInspected} />
      <div className="pointer-events-none absolute left-5 top-5 z-10 rounded-xl border border-emerald-100/10 bg-black/35 px-4 py-3 backdrop-blur-md"><div className="text-[9px] uppercase tracking-[.22em] text-emerald-100/45">Relationship field</div><div className="mt-1 text-sm font-medium">{visible.length.toLocaleString()} connections</div><div className="mt-1 text-[10px] text-white/35">Rendering up to 3,500 actual links</div></div>
      <div className="absolute right-5 top-5 z-10 flex flex-wrap justify-end gap-2">{([['all','All'],['confirmed','Confirmed'],['pending','Pending'],['rejected','Rejected']] as const).map(([v,l])=><button key={v} onClick={()=>{setMode(v);setInspected(null)}} className={`rounded-lg border px-3 py-2 text-[10px] font-semibold transition ${mode===v?'border-emerald-100/70 bg-emerald-100 text-[#07110d]':'border-emerald-100/15 bg-black/40 text-white/65 hover:bg-white/10'}`}>{l}</button>)}</div>
      <div className="pointer-events-none absolute bottom-5 left-5 z-10 hidden items-center gap-4 rounded-lg border border-white/10 bg-black/30 px-3 py-2 text-[9px] uppercase tracking-wider text-white/55 backdrop-blur sm:flex"><span><i className="mr-1 inline-block h-2 w-2 rounded-full bg-emerald-200"/>Confirmed</span><span><i className="mr-1 inline-block h-2 w-2 rounded-full bg-emerald-500"/>Pending</span><span><i className="mr-1 inline-block h-2 w-2 rounded-full bg-rose-400"/>Rejected</span></div>
      {inspected&&<div className="absolute bottom-5 right-5 z-20 w-[min(92vw,440px)] rounded-xl border border-emerald-100/15 bg-[#0b1711]/95 p-4 shadow-2xl backdrop-blur-xl"><div className="flex items-center justify-between gap-3"><span className="text-[9px] uppercase tracking-[.2em] text-white/45">Relationship inspector</span><button onClick={()=>setInspected(null)} className="rounded-md border border-white/10 px-2 py-1 text-[10px] text-white/55 hover:bg-white/10">Close</button></div><div className="mt-2 flex items-center justify-between"><span className="text-xs font-semibold" style={{color:statusColor(inspected.status)}}>{inspected.status}</span><span className="text-xs text-white/55">{inspected.score}% similarity</span></div><div className="mt-3 grid grid-cols-2 gap-3"><div className="rounded-lg border border-white/10 bg-white/[.03] p-3"><div className="text-[9px] uppercase tracking-wider text-emerald-200/50">D1 · {masterSource}</div><div className="mt-2 text-xs font-medium">{shortName(inspected,'left')}</div><div className="mt-1 text-[10px] text-white/40">Row {inspected.left.rowNumber}</div></div><div className="rounded-lg border border-white/10 bg-white/[.03] p-3"><div className="text-[9px] uppercase tracking-wider text-emerald-200/50">D2 · {d2Source}</div><div className="mt-2 text-xs font-medium">{shortName(inspected,'right')}</div><div className="mt-1 text-[10px] text-white/40">Row {inspected.right.rowNumber}</div></div></div><div className="mt-3 text-[10px] text-white/35">This connection represents an actual comparison result. Review or change its decision in Match Review.</div></div>}
      {visible.length===0&&<div className="pointer-events-none absolute inset-0 z-10 flex items-center justify-center"><div className="rounded-2xl border border-emerald-100/10 bg-black/40 px-8 py-6 text-center"><div className="text-sm font-semibold">No relationships in this view</div><p className="mt-2 text-xs text-white/45">Try another filter or select a different D1 and D2 pair.</p></div></div>}
    </section>
    <section className="z-20 grid shrink-0 grid-cols-2 gap-px border-t border-emerald-100/10 bg-emerald-100/10 sm:grid-cols-3 lg:grid-cols-6">{[['D1 RECORDS',d1Count,'#b8e8c9'],['D2 RECORDS',d2Count,'#b8e8c9'],['CANDIDATE LINKS',candidateCount,'#e5f4e9'],['CONFIRMED',confirmedCount,'#baffd5'],['PENDING REVIEW',pendingCount,'#8dd6ae'],['MATCH RATE',`${candidateCount?Math.round(confirmedCount/candidateCount*100):0}%`,'#baffd5']].map(([label,value,color])=><div key={String(label)} className="min-h-[70px] bg-[#09140e] px-4 py-3 sm:px-5"><div className="text-[9px] font-medium tracking-[.16em] text-white/40">{label}</div><div className="mt-1 text-xl font-semibold tabular-nums" style={{color:String(color)}}>{Number.isFinite(Number(value))?Number(value).toLocaleString():value}</div></div>)}</section>
    <footer className="flex shrink-0 items-center justify-between gap-3 border-t border-emerald-100/10 bg-[#07110d] px-4 py-2 text-[9px] text-white/35 sm:px-6"><span>THREE.JS DATA OBSERVATORY / D1 ↔ D2</span><span>{unmatchedD1.toLocaleString()} unmatched D1 · {unmatchedD2.toLocaleString()} unmatched D2 · {rejectedCount.toLocaleString()} rejected</span></footer>
  </main>;
}

function RecordTable({ records, columns, pageSize = 100 }: { records: StudentRecord[]; columns: string[]; pageSize?: number }) {
  const [page, setPage] = useState(0);
  useEffect(() => setPage(0), [records]);
  const pageCount = Math.max(1, Math.ceil(records.length / pageSize));
  const visibleRecords = records.slice(page * pageSize, (page + 1) * pageSize);
  if (!records.length) return <Empty title="No records to show" detail="Try changing the filters or import a spreadsheet." />;
  return <section className="overflow-hidden rounded-2xl border bg-white shadow-sm">
    <div className="overflow-auto">
      <table className="w-full border-collapse text-left text-sm">
        <thead className="bg-slate-100 text-xs uppercase tracking-wide text-slate-500"><tr>
          <th className="whitespace-nowrap px-4 py-3">Source file</th><th className="px-4 py-3">Row</th>
          {columns.map((c) => <th key={c} className="whitespace-nowrap px-4 py-3">{c}</th>)}
        </tr></thead>
        <tbody>{visibleRecords.map((r) => <tr key={r.id} className="border-t border-slate-100 hover:bg-slate-50">
          <td className="whitespace-nowrap px-4 py-3 text-slate-500">{r.source}</td><td className="px-4 py-3 text-slate-500">{r.rowNumber}</td>
          {columns.map((c) => <td key={c} className="max-w-72 whitespace-nowrap px-4 py-3">
            {r.data[c] ? <span>{r.data[c]}</span> : <span className="rounded-md bg-amber-50 px-2 py-1 text-xs text-amber-700">Missing</span>}
          </td>)}
        </tr>)}</tbody>
      </table>
    </div>
    <div className="flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 px-4 py-3 text-xs text-slate-500">
      <span>Showing {(page * pageSize + 1).toLocaleString()}–{Math.min((page + 1) * pageSize, records.length).toLocaleString()} of {records.length.toLocaleString()} records (100 per page).</span>
      <div className="flex items-center gap-2">
        <button className="rounded-lg border px-3 py-1.5 disabled:opacity-40" disabled={page === 0} onClick={() => setPage((p) => Math.max(0, p - 1))}>Previous</button>
        <span>Page {page + 1} of {pageCount}</span>
        <button className="rounded-lg border px-3 py-1.5 disabled:opacity-40" disabled={page >= pageCount - 1} onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))}>Next</button>
      </div>
    </div>
  </section>;
}

function CompareModal({
  match,
  onClose,
  onSave
}: {
  match: ReturnType<typeof compareRecords>[number];
  onClose: () => void;
  onSave: (leftData: Record<string, string>, rightData: Record<string, string>) => void;
}) {
  const fields = Array.from(new Set([...Object.keys(match.left.data), ...Object.keys(match.right.data)]));
  const [leftData, setLeftData] = useState<Record<string, string>>({ ...match.left.data });
  const [rightData, setRightData] = useState<Record<string, string>>({ ...match.right.data });
  const [savedNotice, setSavedNotice] = useState("");

  function update(side: "left" | "right", field: string, value: string) {
    if (side === "left") setLeftData((old) => ({ ...old, [field]: value }));
    else setRightData((old) => ({ ...old, [field]: value }));
  }

  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-3 sm:p-6" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
    <section role="dialog" aria-modal="true" aria-labelledby="compare-title" className="max-h-[90vh] w-full max-w-6xl overflow-auto rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <h2 id="compare-title" className="text-xl font-bold">Compare and edit records</h2>
          <p className="mt-1 text-sm text-slate-500">{match.score}% similarity · {match.reasons.join(" · ")}</p>
          <p className="mt-1 text-xs text-indigo-600">Both records are editable. Changes are saved independently; neither record is merged into the other.</p>
        </div>
        <button onClick={onClose} className={buttonSecondary}>Close</button>
      </div>
      <div className="mt-5 grid gap-4 md:grid-cols-2">
        {[
          { label: "Record A", record: match.left, data: leftData, side: "left" as const },
          { label: "Record B", record: match.right, data: rightData, side: "right" as const }
        ].map(({ label, record, data, side }) => <div key={record.id} className="rounded-xl border border-slate-200 p-4">
          <h3 className="font-bold">{label} · Editable</h3>
          <p className="mt-1 break-all text-xs text-slate-500">{record.source} · row {record.rowNumber}</p>
          {fields.map((field) => {
            const a = normalizeValue(leftData[field] || "");
            const b = normalizeValue(rightData[field] || "");
            const different = a !== b;
            const equivalentAfterNormalization = different && Boolean(a && b) &&
              comparisonValue(field, a) === comparisonValue(field, b);
            const realConflict = different && Boolean(a && b) && !equivalentAfterNormalization;
            const tone = equivalentAfterNormalization ? "bg-emerald-50" : realConflict ? "bg-amber-50" : "bg-slate-50";
            const detail = equivalentAfterNormalization
              ? " · equivalent after normalization"
              : realConflict
                ? " · differs"
                : (!a || !b) && different
                  ? " · missing on one side"
                  : "";
            return <div key={field} className={`mt-3 rounded-lg p-2 ${tone}`}>
              <label className="block text-xs font-semibold text-slate-600">{field}{detail}</label>
              <textarea
                value={data[field] ?? ""}
                onChange={(e) => update(side, field, e.target.value)}
                rows={Math.max(1, Math.min(3, Math.ceil(String(data[field] ?? "").length / 48)))}
                className="mt-1 w-full resize-y rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-900 outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100"
                placeholder={`Enter ${field}...`}
              />
            </div>;
          })}
        </div>)}
      </div>
      <div className="mt-5 flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-4">
        <p className="max-w-2xl text-xs leading-5 text-slate-400">Edits to A and B are saved to their own source rows. Original records are not merged or deleted. Similarity scores are only suggestions, not proof of identity.</p>
        <div className="flex gap-2">
          <button onClick={onClose} className={buttonSecondary}>Cancel</button>
          <button onClick={() => { onSave(leftData, rightData); setSavedNotice("Changes saved."); }} className={buttonPrimary}>Save both records</button>
        </div>
      </div>
      {savedNotice && <p className="mt-2 text-sm font-medium text-emerald-700">{savedNotice}</p>}
    </section>
  </div>;
}

"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  canonicalHeader,
  compareRecords,
  normalizeValue,
  standardFields,
  type StudentRecord
} from "@/lib/matcher";

type Tab = "records" | "matches" | "missing" | "unmatched";
type ReviewDecision = "Confirmed same person" | "Not a match";

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
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [matches, setMatches] = useState<ReturnType<typeof compareRecords>>([]);
  const [matching, setMatching] = useState(false);
  const [selectedMatch, setSelectedMatch] = useState<string | null>(null);
  const [decisions, setDecisions] = useState<Record<string, ReviewDecision>>({});
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [showMapping, setShowMapping] = useState(false);
  const [collapseRepeatedSubmissions, setCollapseRepeatedSubmissions] = useState(false);

  const sourceColumns = useMemo(() => Array.from(new Set(records.flatMap((r) => Object.keys(r.data)))), [records]);
  const sources = useMemo(() => Array.from(new Set(records.map((r) => r.source))), [records]);
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
      const department = compactKey(record.data.Department ?? "");
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

  // Run the expensive matching pass off the main UI thread. This keeps search,
  // scrolling, and the rest of the interface responsive on large workbooks.
  useEffect(() => {
    if (comparisonGroups.records.length < 2) {
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
    worker.postMessage({ records: comparisonGroups.records, threshold });
    return () => worker.terminate();
  }, [comparisonGroups, threshold, sources.length]);
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

  const pendingMatches = useMemo(() => matches.filter((m) => !decisions[m.id]).filter((m) => {
    const text = [m.left.source, m.right.source, ...Object.values(m.left.data), ...Object.values(m.right.data)].join(" ").toLowerCase();
    return text.includes(query.toLowerCase()) &&
      (sourceFilter === "All files" || m.left.source === sourceFilter || m.right.source === sourceFilter);
  }), [matches, decisions, query, sourceFilter]);

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
    try {
      const imported: StudentRecord[] = [];
      for (const file of Array.from(fileList)) {
        const ext = file.name.split(".").pop()?.toLowerCase();
        if (!["xlsx", "xls", "csv"].includes(ext ?? "")) throw new Error(`${file.name} is not an Excel or CSV file.`);
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;
        if (!sheet) continue;
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: false });
        rows.forEach((row, index) => {
          const data: Record<string, string> = {};
          for (const [header, value] of Object.entries(row)) data[header.trim() || "Unnamed column"] = normalizeValue(value);
          imported.push({
            id: `${file.name}:${index + 2}:${crypto.randomUUID()}`,
            source: file.name,
            rowNumber: index + 2,
            data
          });
        });
      }
      if (!imported.length) throw new Error("No data rows found. Check that the first worksheet has a header row and records.");
      setRecords((old) => [...old, ...imported]);
      setTab("records");
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not read these files.");
    } finally {
      setLoading(false);
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
      "Review Decision": decisions[m.id] || "Pending review",
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
    const confirmedMatches = matches.filter((match) => decisions[match.id] === "Confirmed same person");
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

  function resetWorkspace() {
    setRecords([]); setQuery(""); setSourceFilter("All files"); setMissingOnly(false);
    setDecisions({}); setSelectedMatch(null); setMapping({}); setError(""); setTab("records"); setCollapseRepeatedSubmissions(false);
  }

  const tabItems: { id: Tab; label: string; count: number }[] = [
    { id: "records", label: "All records", count: filteredRecords.length },
    { id: "matches", label: "Match review", count: pendingMatches.length },
    { id: "missing", label: "Missing information", count: missingRecords.length },
    { id: "unmatched", label: "Unmatched", count: unmatchedRecords.length }
  ];

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
          <p className="mt-3 text-xs text-slate-400">Files are processed in your browser. Matching runs in a background worker; large datasets may still require substantial browser memory.</p>
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
          <Stat label="Reviewed" value={Object.keys(decisions).length} note="Decisions in this session" />
        </section>

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
          <button className={buttonSecondary} onClick={() => exportRecords(filteredRecords, "filtered-records.xlsx")}>Export Excel</button>
          <button className={buttonSecondary} onClick={() => exportRows(filteredRecords.map((r) => { const members = comparisonGroups.membersByRepresentative.get(r.id) ?? [r]; return { "Source File": r.source, "Representative Source Row": r.rowNumber, "Repeated Submissions Grouped": members.length > 1 ? "Yes" : "No", "Submission Count": members.length, "Source Rows In Group": members.map((member) => `${member.source}: ${member.rowNumber}`).join("; "), ...r.data }; }), "filtered-records.csv")}>Export CSV</button>
          <button className={buttonSecondary} onClick={exportMatches}>Export matches</button>
          <button className={buttonSecondary} onClick={() => exportRecords(unmatchedRecords, "unmatched-records.xlsx")}>Export unmatched</button>
          <button className={buttonPrimary} onClick={exportUnion}>Export Union Master</button>
        </section>

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
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div><h2 className="font-bold">Match sensitivity</h2><p className="mt-1 text-sm text-slate-500">Raise the threshold to show stronger candidates only.</p></div>
              <span className="rounded-lg bg-indigo-50 px-3 py-1.5 text-sm font-bold text-indigo-700">{threshold}%</span>
            </div>
            <input type="range" min={30} max={95} step={5} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))}
              className="mt-4 w-full accent-indigo-600" />
            <div className="mt-1 flex justify-between text-xs text-slate-400"><span>More candidates</span><span>Stricter matches</span></div>
          </section>
        )}

        {tab === "records" && <RecordTable records={filteredRecords} columns={columns} pageSize={100} />}
        {tab === "matches" && (
          <section className="space-y-3">
            {matching && <div className="rounded-xl border border-indigo-100 bg-indigo-50 p-4 text-sm text-indigo-800">Checking {comparisonGroups.records.length.toLocaleString()} comparison records for within-file duplicates and cross-file matches in the background. You can keep using the interface; results appear when the pass completes.</div>}
            {!matching && pendingMatches.length > 200 && <div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">Showing the first 200 of {pendingMatches.length.toLocaleString()} pending candidates to keep the page responsive. Export matches to review the full candidate report, or narrow the list with search and filters.</div>}
            {pendingMatches.length === 0 ? <Empty title="No pending candidates" detail={sources.length < 1 ? "Import an Excel or CSV file to find possible duplicates." : "No candidates found. Check the column mapping or lower the threshold. Matching checks both duplicates within a file and common records across files."} /> :
              pendingMatches.map((m) => <article key={m.id} className="rounded-2xl border bg-white p-5 shadow-sm">
                <div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-start">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="rounded-full bg-indigo-50 px-3 py-1 text-sm font-bold text-indigo-700">{m.score}% score</span>
                      <span className={`rounded-full px-3 py-1 text-xs font-semibold ${m.status === "Likely match" ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>{m.status}</span>
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
                </div>
              </article>)}
          </section>
        )}
        {tab === "missing" && <RecordTable records={missingRecords} columns={columns} pageSize={100} />}
        {tab === "unmatched" && <RecordTable records={unmatchedRecords} columns={columns} pageSize={100} />}

        {activeMatch && <CompareModal match={activeMatch} onClose={() => setSelectedMatch(null)} />}
        <footer className="mt-8 flex flex-col gap-1 border-t border-slate-200 pt-5 text-xs leading-5 text-slate-400 sm:flex-row sm:justify-between">
          <span>Excel Data Matcher · Standalone local-first tool</span>
          <span>Similarity scores are heuristics, not identity verification.</span>
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

function CompareModal({ match, onClose }: { match: ReturnType<typeof compareRecords>[number]; onClose: () => void }) {
  const fields = Array.from(new Set([...Object.keys(match.left.data), ...Object.keys(match.right.data)]));
  return <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/50 p-3 sm:p-6" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
    <section role="dialog" aria-modal="true" aria-labelledby="compare-title" className="max-h-[90vh] w-full max-w-6xl overflow-auto rounded-2xl bg-white p-5 shadow-2xl sm:p-7">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div><h2 id="compare-title" className="text-xl font-bold">Compare records</h2><p className="mt-1 text-sm text-slate-500">{match.score}% similarity · {match.reasons.join(" · ")}</p></div>
        <button onClick={onClose} className={buttonSecondary}>Close</button>
      </div>
      <div className="mt-5 grid gap-4 md:grid-cols-2">
        {[match.left, match.right].map((record, index) => <div key={record.id} className="rounded-xl border border-slate-200 p-4">
          <h3 className="font-bold">{index === 0 ? "Record A" : "Record B"}</h3><p className="mt-1 break-all text-xs text-slate-500">{record.source} · row {record.rowNumber}</p>
          {fields.map((field) => {
            const a = match.left.data[field] || "";
            const b = match.right.data[field] || "";
            const value = record.data[field] || "";
            const different = a !== b;
            return <div key={field} className={`mt-3 rounded-lg p-2 ${different ? "bg-amber-50" : "bg-slate-50"}`}>
              <p className="text-xs font-semibold text-slate-500">{field}{different ? " · differs" : ""}</p>
              <p className="mt-1 break-words text-sm">{value || <span className="text-amber-700">Missing</span>}</p>
            </div>;
          })}
        </div>)}
      </div>
      <p className="mt-4 text-xs leading-5 text-slate-400">Review the original source and identifiers before confirming a match. A similarity score alone does not establish identity.</p>
    </section>
  </div>;
}

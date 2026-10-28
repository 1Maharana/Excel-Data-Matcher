
export type StudentRecord = {
  id: string;
  source: string;
  rowNumber: number;
  data: Record<string, string>;
};

export type MatchCandidate = {
  id: string;
  left: StudentRecord;
  right: StudentRecord;
  score: number;
  reasons: string[];
  status: "Likely match" | "Review";
};

const aliases: Record<string, string[]> = {
  Name: [
    "name",
    "full name",
    "student name",
    "candidate name",
    "student full name",
    "alumni name"
  ],
  Email: [
    "email",
    "email id",
    "email address",
    "mail",
    "e-mail"
  ],
  Mobile: [
    "mobile",
    "phone",
    "phone number",
    "mobile number",
    "contact number",
    "telephone"
  ],
  "Roll No": [
    "roll no",
    "roll number",
    "usn",
    "student id",
    "registration number",
    "register number"
  ],
  Department: [
    "department",
    "dept",
    "branch",
    "course department",
    "program"
  ],
  Gender: ["gender", "sex"],
  "Graduation Year": [
    "graduation year",
    "passout year",
    "passing year",
    "year of graduation",
    "year of graduation (1)"
  ],
  City: ["city", "town"],
  State: ["state", "province", "region"],
  Country: ["country"],
  Degree: ["degree", "qualification", "degree (1)"]
};

export const standardFields = Object.keys(aliases);

export function normalizeHeader(header: string): string {
  return header
    .toLowerCase()
    .trim()
    .replace(/\s*\(\d+\)\s*$/, "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

export function canonicalHeader(header: string): string {
  const normalized = normalizeHeader(header);

  for (const [canonical, names] of Object.entries(aliases)) {
    if (
      names.some((name) => normalizeHeader(name) === normalized)
    ) {
      return canonical;
    }
  }

  return header.trim() || "Unnamed column";
}

export function normalizeValue(value: unknown): string {
  if (value === null || value === undefined) return "";

  return String(value)
    .trim()
    .replace(/\s+/g, " ");
}

export function normalizeRecords(
  records: StudentRecord[],
  mapping: Record<string, string>
): StudentRecord[] {
  return records.map((record) => {
    const data: Record<string, string> = {};

    for (const [original, value] of Object.entries(record.data)) {
      const mapped = mapping[original] || canonicalHeader(original);
      const key = mapped.trim() || original;

      // Preserve a nonempty value if multiple columns map to one field.
      if (!(key in data) || (!data[key] && value)) {
        data[key] = value;
      }
    }

    return { ...record, data };
  });
}

function valueFor(
  record: StudentRecord,
  field: string
): string {
  // Prefer the canonical key if normalization has already occurred.
  if (record.data[field] !== undefined) {
    return normalizeValue(record.data[field]);
  }

  const entry = Object.entries(record.data).find(
    ([key]) => canonicalHeader(key) === field
  );

  return normalizeValue(entry?.[1]);
}

function compact(value: string): string {
  return value
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

export function similarity(a: string, b: string): number {
  const x = compact(a);
  const y = compact(b);

  if (!x || !y) return 0;
  if (x === y) return 1;

  // Save memory by keeping only two rows of the edit-distance matrix.
  let previous = Array.from(
    { length: y.length + 1 },
    (_, i) => i
  );

  for (let i = 1; i <= x.length; i++) {
    const current = new Array<number>(y.length + 1);
    current[0] = i;

    for (let j = 1; j <= y.length; j++) {
      const substitutionCost =
        x[i - 1] === y[j - 1] ? 0 : 1;

      current[j] = Math.min(
        previous[j] + 1,
        current[j - 1] + 1,
        previous[j - 1] + substitutionCost
      );
    }

    previous = current;
  }

  return (
    1 - previous[y.length] / Math.max(x.length, y.length)
  );
}

const MATCH_FIELDS = [
  { name: "Roll No", weight: 40, identifier: true },
  { name: "Mobile", weight: 35, identifier: true },
  { name: "Email", weight: 35, identifier: true },
  { name: "Name", weight: 30, identifier: false },
  { name: "Department", weight: 10, identifier: false },
  { name: "Degree", weight: 10, identifier: false },
  { name: "Graduation Year", weight: 10, identifier: false },
  { name: "Gender", weight: 5, identifier: false }
] as const;

/** Normalize common value variants for comparison only; exports retain original values. */
function comparisonValue(field: string, value: string): string {
  const trimmed = normalizeValue(value).toLowerCase();
  if (!trimmed) return "";

  if (field === "Degree") {
    const token = compact(trimmed);
    const degreeAliases: Record<string, string> = {
      mscengineering: "mscengineering",
      mscengg: "mscengineering",
      msceng: "mscengineering",
      masterofscienceengineering: "mscengineering",
      be: "be",
      btech: "btech",
      me: "me",
      mengg: "me",
      mtech: "mtech",
      phd: "phd",
      doctorofphilosophy: "phd"
    };
    return degreeAliases[token] ?? token;
  }

  if (field === "Department") {
    return compact(trimmed
      .replace(/&/g, " and ")
      .replace(/\bdept\.?\b/g, "department")
      .replace(/\bengg\.?\b/g, "engineering")
      .replace(/\beng\.?\b/g, "engineering"));
  }

  if (field === "Graduation Year") {
    const year = trimmed.match(/\b(19|20)\d{2}\b/);
    return year?.[0] ?? compact(trimmed);
  }

  if (field === "Gender") {
    if (["m", "male"].includes(compact(trimmed))) return "male";
    if (["f", "female"].includes(compact(trimmed))) return "female";
  }

  return compact(trimmed);
}

function addToIndex(index: Map<string, number[]>, key: string, recordIndex: number): void {
  const bucket = index.get(key);
  if (bucket) bucket.push(recordIndex);
  else index.set(key, [recordIndex]);
}

export function compareRecords(records: StudentRecord[], threshold = 60): MatchCandidate[] {
  if (records.length < 2) return [];

  // Cache normalized comparison values once; the source data itself is never changed.
  const normalized = records.map((record) => {
    const values: Record<string, string> = {};
    for (const field of MATCH_FIELDS) {
      values[field.name] = comparisonValue(field.name, valueFor(record, field.name));
    }
    return values;
  });

  const index = new Map<string, number[]>();
  records.forEach((_record, i) => {
    const values = normalized[i];
    for (const field of MATCH_FIELDS) {
      const token = values[field.name];
      if (!token) continue;
      if (field.identifier) addToIndex(index, `${field.name}:${token}`, i);

      if (field.name === "Name") {
        addToIndex(index, `name-exact:${token}`, i);
        // Prefix + word indexes catch common spelling/formatting variations.
        if (token.length >= 4) addToIndex(index, `name-prefix:${token.slice(0, 4)}`, i);
        const originalName = valueFor(records[i], "Name").toLowerCase();
        for (const word of originalName.split(/\s+/)) {
          const wordToken = compact(word);
          if (wordToken.length >= 3) addToIndex(index, `name-word:${wordToken}`, i);
        }
      }
    }
  });

  // Generate both cross-file matches and within-file duplicate candidates.
  // Buckets with too many records are skipped to avoid quadratic explosions.
  // Identifier buckets are inserted first, so the strongest evidence is prioritized.
  const pairs = new Set<string>();
  const MAX_PAIRS_PER_BUCKET = 20_000;
  const MAX_TOTAL_CANDIDATE_PAIRS = 250_000;

  for (const bucket of index.values()) {
    if (bucket.length < 2) continue;
    const estimatedPairs = (bucket.length * (bucket.length - 1)) / 2;
    if (estimatedPairs > MAX_PAIRS_PER_BUCKET) continue;

    for (let a = 0; a < bucket.length; a++) {
      for (let b = a + 1; b < bucket.length; b++) {
        if (pairs.size >= MAX_TOTAL_CANDIDATE_PAIRS) break;
        const i = Math.min(bucket[a], bucket[b]);
        const j = Math.max(bucket[a], bucket[b]);
        if (i !== j) pairs.add(`${i}:${j}`);
      }
      if (pairs.size >= MAX_TOTAL_CANDIDATE_PAIRS) break;
    }
    if (pairs.size >= MAX_TOTAL_CANDIDATE_PAIRS) break;
  }

  const results: MatchCandidate[] = [];
  for (const pair of pairs) {
    const separator = pair.indexOf(":");
    const i = Number(pair.slice(0, separator));
    const j = Number(pair.slice(separator + 1));
    const left = records[i];
    const right = records[j];
    let points = 0;
    let availableWeight = 0;
    let conflictingStrongIds = 0;
    let exactStrongId = false;
    const reasons: string[] = [];

    for (const field of MATCH_FIELDS) {
      const a = normalized[i][field.name];
      const b = normalized[j][field.name];
      if (!a || !b) continue;
      availableWeight += field.weight;
      const sim = similarity(a, b);
      if (field.identifier && sim === 1) exactStrongId = true;
      if (field.identifier && sim < 0.85) conflictingStrongIds++;

      const accepted = sim === 1 || (field.name === "Name" && sim >= 0.7) || (!field.identifier && sim >= 0.8);
      if (accepted) {
        points += field.weight * (field.identifier ? 1 : sim);
        reasons.push(sim === 1 ? `Same ${field.name}` : `Similar ${field.name}`);
      } else if (!field.identifier && field.name !== "Gender" && sim < 0.5) {
        // Keep major conflicts visible in the candidate review details.
        reasons.push(`Different ${field.name}`);
      }
    }

    if (availableWeight === 0 || !reasons.some((r) => r.startsWith("Same ") || r.startsWith("Similar "))) continue;
    let score = Math.round((points / availableWeight) * 100);
    if (conflictingStrongIds > 0 && !exactStrongId) score = Math.min(score, 59);
    if (score < threshold) continue;

    const sameSource = left.source === right.source;
    results.push({
      id: [left.id, right.id].sort().join("::"),
      left,
      right,
      score,
      reasons: [
        ...(sameSource ? ["Within-file duplicate candidate"] : ["Cross-file match candidate"]),
        ...reasons,
        ...(conflictingStrongIds > 0 ? ["Conflicting identifier(s)—review carefully"] : [])
      ],
      status: score >= 80 && (exactStrongId || normalized[i].Name === normalized[j].Name)
        ? "Likely match"
        : "Review"
    });
  }

  return results.sort((a, b) => b.score - a.score);
}

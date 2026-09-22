"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  ArrowUpRight,
  Check,
  CircleStop,
  Database,
  Download,
  FileCheck2,
  FileSpreadsheet,
  LoaderCircle,
  MapPin,
  Pencil,
  Plus,
  RotateCcw,
  Search,
  SearchCheck,
  ShieldCheck,
  Sparkles,
  Stethoscope,
  UserCheck,
  UploadCloud,
  X,
} from "lucide-react";
import * as XLSX from "xlsx";

import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import {
  Pagination,
  PaginationContent,
  PaginationItem,
  PaginationLink,
  PaginationNext,
  PaginationPrevious,
} from "@/components/ui/pagination";
import { Progress } from "@/components/ui/progress";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";

type Cell = string | number | boolean | Date | null | undefined;
type ContactRow = Record<string, Cell>;

type NpiTaxonomy = {
  code?: string;
  desc?: string;
  primary?: boolean;
  state?: string;
  license?: string;
};

type NpiResult = {
  number?: string | number;
  enumeration_type?: string;
  basic?: {
    first_name?: string;
    last_name?: string;
    middle_name?: string;
    credential?: string;
    status?: string;
    organization_name?: string;
  };
  addresses?: Array<{
    state?: string;
    city?: string;
    country_code?: string;
    country_name?: string;
    address_purpose?: string;
  }>;
  taxonomies?: NpiTaxonomy[];
};

type CountryDecision = "confirmed-us" | "likely-us" | "confirmed-non-us" | "ambiguous" | "unresolved";
type DoctorDecision = "confirmed" | "likely" | "not-physician" | "ambiguous" | "unresolved";
type SpecialtyDecision = "confirmed" | "conflict" | "insufficient" | "unreviewed";
type AiConfidence = "high" | "medium" | "low" | "insufficient";

type AiVerificationResponse = {
  identityMatch: boolean | null;
  practicesInUs: boolean | null;
  usPracticeReason: string;
  isPhysician: boolean | null;
  usDoctorConfidence: number;
  verifiedSpecialty: string;
  specialtyMatchScore: number | null;
  specialtyMatchReason: string;
  confidence: AiConfidence;
  reason: string;
  evidenceUrls: string[];
};

type AiNpiDiscoveryResponse = {
  candidateNpi: string;
  registryFirstName: string;
  registryLastName: string;
  alternateName: string;
  confidence: AiConfidence;
  reason: string;
  evidenceUrls: string[];
};

type EnrichedRow = ContactRow & {
  __index: number;
  __eligible: boolean;
  __displayName: string;
  __state: string;
  __status: "queued" | "running" | "verified" | "review" | "not-found" | "not-physician" | "error" | "skipped";
  __countryDecision: CountryDecision;
  __countryReason: string;
  __doctorDecision: DoctorDecision;
  __doctorReason: string;
  __specialtyDecision: SpecialtyDecision;
  __manualCountry?: boolean;
  __manualDoctor?: boolean;
  __manualSpecialty?: boolean;
  __aiStatus?: "not-run" | "checking" | "verified" | "uncertain" | "error";
  __aiConfidence?: AiConfidence;
  __aiReason?: string;
  __aiEvidenceUrls?: string;
  __aiUsDoctor?: boolean;
  __npiEvidenceUrls?: string;
  __excluded?: boolean;
  __manualAdded?: boolean;
};

type NewContact = {
  firstName: string;
  lastName: string;
  email: string;
  company: string;
  country: string;
  state: string;
  license: string;
  specialty: string;
  linkedIn: string;
  npi: string;
};

type FieldMap = {
  email?: string;
  firstName?: string;
  lastName?: string;
  fullName?: string;
  license?: string;
  country?: string;
  state?: string;
  company?: string;
  claimedSpecialty?: string;
  categoryDetail?: string;
  linkedIn?: string;
};

declare global {
  interface Document {
    modelContext?: {
      registerTool: (
        tool: {
          name: string;
          title?: string;
          description: string;
          inputSchema: object;
          annotations?: { readOnlyHint?: boolean; untrustedContentHint?: boolean };
          execute: (input: unknown) => unknown | Promise<unknown>;
        },
        options?: { signal?: AbortSignal },
      ) => void | Promise<void>;
    };
  }
}

const FIELD_ALIASES: Record<keyof FieldMap, string[]> = {
  email: ["email address", "email", "work email"],
  firstName: ["first name", "firstname", "first_name", "given name"],
  lastName: ["last name", "lastname", "last_name", "surname", "family name"],
  fullName: ["full name", "name", "contact name"],
  license: ["license key", "licensekey", "license", "licence key"],
  country: ["country", "country code", "nation"],
  state: ["us state", "state", "state code", "region"],
  company: ["company", "organization", "organisation", "affiliation", "practice"],
  claimedSpecialty: ["cancer specialty", "specialty", "speciality", "cancer speciality"],
  categoryDetail: ["category detail", "category details", "hcp type"],
  linkedIn: ["linkedin or website", "person linkedin url", "linkedin url", "linkedin", "website"],
};

const DEFAULT_FIELDS: Required<FieldMap> = {
  email: "Email Address",
  firstName: "First name",
  lastName: "Last name",
  fullName: "Full name",
  license: "License key",
  country: "Country",
  state: "US State",
  company: "Company",
  claimedSpecialty: "Cancer specialty",
  categoryDetail: "Category Detail",
  linkedIn: "Linkedin or Website",
};

const EMPTY_CONTACT: NewContact = {
  firstName: "",
  lastName: "",
  email: "",
  company: "",
  country: "",
  state: "",
  license: "",
  specialty: "",
  linkedIn: "",
  npi: "",
};

type ReviewDraft = {
  country: CountryDecision;
  doctor: DoctorDecision;
  specialty: SpecialtyDecision;
  evidenceUrl: string;
  notes: string;
};

const EMPTY_REVIEW: ReviewDraft = {
  country: "unresolved",
  doctor: "unresolved",
  specialty: "unreviewed",
  evidenceUrl: "",
  notes: "",
};

const normalize = (value: Cell) =>
  String(value ?? "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

const textValue = (row: ContactRow, key?: string) =>
  key ? String(row[key] ?? "").trim() : "";

function findHeader(headers: string[], aliases: string[]) {
  const normalized = headers.map((header) => normalize(header));
  for (const alias of aliases) {
    const exact = normalized.indexOf(alias);
    if (exact >= 0) return headers[exact];
  }
  return undefined;
}

function mapFields(headers: string[]): FieldMap {
  return Object.fromEntries(
    Object.entries(FIELD_ALIASES).map(([field, aliases]) => [field, findHeader(headers, aliases)]),
  ) as FieldMap;
}

function columnLetter(index: number) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function isValidNpi(value: string) {
  if (!/^\d{10}$/.test(value)) return false;
  const digits = `80840${value}`.split("").map(Number).reverse();
  const sum = digits.reduce((total, digit, index) => {
    if (index % 2 === 0) return total + digit;
    const doubled = digit * 2;
    return total + (doubled > 9 ? doubled - 9 : doubled);
  }, 0);
  return sum % 10 === 0;
}

function isUS(value: string) {
  return ["usa", "us", "united states", "united states of america"].includes(normalize(value));
}

function activeNpi(result: NpiResult) {
  return result.basic?.status !== "D";
}

function primaryTaxonomy(result: NpiResult) {
  return result.taxonomies?.find((taxonomy) => taxonomy.primary) ?? result.taxonomies?.[0];
}

function registryName(result: NpiResult) {
  return [result.basic?.first_name, result.basic?.middle_name, result.basic?.last_name]
    .filter(Boolean)
    .join(" ")
    .trim();
}

function organizationDomain(row: ContactRow, fields: FieldMap) {
  const email = textValue(row, fields.email);
  const domain = email.split("@")[1]?.trim().toLowerCase() ?? "";
  return /^[a-z0-9.-]+\.[a-z]{2,}$/i.test(domain) ? domain : "";
}

function hasPhysicianCredential(result: NpiResult) {
  const credential = normalize(result.basic?.credential);
  return /(^| )(md|do|mbbs|mb chb|mbchb)( |$)/.test(credential);
}

function hasPhysicianTaxonomy(result: NpiResult) {
  return (result.taxonomies ?? []).some((taxonomy) =>
    /^20/.test(String(taxonomy.code ?? "")),
  );
}

function practiceAddress(result: NpiResult) {
  return result.addresses?.find((address) => address.address_purpose === "LOCATION") ?? result.addresses?.[0];
}

function addressIsUS(result: NpiResult) {
  const address = practiceAddress(result);
  const country = normalize(address?.country_code || address?.country_name);
  return country === "us" || country === "usa" || country === "united states" || (!country && /^[A-Za-z]{2}$/.test(address?.state ?? ""));
}

function doctorAssessment(result: NpiResult, clearMatch: boolean): { decision: DoctorDecision; reason: string } {
  if (!clearMatch) return { decision: "ambiguous", reason: "The NPI candidate exists, but the identity match is not strong enough." };
  if (result.enumeration_type !== "NPI-1") return { decision: "not-physician", reason: "This NPI belongs to an organization, not an individual." };
  if (!activeNpi(result)) return { decision: "unresolved", reason: "The matched NPI is deactivated." };
  if (hasPhysicianTaxonomy(result)) return { decision: "confirmed", reason: "Active individual NPI with an Allopathic & Osteopathic Physicians taxonomy (code family 20)." };
  if (hasPhysicianCredential(result)) return { decision: "likely", reason: "Physician credential found, but the registered taxonomy is not a physician taxonomy." };
  return { decision: "not-physician", reason: "The active individual NPI does not have a physician taxonomy." };
}

function nameFor(row: ContactRow, fields: FieldMap) {
  const combined = `${textValue(row, fields.firstName)} ${textValue(row, fields.lastName)}`.trim();
  return combined || textValue(row, fields.fullName) || "Unnamed contact";
}

function scoreCandidate(result: NpiResult, row: ContactRow, fields: FieldMap) {
  const first = normalize(textValue(row, fields.firstName));
  const last = normalize(textValue(row, fields.lastName));
  const state = normalize(textValue(row, fields.state));
  const company = normalize(textValue(row, fields.company));
  let score = 0;

  if (first && normalize(result.basic?.first_name) === first) score += 30;
  if (last && normalize(result.basic?.last_name) === last) score += 40;
  if (state && result.addresses?.some((address) => normalize(address.state) === state)) score += 20;
  if (company && normalize(result.basic?.organization_name).includes(company)) score += 5;
  if (activeNpi(result)) score += 5;
  if (hasPhysicianTaxonomy(result)) score += 5;
  return Math.min(score, 100);
}

function initialCountryDecision(country: string): { decision: CountryDecision; reason: string } {
  if (isUS(country)) return { decision: "likely-us", reason: "Source lists United States." };
  if (!country.trim()) return { decision: "unresolved", reason: "Country is blank; NPPES practice-address research is required." };
  return { decision: "confirmed-non-us", reason: `The source file lists ${country}.` };
}

function countryAssessment(sourceCountry: string, result: NpiResult | undefined, clearMatch: boolean): { decision: CountryDecision; reason: string } {
  if (isUS(sourceCountry)) {
    if (result && clearMatch && addressIsUS(result)) {
      const address = practiceAddress(result);
      const location = [address?.city, address?.state].filter(Boolean).join(", ");
      return { decision: "likely-us", reason: `Source and NPPES support U.S. practice${location ? ` in ${location}` : ""}.` };
    }
    if (result && clearMatch && !addressIsUS(result)) {
      return { decision: "ambiguous", reason: "Source lists United States, but the matched NPPES address does not; AI review is required." };
    }
    return { decision: "likely-us", reason: "Source lists United States." };
  }
  if (sourceCountry.trim()) {
    if (result && clearMatch && addressIsUS(result)) {
      const address = practiceAddress(result);
      const location = [address?.city, address?.state, address?.country_name || address?.country_code].filter(Boolean).join(", ");
      return { decision: "likely-us", reason: `The source lists ${sourceCountry}, but the verified NPI has a current U.S. practice address${location ? ` (${location})` : ""}.` };
    }
    return { decision: "confirmed-non-us", reason: `The source file lists ${sourceCountry}.` };
  }
  if (!result) return { decision: "unresolved", reason: "Country is blank and no NPPES candidate was found." };
  const address = practiceAddress(result);
  const location = [address?.city, address?.state, address?.country_name || address?.country_code].filter(Boolean).join(", ");
  if (!clearMatch) return { decision: "ambiguous", reason: `A possible NPPES location was found${location ? ` (${location})` : ""}, but identity needs review.` };
  if (addressIsUS(result)) return { decision: "likely-us", reason: `A strong NPI match has a U.S. practice address${location ? ` (${location})` : ""}.` };
  return { decision: "confirmed-non-us", reason: `The matched NPPES practice address is outside the U.S.${location ? ` (${location})` : ""}.` };
}

function specialtyAssessment(claimed: string, result: NpiResult) {
  const taxonomy = primaryTaxonomy(result);
  const claim = normalize(claimed);
  const npiSpecialty = normalize(taxonomy?.desc);
  if (!claim) return "No claimed specialty to compare";
  if (!npiSpecialty) return "NPI specialty unavailable";

  const oncologyClaim = /(oncol|hematol|cancer|tumor|tumour|myeloma|leukemia|lymphoma)/.test(claim);
  const oncologyNpi = /(oncol|hematol)/.test(npiSpecialty);
  const diseaseFocus = /(breast|lung|thoracic|prostate|genitourinary|gastro|myeloma|leukemia|lymphoma|neuro)/.test(claim);
  const claimTokens = new Set(claim.split(" ").filter((token) => token.length > 3));
  const overlap = npiSpecialty.split(" ").some((token) => claimTokens.has(token));

  if (overlap || (oncologyClaim && oncologyNpi)) {
    return diseaseFocus
      ? "NPI aligns with oncology; disease focus needs web or LinkedIn review"
      : "NPI specialty aligns";
  }
  return "Specialty conflict or manual review required";
}

function linkedinUrl(row: ContactRow, fields: FieldMap) {
  const existing = textValue(row, fields.linkedIn);
  if (/linkedin\.com/i.test(existing)) return existing;
  const query = encodeURIComponent(`${nameFor(row, fields)} ${textValue(row, fields.company)}`.trim());
  return `https://www.linkedin.com/search/results/people/?keywords=${query}`;
}

function webUrl(row: ContactRow, fields: FieldMap) {
  const existing = textValue(row, fields.linkedIn);
  if (/^https?:\/\//i.test(existing) && !/linkedin\.com/i.test(existing)) return existing;
  const query = encodeURIComponent(
    `${nameFor(row, fields)} ${textValue(row, fields.company)} physician specialty`,
  );
  return `https://www.google.com/search?q=${query}`;
}

function locationUrl(row: ContactRow, fields: FieldMap) {
  const query = encodeURIComponent(
    `${nameFor(row, fields)} ${textValue(row, fields.company)} physician location United States`,
  );
  return `https://www.google.com/search?q=${query}`;
}

async function requestNpiResults(params: URLSearchParams, signal?: AbortSignal) {
  let response: Response | undefined;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    response = await fetch(`/api/npi?${params.toString()}`, { cache: "no-store", signal });
    if (response.status !== 429 && response.status < 500) break;
    await new Promise((resolve) => setTimeout(resolve, 900 * 2 ** attempt));
  }
  if (!response?.ok) throw new Error("NPI lookup failed");
  const payload = (await response.json()) as { results?: NpiResult[] };
  return payload.results ?? [];
}

async function responseErrorMessage(response: Response, fallback: string) {
  try {
    const payload = await response.json() as { error?: string; detail?: string };
    return payload.detail || payload.error || fallback;
  } catch {
    return fallback;
  }
}

function retryDelay(milliseconds: number, signal?: AbortSignal) {
  return new Promise<void>((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException("Request cancelled", "AbortError"));
      return;
    }
    const timeout = window.setTimeout(resolve, milliseconds);
    signal?.addEventListener("abort", () => {
      window.clearTimeout(timeout);
      reject(new DOMException("Request cancelled", "AbortError"));
    }, { once: true });
  });
}

async function fetchWithRetry(input: RequestInfo | URL, init: RequestInit, attempts = 3) {
  let response: Response | undefined;
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    response = await fetch(input, init);
    const retryable = response.status === 429 || response.status === 503 || response.status === 504;
    if (!retryable) return response;
    if (attempt < attempts - 1) {
      await response.body?.cancel();
      await retryDelay(700 * 2 ** attempt, init.signal ?? undefined);
    }
  }
  return response as Response;
}

function cleanExportRow(row: EnrichedRow): ContactRow {
  const hiddenExportFields = new Set([
    "NPI Entity Type",
    "NPI Taxonomy Code",
    "NPI Status",
    "Doctor Confirmation Reason",
    "Suggested NPI",
    "Country Resolution",
    "Country Research URL",
    "Specialty Comparison",
    "NPI Candidate Count",
    "Specialty Decision",
    "Enriched At (UTC)",
    "Specialty Evidence",
    "Country Evidence",
    "Doctor Confirmation",
    "Registered NPI Specialty",
    "Identity Match",
    "Identity Match Reason",
    "AI Specialty Match Reason",
  ]);
  return Object.fromEntries(
    Object.entries(row).filter(([key]) => !key.startsWith("__") && !hiddenExportFields.has(key)),
  );
}

function yesNoMaybe(value: boolean | null | undefined) {
  return value === true ? "Yes" : value === false ? "No" : "Maybe";
}

function confidenceLabel(value: boolean | null | undefined, score: number) {
  const bounded = Math.max(0, Math.min(100, Math.round(score)));
  return `${yesNoMaybe(value)} · ${bounded}%`;
}

function specialtyMatchLabel(score: number | null | undefined) {
  if (typeof score !== "number") return "";
  const bounded = Math.max(0, Math.min(100, Math.round(score)));
  const label = bounded >= 85 ? "Yes" : bounded >= 60 ? "Mostly" : "No";
  return `${label} · ${bounded}%`;
}

function shortQueueReason(value: unknown) {
  const text = String(value ?? "").replace(/^ChatGPT web verification:\s*/i, "").replace(/\s+/g, " ").trim();
  if (text.length <= 180) return text;
  const shortened = text.slice(0, 180);
  const wordBoundary = shortened.lastIndexOf(" ");
  return `${shortened.slice(0, wordBoundary > 120 ? wordBoundary : 179).trimEnd()}…`;
}

function downloadExcelWorkbook(allUsRows: EnrichedRow[], doctorRows: EnrichedRow[]) {
  if (!allUsRows.length) return;
  const allContacts = allUsRows.map(cleanExportRow);
  const doctors = doctorRows.map(cleanExportRow);
  const allHeaders = Array.from(new Set(allContacts.flatMap((row) => Object.keys(row))));
  const decisionHeaders = [
    "Claimed Specialty",
    "NPI Primary Specialty",
    "Verified Specialty",
    "AI Specialty Match",
    "US Doctor",
    "US Doctor Reason",
  ];
  const headers = [
    ...allHeaders.filter((header) => !decisionHeaders.includes(header)),
    ...decisionHeaders.filter((header) => allHeaders.includes(header)),
  ];
  const workbook = XLSX.utils.book_new();
  const allSheet = XLSX.utils.json_to_sheet(allContacts, { header: headers });
  const doctorSheet = XLSX.utils.json_to_sheet(doctors, { header: headers });
  const columnWidths = headers.map((header) => ({
    wch: Math.min(55, Math.max(14, header.length + 2, ...allContacts.slice(0, 250).map((row) => String(row[header] ?? "").length))),
  }));
  allSheet["!cols"] = columnWidths;
  doctorSheet["!cols"] = columnWidths;
  if (allSheet["!ref"]) allSheet["!autofilter"] = { ref: allSheet["!ref"] };
  if (doctorSheet["!ref"]) doctorSheet["!autofilter"] = { ref: doctorSheet["!ref"] };
  XLSX.utils.book_append_sheet(workbook, allSheet, "All US Contacts");
  XLSX.utils.book_append_sheet(workbook, doctorSheet, "US Doctors");
  XLSX.writeFile(workbook, "us-contact-enrichment.xlsx", { compression: true });
}

function StatusPill({ status }: { status: EnrichedRow["__status"] }) {
  const label: Record<EnrichedRow["__status"], string> = {
    queued: "Queued",
    running: "Checking",
    verified: "Verified physician",
    review: "Review match",
    "not-found": "No exact match",
    "not-physician": "Not confirmed as physician",
    error: "API error",
    skipped: "Not in scope",
  };
  return <span className={`status-pill status-${status}`}>{label[status]}</span>;
}

function AnswerPill({ value }: { value: unknown }) {
  const display = String(value || "Not checked");
  const tone = display.startsWith("Yes") ? "yes" : display.startsWith("No") ? "no" : display.startsWith("Mostly") || display.startsWith("Maybe") ? "maybe" : "unchecked";
  return <span className={`answer-pill answer-${tone}`}>{display}</span>;
}

const COUNTRY_LABELS: Record<CountryDecision, string> = {
  "confirmed-us": "Confirmed US",
  "likely-us": "US",
  "confirmed-non-us": "Confirmed non-US",
  ambiguous: "Country ambiguous",
  unresolved: "Country unresolved",
};

const DOCTOR_LABELS: Record<DoctorDecision, string> = {
  confirmed: "Confirmed physician",
  likely: "Likely physician",
  "not-physician": "Not a physician",
  ambiguous: "Identity ambiguous",
  unresolved: "Doctor check unresolved",
};

const SPECIALTY_LABELS: Record<SpecialtyDecision, string> = {
  confirmed: "Confirmed",
  conflict: "Conflict",
  insufficient: "Not enough evidence",
  unreviewed: "Not reviewed",
};

function DecisionPill({ kind, value }: { kind: "country" | "doctor" | "specialty"; value: CountryDecision | DoctorDecision | SpecialtyDecision }) {
  const label = kind === "country"
    ? COUNTRY_LABELS[value as CountryDecision]
    : kind === "doctor"
      ? DOCTOR_LABELS[value as DoctorDecision]
      : SPECIALTY_LABELS[value as SpecialtyDecision];
  return <span className={`decision-pill decision-${value}`}>{label}</span>;
}

export default function Home() {
  const [fileName, setFileName] = useState("");
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState("");
  const workbookRef = useRef<XLSX.WorkBook | null>(null);
  const [rows, setRows] = useState<EnrichedRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fields, setFields] = useState<FieldMap>({});
  const [message, setMessage] = useState("Upload the current contact file to begin.");
  const [isRunning, setIsRunning] = useState(false);
  const [isAiRunning, setIsAiRunning] = useState(false);
  const [aiConfigured, setAiConfigured] = useState(false);
  const [aiConfigurationChecked, setAiConfigurationChecked] = useState(false);
  const [processed, setProcessed] = useState(0);
  const [runTotal, setRunTotal] = useState(0);
  const [searchQuery, setSearchQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [pageSize, setPageSize] = useState("250");
  const [page, setPage] = useState(1);
  const [lastExcluded, setLastExcluded] = useState<number | null>(null);
  const [npiDialogRow, setNpiDialogRow] = useState<EnrichedRow | null>(null);
  const [manualNpi, setManualNpi] = useState("");
  const [manualNpiError, setManualNpiError] = useState("");
  const [addDialogOpen, setAddDialogOpen] = useState(false);
  const [newContact, setNewContact] = useState<NewContact>(EMPTY_CONTACT);
  const [newContactError, setNewContactError] = useState("");
  const [reviewDialogRow, setReviewDialogRow] = useState<EnrichedRow | null>(null);
  const [reviewDraft, setReviewDraft] = useState<ReviewDraft>(EMPTY_REVIEW);
  const cancelRef = useRef(false);
  const abortControllerRef = useRef<AbortController | null>(null);
  const aiVerificationCacheRef = useRef(new Map<string, Promise<Partial<EnrichedRow>>>());

  const eligibleRows = useMemo(() => rows.filter((row) => row.__eligible && !row.__excluded), [rows]);
  const excludedCount = rows.filter((row) => row.__excluded).length;
  const manualCount = eligibleRows.filter((row) => row.__manualAdded).length;
  const completeRows = useMemo(
    () => eligibleRows.filter((row) => !["queued", "running"].includes(row.__status)),
    [eligibleRows],
  );
  const verified = eligibleRows.filter((row) => row.__doctorDecision === "confirmed").length;
  const aiVerified = eligibleRows.filter((row) => row.__aiStatus === "verified").length;
  const countryReview = eligibleRows.filter((row) => ["unresolved", "ambiguous"].includes(row.__countryDecision)).length;
  const specialtyReview = eligibleRows.filter((row) => ["unreviewed", "insufficient"].includes(row.__specialtyDecision)).length;
  const needsReview = completeRows.filter((row) =>
    ["review", "not-found", "not-physician", "error"].includes(row.__status),
  ).length;
  const exportRows = useMemo(() => eligibleRows.filter((row) =>
    row.__doctorDecision === "confirmed"
    && ["confirmed-us", "likely-us"].includes(row.__countryDecision)
    && Boolean(row["NPI Number"]),
  ), [eligibleRows]);
  const usContactRows = useMemo(() => eligibleRows.filter((row) =>
    ["confirmed-us", "likely-us"].includes(row.__countryDecision),
  ), [eligibleRows]);
  const aiTargetsAvailable = eligibleRows.some((row) =>
    row.__doctorDecision === "confirmed"
    && Boolean(row["NPI Number"])
    && row.__aiStatus !== "verified"
    && !row.__manualSpecialty,
  );
  const licenseColumn = fields.license ? columnLetter(headers.indexOf(fields.license)) : "—";

  useEffect(() => {
    let active = true;
    void fetch("/api/ai-verify", { cache: "no-store" })
      .then((response) => response.ok ? response.json() : { configured: false })
      .then((result) => {
        if (active) setAiConfigured(Boolean((result as { configured?: boolean }).configured));
      })
      .catch(() => {
        if (active) setAiConfigured(false);
      })
      .finally(() => {
        if (active) setAiConfigurationChecked(true);
      });
    return () => { active = false; };
  }, []);

  const loadSheet = useCallback((workbook: XLSX.WorkBook, selected: string) => {
    aiVerificationCacheRef.current.clear();
    const sheet = workbook.Sheets[selected];
    const parsed = XLSX.utils.sheet_to_json<ContactRow>(sheet, { defval: "", raw: false });
    const parsedHeaders = parsed.length ? Object.keys(parsed[0]) : [];
    const mapped = mapFields(parsedHeaders);
    const decorated = parsed.map((row, index) => {
      const country = textValue(row, mapped.country);
      const countryCheck = initialCountryDecision(country);
      const eligible = Boolean(textValue(row, mapped.license)) && (isUS(country) || !country);
      return {
        ...row,
        __index: index,
        __eligible: eligible,
        __displayName: nameFor(row, mapped),
        __state: textValue(row, mapped.state),
        __status: eligible ? "queued" : "skipped",
        __countryDecision: countryCheck.decision,
        __countryReason: countryCheck.reason,
        __doctorDecision: "unresolved",
        __doctorReason: "NPI lookup has not been run.",
        __specialtyDecision: "unreviewed",
        "Resolved Practice Country": ["confirmed-us", "likely-us"].includes(countryCheck.decision) ? "United States" : countryCheck.decision === "confirmed-non-us" ? country : "",
        "Country Resolution": COUNTRY_LABELS[countryCheck.decision],
        "Country Evidence": countryCheck.reason,
        "Doctor Confirmation": DOCTOR_LABELS.unresolved,
        "Doctor Confirmation Reason": "NPI lookup has not been run.",
        "Verified Specialty": "",
        "Claimed Specialty": textValue(row, mapped.claimedSpecialty) || textValue(row, mapped.categoryDetail),
        "Identity Match": "Maybe",
        "Identity Match Reason": "Identity verification has not been run.",
        "US Doctor": "Maybe · 0%",
        "US Doctor Reason": "NPI and identity verification have not been run.",
        "AI Specialty Match": "",
        "AI Specialty Match Reason": "AI specialty comparison has not been run.",
        "Specialty Decision": SPECIALTY_LABELS.unreviewed,
      } satisfies EnrichedRow;
    });
    setSheetName(selected);
    setRows(decorated);
    setHeaders(parsedHeaders);
    setFields(mapped);
    setProcessed(0);
    setRunTotal(0);
    const eligible = decorated.filter((row) => row.__eligible).length;
    const blankCountry = decorated.filter((row) => row.__eligible && row.__countryDecision === "unresolved").length;
    const missing = ["firstName", "lastName", "license"]
      .filter((key) => !mapped[key as keyof FieldMap])
      .join(", ");
    setMessage(
      missing
        ? `Missing required fields: ${missing}. Rename the headers and upload again.`
        : `${eligible.toLocaleString()} licensed contacts are in scope. ${blankCountry.toLocaleString()} have a blank country and will be researched through NPPES.`,
    );
  }, []);

  async function handleFile(file?: File) {
    if (!file) return;
    if (file.size > 30 * 1024 * 1024) {
      setMessage("This file is over the 30 MB browser-processing limit.");
      return;
    }
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      workbookRef.current = workbook;
      setFileName(file.name);
      setSheetNames(workbook.SheetNames);
      const preferred = workbook.SheetNames.find((name) => normalize(name) === "master list") ?? workbook.SheetNames[0];
      loadSheet(workbook, preferred);
    } catch {
      setMessage("The file could not be read. Use a standard .xlsx, .xls, or .csv file.");
    }
  }

  const verifyWithAi = useCallback(async (row: EnrichedRow, signal?: AbortSignal): Promise<Partial<EnrichedRow>> => {
    const sourceSaysUs = isUS(textValue(row, fields.country));
    const response = await fetchWithRetry("/api/ai-verify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      cache: "no-store",
      signal,
      body: JSON.stringify({
        name: row.__displayName,
        npi: String(row["NPI Number"] || ""),
        npiSpecialty: String(row["NPI Primary Specialty"] || ""),
        claimedSpecialty: String(row["Claimed Specialty"] || textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail) || ""),
        organization: textValue(row, fields.company),
        organizationDomain: organizationDomain(row, fields),
        state: row.__state,
        country: textValue(row, fields.country),
        linkedInUrl: textValue(row, fields.linkedIn),
        practiceLocation: String(row["Resolved Practice Country"] || ""),
        verifyUsLocation: !sourceSaysUs,
      }),
    });
    if (response.status === 503) setAiConfigured(false);
    if (!response.ok) throw new Error(await responseErrorMessage(response, response.status === 429 ? "OpenAI rate limit reached." : "AI verification failed."));

    const result = await response.json() as AiVerificationResponse;
    const reliable = result.confidence === "high" || result.confidence === "medium";
    const identitySupported = reliable && result.identityMatch === true;
    const hasVerifiedSpecialty = identitySupported && result.isPhysician === true && Boolean(result.verifiedSpecialty.trim());
    const usLocationSupported = sourceSaysUs || result.practicesInUs === true;
    const usDoctor = identitySupported && usLocationSupported && result.isPhysician === true
      ? true
      : identitySupported && ((!sourceSaysUs && result.practicesInUs === false) || result.isPhysician === false)
        ? false
        : null;
    const identityReason = result.reason || "AI identity verification returned no explanation.";
    const usPracticeReason = result.usPracticeReason || "AI could not independently confirm current U.S. practice.";
    const specialtyReason = result.specialtyMatchReason || (result.verifiedSpecialty
      ? "AI returned a verified specialty but no comparison explanation."
      : "The specialty could not be independently verified.");
    const usDoctorReason = usDoctor === true
      ? sourceSaysUs ? "Source lists U.S.; AI confirmed physician identity and specialty." : "AI confirmed this contact as a current U.S. physician."
      : result.identityMatch === false
        ? "AI found an identity conflict; review the NPI and evidence."
        : result.isPhysician === false
          ? "AI did not confirm this contact as a physician."
          : result.practicesInUs === false
            ? "AI found the current practice is outside the U.S."
            : "AI could not confirm current U.S. physician practice.";
    const update: Partial<EnrichedRow> = {
      __aiStatus: hasVerifiedSpecialty ? "verified" : "uncertain",
      __aiConfidence: result.confidence,
      __aiReason: result.reason,
      __aiEvidenceUrls: result.evidenceUrls.join("\n"),
      __aiUsDoctor: usDoctor ?? undefined,
      "Identity Match": yesNoMaybe(reliable ? result.identityMatch : null),
      "Identity Match Reason": identityReason,
      "US Doctor": confidenceLabel(usDoctor, result.usDoctorConfidence),
      "US Doctor Reason": usDoctorReason,
      "AI Specialty Match": specialtyMatchLabel(hasVerifiedSpecialty ? result.specialtyMatchScore : null),
      "AI Specialty Match Reason": specialtyReason,
    };

    if (hasVerifiedSpecialty) {
      update["Verified Specialty"] = result.verifiedSpecialty.trim();
      update.__specialtyDecision = "confirmed";
      update["Specialty Decision"] = SPECIALTY_LABELS.confirmed;
      update["Specialty Evidence"] = `ChatGPT web verification: ${result.reason}`;
    }
    if (!sourceSaysUs && identitySupported && result.practicesInUs === true) {
      update.__countryDecision = "confirmed-us";
      update.__countryReason = "AI confirmed current U.S. practice.";
      update["Resolved Practice Country"] = "United States";
      update["Country Resolution"] = COUNTRY_LABELS["confirmed-us"];
      update["Country Evidence"] = usPracticeReason;
    }
    if (!sourceSaysUs && identitySupported && result.practicesInUs === false) {
      update.__countryDecision = "confirmed-non-us";
      update.__countryReason = "AI found the current practice is outside the U.S.";
      update["Resolved Practice Country"] = "Non-US";
      update["Country Resolution"] = COUNTRY_LABELS["confirmed-non-us"];
      update["Country Evidence"] = usPracticeReason;
    }
    if (reliable && result.identityMatch === false) {
      const currentNpi = String(row["NPI Number"] || "");
      update["Suggested NPI"] = currentNpi;
      update["NPI Number"] = "";
      update["NPI Match Status"] = "Identity conflict — review candidate";
      update.__doctorDecision = "ambiguous";
      update.__doctorReason = identityReason;
      update["Doctor Confirmation"] = DOCTOR_LABELS.ambiguous;
      update["Doctor Confirmation Reason"] = identityReason;
      update.__status = "review";
      update.__specialtyDecision = "insufficient";
      update["Specialty Decision"] = SPECIALTY_LABELS.insufficient;
      update["Verified Specialty"] = "";
      if (!textValue(row, fields.country)) {
        update.__countryDecision = "ambiguous";
        update.__countryReason = "The proposed NPI belongs to a different person, so it cannot establish this contact's country.";
        update["Resolved Practice Country"] = "";
        update["Country Resolution"] = COUNTRY_LABELS.ambiguous;
        update["Country Evidence"] = identityReason;
      }
    } else if (identitySupported && result.isPhysician === false) {
      update.__doctorDecision = "not-physician";
      update.__doctorReason = identityReason;
      update["Doctor Confirmation"] = DOCTOR_LABELS["not-physician"];
      update["Doctor Confirmation Reason"] = identityReason;
      update.__status = "not-physician";
    }
    return update;
  }, [fields]);

  const verifyWithAiCached = useCallback((row: EnrichedRow, signal?: AbortSignal) => {
    const key = JSON.stringify([
      String(row["NPI Number"] || ""),
      normalize(row.__displayName),
      normalize(row["Claimed Specialty"]),
      normalize(row["NPI Primary Specialty"]),
      normalize(textValue(row, fields.company)),
      normalize(row.__state),
      normalize(textValue(row, fields.country)),
      normalize(textValue(row, fields.linkedIn)),
    ]);
    const cached = aiVerificationCacheRef.current.get(key);
    if (cached) return cached;
    const request = verifyWithAi(row, signal);
    aiVerificationCacheRef.current.set(key, request);
    void request.catch(() => {
      if (aiVerificationCacheRef.current.get(key) === request) aiVerificationCacheRef.current.delete(key);
    });
    return request;
  }, [fields, verifyWithAi]);

  const discoverNpiWithAi = useCallback(async (row: EnrichedRow, npiCandidates: NpiResult[], signal?: AbortSignal): Promise<AiNpiDiscoveryResponse> => {
    const response = await fetchWithRetry("/api/ai-npi-search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      cache: "no-store",
      signal,
      body: JSON.stringify({
        name: row.__displayName,
        organization: textValue(row, fields.company),
        organizationDomain: organizationDomain(row, fields),
        state: row.__state,
        country: textValue(row, fields.country),
        claimedSpecialty: textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail),
        linkedInUrl: textValue(row, fields.linkedIn),
        npiCandidates: npiCandidates.map((candidate) => {
          const taxonomy = primaryTaxonomy(candidate);
          const address = practiceAddress(candidate);
          return {
            number: String(candidate.number ?? ""),
            registryName: registryName(candidate),
            specialty: taxonomy?.desc ?? "",
            city: address?.city ?? "",
            state: address?.state ?? taxonomy?.state ?? "",
          };
        }),
      }),
    });
    if (response.status === 503) setAiConfigured(false);
    if (!response.ok) throw new Error(await responseErrorMessage(response, response.status === 429 ? "OpenAI rate limit reached." : "AI NPI search failed."));
    return response.json() as Promise<AiNpiDiscoveryResponse>;
  }, [fields]);

  const enrichOne = useCallback(
    async (row: EnrichedRow, signal?: AbortSignal): Promise<Partial<EnrichedRow>> => {
      const firstName = textValue(row, fields.firstName);
      const lastName = textValue(row, fields.lastName);
      const state = textValue(row, fields.state);
      const sourceCountry = textValue(row, fields.country);
      if (!firstName || !lastName) {
        return {
          __status: "review",
          __doctorDecision: "unresolved",
          __doctorReason: "First and last name are required for NPPES lookup.",
          "NPI Match Status": "Missing first or last name",
          "Doctor Confirmation": DOCTOR_LABELS.unresolved,
          "Doctor Confirmation Reason": "First and last name are required for NPPES lookup.",
          "Verified Specialty": "",
          "Claimed Specialty": textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail),
          "Identity Match": "Maybe",
          "Identity Match Reason": "First and last name are required for identity verification.",
          "US Doctor": "Maybe · 0%",
          "US Doctor Reason": "First and last name are required for identity verification.",
          "AI Specialty Match": "",
          "AI Specialty Match Reason": "Specialty verification cannot run without a complete name.",
          "Specialty Decision": SPECIALTY_LABELS.unreviewed,
          "LinkedIn Verification URL": linkedinUrl(row, fields),
          "Web Verification URL": webUrl(row, fields),
          "Country Research URL": locationUrl(row, fields),
          "Enriched At (UTC)": new Date().toISOString(),
        };
      }

      const nameParams = new URLSearchParams({ first_name: firstName, last_name: lastName });
      if (state) nameParams.set("state", state);
      let npiResults = await requestNpiResults(nameParams, signal);
      let stateFilterRelaxed = false;
      let aiDiscovery: AiNpiDiscoveryResponse | undefined;
      let aiDiscoveryNote = "";
      let aiDiscoveryOutcome: "not-run" | "verified" | "not-found" | "rejected" | "error" = "not-run";

      if (!npiResults.length && state) {
        npiResults = await requestNpiResults(new URLSearchParams({ first_name: firstName, last_name: lastName }), signal);
        stateFilterRelaxed = npiResults.length > 0;
      }

      if (!npiResults.length && aiConfigured) {
        try {
          const broadParams = new URLSearchParams({ first_name: firstName });
          if (state) broadParams.set("state", state);
          let broadCandidates = await requestNpiResults(broadParams, signal);
          if (!broadCandidates.length && state) {
            broadCandidates = await requestNpiResults(new URLSearchParams({ first_name: firstName }), signal);
          }
          aiDiscovery = await discoverNpiWithAi(row, broadCandidates, signal);
          const reliable = aiDiscovery.confidence === "high" || aiDiscovery.confidence === "medium";
          if (reliable && isValidNpi(aiDiscovery.candidateNpi)) {
            const exactResults = await requestNpiResults(new URLSearchParams({ number: aiDiscovery.candidateNpi }), signal);
            const exact = exactResults.find((result) => String(result.number ?? "") === aiDiscovery?.candidateNpi);
            const registryIdentityMatches = Boolean(
              exact
              && normalize(exact.basic?.first_name) === normalize(aiDiscovery.registryFirstName)
              && normalize(exact.basic?.last_name) === normalize(aiDiscovery.registryLastName),
            );
            if (exact && registryIdentityMatches && exact.enumeration_type === "NPI-1") {
              npiResults = [exact];
              aiDiscoveryOutcome = "verified";
              aiDiscoveryNote = `AI found the registry identity${aiDiscovery.alternateName ? ` through ${aiDiscovery.alternateName}` : ""}; NPI verified directly in NPPES.`;
            } else {
              aiDiscoveryOutcome = "rejected";
              aiDiscoveryNote = "The AI candidate was rejected because its number or registry name did not match NPPES.";
            }
          } else {
            aiDiscoveryOutcome = "not-found";
            aiDiscoveryNote = aiDiscovery.reason || "AI research did not find a sufficiently reliable NPI candidate.";
          }
        } catch (error) {
          if (signal?.aborted) throw error;
          aiDiscoveryOutcome = "error";
          const detail = error instanceof Error ? error.message : "Unknown AI error.";
          aiDiscoveryNote = `AI research could not be completed: ${detail}`;
        }
      }

      const aiMatched = aiDiscoveryOutcome === "verified";
      const candidates = npiResults
        .map((result) => ({ result, score: scoreCandidate(result, row, fields) }))
        .map((candidate) => aiMatched
          ? { ...candidate, score: Math.max(candidate.score, aiDiscovery?.confidence === "high" ? 90 : 82) }
          : candidate)
        .sort((a, b) => b.score - a.score);

      if (!candidates.length) {
        const country = countryAssessment(sourceCountry, undefined, false);
        const noMatchReason = aiConfigured
          ? aiDiscoveryNote || "Exact NPPES searches and AI research found no verifiable candidate."
          : "No exact NPPES match was found. Enable ChatGPT verification to research alternate names and possible NPIs.";
        const noMatchStatus = aiDiscoveryOutcome === "error"
          ? "No exact NPPES match; AI search unavailable"
          : aiDiscoveryOutcome === "rejected"
            ? "No exact NPPES match; AI candidate rejected by NPPES"
            : aiDiscoveryOutcome === "not-found"
              ? "No exact NPPES match; AI found no verified candidate"
              : "No exact NPPES match found";
        return {
          __status: "not-found",
          __countryDecision: country.decision,
          __countryReason: country.reason,
          __doctorDecision: "unresolved",
          __doctorReason: noMatchReason,
          __specialtyDecision: "insufficient",
          __npiEvidenceUrls: aiDiscovery?.evidenceUrls.join("\n") || "",
          "NPI Match Status": noMatchStatus,
          "NPI Candidate Count": 0,
          "Country Resolution": COUNTRY_LABELS[country.decision],
          "Country Evidence": country.reason,
          "Doctor Confirmation": DOCTOR_LABELS.unresolved,
          "Doctor Confirmation Reason": noMatchReason,
          "Verified Specialty": "",
          "Claimed Specialty": textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail),
          "Identity Match": "Maybe",
          "Identity Match Reason": noMatchReason,
          "US Doctor": "Maybe · 20%",
          "US Doctor Reason": "No verified NPI or identity link was found.",
          "AI Specialty Match": "",
          "AI Specialty Match Reason": "No NPI specialty is available to compare.",
          "Specialty Comparison": "No NPI specialty is available.",
          "Specialty Decision": SPECIALTY_LABELS.insufficient,
          "LinkedIn Verification URL": linkedinUrl(row, fields),
          "Web Verification URL": webUrl(row, fields),
          "Country Research URL": locationUrl(row, fields),
          "Enriched At (UTC)": new Date().toISOString(),
        };
      }

      const best = candidates[0];
      const runnerUp = candidates[1];
      const clearMatch = best.score >= 75 && (!runnerUp || best.score - runnerUp.score >= 15);
      const taxonomy = primaryTaxonomy(best.result);
      const doctor = doctorAssessment(best.result, clearMatch);
      const country = countryAssessment(sourceCountry, best.result, clearMatch);
      const npi = String(best.result.number ?? "");
      const claimed = textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail);
      const registryUrl = npi ? `https://npiregistry.cms.hhs.gov/provider-view/${npi}` : "";
      const status: EnrichedRow["__status"] = !clearMatch
        ? "review"
        : doctor.decision === "confirmed"
          ? "verified"
          : doctor.decision === "not-physician"
            ? "not-physician"
            : "review";
      const practice = practiceAddress(best.result);
      const baselineUsDoctor = clearMatch && doctor.decision === "confirmed" && ["confirmed-us", "likely-us"].includes(country.decision)
        ? "Yes · 80%"
        : country.decision === "confirmed-non-us" || doctor.decision === "not-physician"
          ? "No · 90%"
          : "Maybe · 45%";

      const baseResult: Partial<EnrichedRow> = {
        __status: status,
        __countryDecision: country.decision,
        __countryReason: country.reason,
        __doctorDecision: doctor.decision,
        __doctorReason: doctor.reason,
        __specialtyDecision: "unreviewed",
        "NPI Number": clearMatch ? npi : "",
        "Suggested NPI": clearMatch ? "" : npi,
        __npiEvidenceUrls: aiDiscovery?.evidenceUrls.join("\n") || "",
        "NPI Match Status": clearMatch
          ? aiMatched
            ? "AI-discovered identity — NPI verified in NPPES"
            : stateFilterRelaxed
              ? "High-confidence match after retry without state"
              : "High-confidence match"
          : "Ambiguous — review candidate",
        "NPI Match Score": best.score,
        "NPI Status": activeNpi(best.result) ? "Active" : "Deactivated",
        "NPI Entity Type": best.result.enumeration_type === "NPI-1" ? "Individual" : "Organization",
        "NPI Primary Specialty": taxonomy?.desc ?? "",
        "NPI Taxonomy Code": taxonomy?.code ?? "",
        "NPI Credential": best.result.basic?.credential ?? "",
        "NPI Registry Name": registryName(best.result),
        "NPI License State": taxonomy?.state ?? "",
        "NPI Candidate Count": candidates.length,
        "Resolved Practice Country": ["confirmed-us", "likely-us"].includes(country.decision) ? "United States" : practice?.country_name || practice?.country_code || "",
        "Country Resolution": COUNTRY_LABELS[country.decision],
        "Country Evidence": country.reason,
        "Doctor Confirmation": DOCTOR_LABELS[doctor.decision],
        "Doctor Confirmation Reason": doctor.reason,
        "Verified Specialty": "",
        "Claimed Specialty": claimed,
        "Registered NPI Specialty": taxonomy?.desc ?? "",
        "Identity Match": "Maybe",
        "Identity Match Reason": clearMatch
          ? "NPPES produced a strong registry match; AI identity verification may confirm the exact person."
          : "The NPPES candidate is ambiguous and needs review.",
        "US Doctor": baselineUsDoctor,
        "US Doctor Reason": baselineUsDoctor.startsWith("Yes")
          ? "NPPES supports a U.S. physician match."
          : baselineUsDoctor.startsWith("No")
            ? "NPPES does not support a current U.S. physician match."
            : "The NPI, physician status, or U.S. practice still needs confirmation.",
        "AI Specialty Match": "",
        "AI Specialty Match Reason": "AI specialty comparison has not been completed.",
        "Specialty Comparison": specialtyAssessment(claimed, best.result),
        "Specialty Decision": SPECIALTY_LABELS.unreviewed,
        "Specialty Evidence": `NPPES primary taxonomy${taxonomy?.desc ? `: ${taxonomy.desc}` : " unavailable"}`,
        "NPI Registry URL": registryUrl,
        "LinkedIn Verification URL": linkedinUrl(row, fields),
        "Web Verification URL": webUrl(row, fields),
        "Country Research URL": locationUrl(row, fields),
        "Enriched At (UTC)": new Date().toISOString(),
      };
      if (aiConfigured && clearMatch && doctor.decision === "confirmed") {
        try {
          const aiResult = await verifyWithAiCached({ ...row, ...baseResult } as EnrichedRow, signal);
          return { ...baseResult, ...aiResult };
        } catch (error) {
          if (signal?.aborted) throw error;
          const detail = error instanceof Error ? error.message : "Unknown AI error.";
          return {
            ...baseResult,
            __aiStatus: "error",
            __aiReason: `ChatGPT web verification could not be completed: ${detail}`,
          };
        }
      }
      return baseResult;
    },
    [aiConfigured, discoverNpiWithAi, fields, verifyWithAiCached],
  );

  const runEnrichment = useCallback(
    async (limit?: number) => {
      if (isRunning || isAiRunning || !eligibleRows.length) return;
      const targets = eligibleRows.filter((row) => ["queued", "error", "not-found"].includes(row.__status)).slice(0, limit);
      if (!targets.length) {
        setMessage("Every eligible row in this scope has already been checked.");
        return;
      }
      cancelRef.current = false;
      const controller = new AbortController();
      abortControllerRef.current = controller;
      setIsRunning(true);
      setProcessed(0);
      setRunTotal(targets.length);
      setMessage(aiConfigured
        ? `Checking ${targets.length.toLocaleString()} contacts against NPPES, then running ChatGPT web verification for confirmed matches…`
        : `Checking ${targets.length.toLocaleString()} contacts against NPPES…`);

      const targetIndexes = new Set(targets.map((row) => row.__index));
      setRows((current) => current.map((row) => targetIndexes.has(row.__index) ? { ...row, __status: "running" } : row));

      let cursor = 0;
      let finished = 0;
      let aiFailures = 0;
      const worker = async () => {
        while (!cancelRef.current) {
          const position = cursor;
          cursor += 1;
          if (position >= targets.length) return;
          const target = targets[position];
          try {
            const result = await enrichOne(target, controller.signal);
            if (result.__aiStatus === "error" || String(result["NPI Match Status"] || "").includes("AI search unavailable")) {
              aiFailures += 1;
            }
            setRows((current) => current.map((row) => row.__index === target.__index ? { ...row, ...result } : row));
          } catch {
            if (controller.signal.aborted) return;
            setRows((current) => current.map((row) => row.__index === target.__index ? {
              ...row,
              __status: "error",
              "NPI Match Status": "API error — retry",
              __doctorDecision: "unresolved",
              __doctorReason: "The CMS request failed; retry the NPI lookup.",
              "Doctor Confirmation": DOCTOR_LABELS.unresolved,
              "Doctor Confirmation Reason": "The CMS request failed; retry the NPI lookup.",
              "LinkedIn Verification URL": linkedinUrl(row, fields),
              "Web Verification URL": webUrl(row, fields),
              "Country Research URL": locationUrl(row, fields),
            } : row));
          }
          finished += 1;
          setProcessed(finished);
        }
      };

      await Promise.all([worker(), worker(), worker(), worker()]);
      if (cancelRef.current) {
        setRows((current) => current.map((row) =>
          targetIndexes.has(row.__index) && row.__status === "running"
            ? { ...row, __status: "queued" }
            : row,
        ));
      }
      setIsRunning(false);
      if (abortControllerRef.current === controller) abortControllerRef.current = null;
      setMessage(cancelRef.current
        ? `Stopped after ${finished.toLocaleString()} contacts. Completed results are retained; press Enrich all eligible to continue.`
        : aiConfigured
          ? `Completed ${finished.toLocaleString()} NPI checks.${aiFailures ? ` ${aiFailures.toLocaleString()} AI request${aiFailures === 1 ? "" : "s"} failed; see the affected row for the exact error.` : " ChatGPT research completed where needed."}`
          : `Completed ${finished.toLocaleString()} NPI checks. Add OPENAI_API_KEY to activate ChatGPT web verification.`);
    },
    [aiConfigured, eligibleRows, enrichOne, fields, isAiRunning, isRunning],
  );

  const runAiVerification = useCallback(async () => {
    if (isRunning || isAiRunning || !aiConfigured) return;
    const targets = eligibleRows.filter((row) =>
      row.__doctorDecision === "confirmed"
      && Boolean(row["NPI Number"])
      && row.__aiStatus !== "verified"
      && !row.__manualSpecialty,
    );
    if (!targets.length) {
      setMessage("Every matched physician has already completed ChatGPT web verification.");
      return;
    }

    cancelRef.current = false;
    const controller = new AbortController();
    abortControllerRef.current = controller;
    setIsAiRunning(true);
    setProcessed(0);
    setRunTotal(targets.length);
    setMessage(`ChatGPT is researching ${targets.length.toLocaleString()} matched physicians…`);
    const targetIndexes = new Set(targets.map((row) => row.__index));
    setRows((current) => current.map((row) => targetIndexes.has(row.__index) ? { ...row, __aiStatus: "checking" } : row));

    let cursor = 0;
    let finished = 0;
    const worker = async () => {
      while (!cancelRef.current) {
        const position = cursor;
        cursor += 1;
        if (position >= targets.length) return;
        const target = targets[position];
        try {
          const result = await verifyWithAiCached(target, controller.signal);
          setRows((current) => current.map((row) => row.__index === target.__index ? { ...row, ...result } : row));
        } catch (error) {
          if (controller.signal.aborted) return;
          const detail = error instanceof Error ? error.message : "Unknown AI error.";
          setRows((current) => current.map((row) => row.__index === target.__index ? {
            ...row,
            __aiStatus: "error",
            __aiReason: `ChatGPT web verification failed: ${detail}`,
          } : row));
        }
        finished += 1;
        setProcessed(finished);
      }
    };

    await Promise.all([worker(), worker(), worker(), worker()]);

    if (cancelRef.current) {
      setRows((current) => current.map((row) => row.__aiStatus === "checking" ? { ...row, __aiStatus: "not-run" } : row));
    }
    setIsAiRunning(false);
    if (abortControllerRef.current === controller) abortControllerRef.current = null;
    setMessage(cancelRef.current
      ? `Stopped after ${finished.toLocaleString()} AI checks. Completed results are retained; press AI verify matched to continue.`
      : `Completed ${finished.toLocaleString()} ChatGPT web checks. Only confirmed U.S. doctors are included in the export.`);
  }, [aiConfigured, eligibleRows, isAiRunning, isRunning, verifyWithAiCached]);

  function stopActiveRun() {
    cancelRef.current = true;
    abortControllerRef.current?.abort();
    abortControllerRef.current = null;
    setRows((current) => current.map((row) => ({
      ...row,
      __status: row.__status === "running" ? "queued" : row.__status,
      __aiStatus: row.__aiStatus === "checking" ? "not-run" : row.__aiStatus,
    })));
    setIsRunning(false);
    setIsAiRunning(false);
    setMessage("Stopped. Completed results were kept; unfinished contacts are ready to continue.");
  }

  function openReview(row: EnrichedRow) {
    setReviewDialogRow(row);
    setReviewDraft({
      country: row.__countryDecision,
      doctor: row.__doctorDecision,
      specialty: row.__specialtyDecision,
      evidenceUrl: String(row["Manual Evidence URL"] || ""),
      notes: String(row["Reviewer Notes"] || ""),
    });
  }

  function saveReview() {
    if (!reviewDialogRow) return;
    setRows((current) => current.map((row) => row.__index === reviewDialogRow.__index ? {
      ...row,
      __countryDecision: reviewDraft.country,
      __doctorDecision: reviewDraft.doctor,
      __specialtyDecision: reviewDraft.specialty,
      __countryReason: reviewDraft.notes.trim() ? `Manual review: ${reviewDraft.notes.trim()}` : `${row.__countryReason} Manually reviewed.`,
      __doctorReason: reviewDraft.notes.trim() ? `Manual review: ${reviewDraft.notes.trim()}` : `${row.__doctorReason} Manually reviewed.`,
      __manualCountry: true,
      __manualDoctor: true,
      __manualSpecialty: true,
      __status: reviewDraft.doctor === "confirmed" ? "verified" : reviewDraft.doctor === "not-physician" ? "not-physician" : "review",
      "Resolved Practice Country": ["confirmed-us", "likely-us"].includes(reviewDraft.country) ? "United States" : reviewDraft.country === "confirmed-non-us" ? "Non-US" : "",
      "Country Resolution": `${COUNTRY_LABELS[reviewDraft.country]} — manual review`,
      "Doctor Confirmation": `${DOCTOR_LABELS[reviewDraft.doctor]} — manual review`,
      "Doctor Confirmation Reason": reviewDraft.notes || row["Doctor Confirmation Reason"] || "Reviewed manually.",
      "US Doctor": reviewDraft.doctor === "confirmed" && ["confirmed-us", "likely-us"].includes(reviewDraft.country)
        ? "Yes · 100%"
        : reviewDraft.doctor === "not-physician" || reviewDraft.country === "confirmed-non-us"
          ? "No · 100%"
          : "Maybe · 50%",
      "US Doctor Reason": reviewDraft.notes || "Determined by manual country and physician review.",
      "Specialty Decision": `${SPECIALTY_LABELS[reviewDraft.specialty]} — manual review`,
      "Manual Evidence URL": reviewDraft.evidenceUrl.trim(),
      "Reviewer Notes": reviewDraft.notes.trim(),
      "Reviewed At (UTC)": new Date().toISOString(),
    } : row));
    setMessage(`Review decisions were saved for ${reviewDialogRow.__displayName}.`);
    setReviewDialogRow(null);
  }

  function excludeRow(row: EnrichedRow) {
    setRows((current) => current.map((item) => item.__index === row.__index ? { ...item, __excluded: true } : item));
    setLastExcluded(row.__index);
    setMessage(`${row.__displayName} was removed from the queue and exports.`);
  }

  function undoExclude() {
    if (lastExcluded === null) return;
    setRows((current) => current.map((item) => item.__index === lastExcluded ? { ...item, __excluded: false } : item));
    setLastExcluded(null);
    setMessage("The last removed contact was restored.");
  }

  function openManualNpi(row: EnrichedRow) {
    setNpiDialogRow(row);
    setManualNpi(String(row["NPI Number"] || row["Suggested NPI"] || ""));
    setManualNpiError("");
  }

  function saveManualNpi() {
    const npi = manualNpi.replace(/\D/g, "");
    if (!npiDialogRow || !isValidNpi(npi)) {
      setManualNpiError("Enter a valid 10-digit NPI with a correct check digit.");
      return;
    }
    setRows((current) => current.map((row) => row.__index === npiDialogRow.__index ? {
      ...row,
      __status: "review",
      __doctorDecision: "unresolved",
      __doctorReason: "The NPI was entered manually and has not been matched to this person.",
      __specialtyDecision: "unreviewed",
      "NPI Number": npi,
      "Suggested NPI": "",
      "NPI Match Status": "Manually entered",
      "NPI Match Score": "Manual",
      "NPI Status": "Not checked",
      "Doctor Confirmation": DOCTOR_LABELS.unresolved,
      "Doctor Confirmation Reason": "The NPI was entered manually and has not been matched to this person.",
      "Identity Match": "Maybe",
      "Identity Match Reason": "The NPI was entered manually and has not been matched to this person.",
      "US Doctor": "Maybe · 25%",
      "US Doctor Reason": "The manually entered NPI still needs physician and U.S. practice verification.",
      "AI Specialty Match": "",
      "AI Specialty Match Reason": "AI specialty comparison has not been run for this NPI.",
      "NPI Registry URL": `https://npiregistry.cms.hhs.gov/provider-view/${npi}`,
      "Verified Specialty": "",
      "Specialty Decision": SPECIALTY_LABELS.unreviewed,
      "Enriched At (UTC)": new Date().toISOString(),
    } : row));
    setMessage(`NPI ${npi} was added manually for ${npiDialogRow.__displayName}.`);
    setNpiDialogRow(null);
    setManualNpi("");
  }

  function addContact() {
    const state = newContact.state.trim().toUpperCase();
    const countryValue = newContact.country.trim();
    const countryCheck = initialCountryDecision(countryValue);
    const npi = newContact.npi.replace(/\D/g, "");
    if (!newContact.firstName.trim() || !newContact.lastName.trim() || !newContact.license.trim()) {
      setNewContactError("First name, last name, and license key are required.");
      return;
    }
    if (state && !/^[A-Z]{2}$/.test(state)) {
      setNewContactError("Use a two-letter U.S. state code or leave the state blank for research.");
      return;
    }
    if (npi && !isValidNpi(npi)) {
      setNewContactError("The optional NPI must be a valid 10-digit number with a correct check digit.");
      return;
    }

    const resolved = Object.fromEntries(
      Object.entries(DEFAULT_FIELDS).map(([key, fallback]) => [key, fields[key as keyof FieldMap] ?? fallback]),
    ) as Required<FieldMap>;
    const emailHeader = findHeader(headers, ["email address", "email"]) ?? "Email Address";
    const nextIndex = rows.reduce((max, row) => Math.max(max, row.__index), -1) + 1;
    const base: ContactRow = {
      [emailHeader]: newContact.email.trim(),
      [resolved.fullName]: `${newContact.firstName.trim()} ${newContact.lastName.trim()}`,
      [resolved.firstName]: newContact.firstName.trim(),
      [resolved.lastName]: newContact.lastName.trim(),
      [resolved.license]: newContact.license.trim(),
      [resolved.country]: countryValue,
      [resolved.state]: state,
      [resolved.company]: newContact.company.trim(),
      [resolved.claimedSpecialty]: newContact.specialty.trim(),
      [resolved.linkedIn]: newContact.linkedIn.trim(),
      "Record Source": "Added manually",
    };
    const added: EnrichedRow = {
      ...base,
      __index: nextIndex,
      __eligible: true,
      __manualAdded: true,
      __displayName: `${newContact.firstName.trim()} ${newContact.lastName.trim()}`,
      __state: state,
      __status: npi ? "review" : "queued",
      __countryDecision: countryCheck.decision,
      __countryReason: countryCheck.reason,
      __doctorDecision: "unresolved",
      __doctorReason: npi ? "Manual NPI needs identity and taxonomy review." : "NPI lookup has not been run.",
      __specialtyDecision: "unreviewed",
      "NPI Number": npi,
      "NPI Match Status": npi ? "Manually entered" : "Manually added — NPI not checked",
      "NPI Match Score": npi ? "Manual" : "",
      "NPI Status": npi ? "Not checked" : "",
      "Resolved Practice Country": ["confirmed-us", "likely-us"].includes(countryCheck.decision) ? "United States" : countryCheck.decision === "confirmed-non-us" ? countryValue : "",
      "Country Resolution": COUNTRY_LABELS[countryCheck.decision],
      "Country Evidence": countryCheck.reason,
      "Doctor Confirmation": DOCTOR_LABELS.unresolved,
      "Doctor Confirmation Reason": npi ? "Manual NPI needs identity and taxonomy review." : "NPI lookup has not been run.",
      "Verified Specialty": "",
      "Claimed Specialty": newContact.specialty.trim(),
      "Identity Match": "Maybe",
      "Identity Match Reason": npi ? "A manually entered NPI still needs identity verification." : "Identity verification has not been run.",
      "US Doctor": "Maybe · 25%",
      "US Doctor Reason": npi ? "A manually entered NPI still needs physician and U.S. practice verification." : "NPI and identity verification have not been run.",
      "AI Specialty Match": "",
      "AI Specialty Match Reason": "AI specialty comparison has not been run.",
      "Specialty Decision": SPECIALTY_LABELS.unreviewed,
      "NPI Registry URL": npi ? `https://npiregistry.cms.hhs.gov/provider-view/${npi}` : "",
      "LinkedIn Verification URL": linkedinUrl(base, resolved),
      "Web Verification URL": webUrl(base, resolved),
      "Country Research URL": locationUrl(base, resolved),
      "Enriched At (UTC)": new Date().toISOString(),
    };

    setRows((current) => [...current, added]);
    setFields(resolved);
    setHeaders((current) => Array.from(new Set([...current, ...Object.keys(base)])));
    setNewContact(EMPTY_CONTACT);
    setNewContactError("");
    setAddDialogOpen(false);
    setPage(1);
    setStatusFilter("all");
    setSearchQuery("");
    setMessage(`${added.__displayName} was added to the queue and will be included in exports.`);
  }

  useEffect(() => {
    const context = document.modelContext;
    if (!context?.registerTool) return;
    const lifecycle = new AbortController();
    void Promise.resolve(context.registerTool({
      name: "get_enrichment_summary",
      title: "Get enrichment summary",
      description: "Return counts for uploaded, eligible, processed, verified, and review-required contacts.",
      inputSchema: { type: "object", properties: {}, additionalProperties: false },
      annotations: { readOnlyHint: true, untrustedContentHint: false },
      execute: () => ({ fileName, totalRows: rows.length, reviewScope: eligibleRows.length, processedRows: completeRows.length, confirmedPhysicians: verified, countryNeedsReview: countryReview, specialtyNeedsReview: specialtyReview, needsReview, manuallyAdded: manualCount, excluded: excludedCount }),
    }, { signal: lifecycle.signal })).catch(() => undefined);
    return () => lifecycle.abort();
  }, [completeRows.length, countryReview, eligibleRows.length, excludedCount, fileName, manualCount, needsReview, rows.length, specialtyReview, verified]);

  const filteredRows = eligibleRows.filter((row) => {
    const query = normalize(searchQuery);
    const matchesSearch = !query || normalize([
      row.__displayName,
      textValue(row, fields.company),
      row.__state,
      row["NPI Number"],
      row["Suggested NPI"],
      row["NPI Primary Specialty"],
      row["Verified Specialty"],
      COUNTRY_LABELS[row.__countryDecision],
      DOCTOR_LABELS[row.__doctorDecision],
    ].join(" ")).includes(query);
    const matchesStatus = statusFilter === "all"
      || (statusFilter === "queued" && ["queued", "running"].includes(row.__status))
      || (statusFilter === "verified" && row.__doctorDecision === "confirmed")
      || (statusFilter === "not-physician" && row.__doctorDecision === "not-physician")
      || (statusFilter === "country-review" && ["unresolved", "ambiguous"].includes(row.__countryDecision))
      || (statusFilter === "specialty-review" && ["unreviewed", "insufficient"].includes(row.__specialtyDecision))
      || (statusFilter === "review" && ["review", "not-found", "error"].includes(row.__status))
      || (statusFilter === "manual" && row.__manualAdded);
    return matchesSearch && matchesStatus;
  });
  const pageSizeNumber = pageSize === "all" ? Math.max(filteredRows.length, 1) : Number(pageSize);
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSizeNumber));
  const safePage = Math.min(page, totalPages);
  const visibleRows = filteredRows.slice((safePage - 1) * pageSizeNumber, safePage * pageSizeNumber);

  useEffect(() => {
    setPage(1);
  }, [searchQuery, statusFilter, pageSize, eligibleRows.length]);

  const hasRequiredFields = Boolean(fields.firstName && fields.lastName && fields.license);
  const progress = runTotal ? Math.round((processed / runTotal) * 100) : 0;

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="topbar">
        <div className="brand-mark" aria-hidden="true"><Stethoscope /></div>
        <div>
          <p className="eyebrow">Chief Commercialization Office</p>
          <h1>NPI Contact Enrichment</h1>
        </div>
        <div className="ml-auto hidden items-center gap-2 text-sm text-slate-300 sm:flex">
          <ShieldCheck className="size-4 text-cyan-300" />
          Private working file · no contact data stored
        </div>
      </header>

      <div className="workspace-shell">
        <aside className="workflow-rail" aria-label="Enrichment workflow">
          <p className="rail-label">Workflow</p>
          {[
            ["01", "Upload", Boolean(rows.length)],
            ["02", "Country check", countryReview === 0 && rows.length > 0],
            ["03", "NPI match", completeRows.length > 0],
            ["04", "Doctor + AI specialty", aiVerified > 0],
            ["05", "Export US contacts", usContactRows.length > 0],
          ].map(([number, label, done]) => (
            <div className="rail-step" key={String(number)}>
              <span className={done ? "rail-number done" : "rail-number"}>{done ? <Check /> : number}</span>
              <span>{label}</span>
            </div>
          ))}
          <div className="rail-note">
            <Database className="size-4" />
            <p>An NPI identifies a healthcare provider; it does not by itself prove the person is a doctor. Physician taxonomy and identity must also match.</p>
          </div>
        </aside>

        <section className="content-column">
          <div className="section-heading">
            <div>
              <p className="eyebrow text-cyan-700">Doctor prospecting list</p>
              <h2>Validate licensed U.S. contacts before CRM import</h2>
            </div>
            <a className="source-link" href="https://npiregistry.cms.hhs.gov/api-page" target="_blank" rel="noreferrer">
              NPPES API reference <ArrowUpRight />
            </a>
          </div>

          <section className="upload-panel">
            <label className="drop-zone">
              <input type="file" accept=".xlsx,.xls,.csv" onChange={(event) => void handleFile(event.target.files?.[0])} className="sr-only" />
              <span className="upload-icon"><UploadCloud /></span>
              <span className="font-semibold">{fileName || "Choose a contact file"}</span>
              <span className="text-sm text-slate-500">Excel or CSV · processed in this browser · 30 MB maximum</span>
            </label>

            <div className="detection-grid">
              <div><span>Excel tab</span><strong>{sheetName || "—"}</strong></div>
              <div><span>Rows</span><strong>{rows.length.toLocaleString()}</strong></div>
              <div><span>License field</span><strong>{fields.license ? `${fields.license} (Column ${licenseColumn})` : "Not detected"}</strong></div>
              <div><span>Review scope</span><strong>{eligibleRows.length.toLocaleString()} licensed U.S. or blank-country contacts</strong></div>
            </div>

            {sheetNames.length > 1 && (
              <details className="sheet-options">
                <summary>Choose a different Excel tab</summary>
                <div className="sheet-picker">
                  <div><span>Excel tab to process</span><small>“Master list” is selected automatically when present.</small></div>
                  <Select value={sheetName} onValueChange={(value) => workbookRef.current && loadSheet(workbookRef.current, value)}>
                    <SelectTrigger className="w-[min(340px,100%)] bg-white"><SelectValue placeholder="Choose Excel tab" /></SelectTrigger>
                    <SelectContent>
                      {sheetNames.map((name) => <SelectItem value={name} key={name}>{name}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
              </details>
            )}
          </section>

          <section className="status-strip" aria-live="polite">
            {isRunning ? <LoaderCircle className="size-5 animate-spin" /> : rows.length ? <FileCheck2 className="size-5" /> : <AlertCircle className="size-5" />}
            <p>{message}</p>
          </section>

          <section className="metric-row" aria-label="Enrichment summary">
            <div className="metric-card"><span>In review scope</span><strong>{eligibleRows.length.toLocaleString()}</strong><small>License + U.S. or blank country</small></div>
            <div className="metric-card review"><span>Country unresolved</span><strong>{countryReview.toLocaleString()}</strong><small>Blank or ambiguous</small></div>
            <div className="metric-card verified"><span>Confirmed doctors</span><strong>{verified.toLocaleString()}</strong><small>Identity + taxonomy</small></div>
            <div className="metric-card"><span>AI web verified</span><strong>{aiVerified.toLocaleString()}</strong><small>Identity and specialty evidence</small></div>
          </section>

          <section className="decision-guide" aria-label="What each check means">
            <div><MapPin /><span><strong>U.S. practice</strong><small>Separate from identity and doctor status</small></span></div>
            <div><SearchCheck /><span><strong>Identity match</strong><small>Checks that the NPI belongs to this exact contact</small></span></div>
            <div><UserCheck /><span><strong>NPI primary specialty</strong><small>The primary taxonomy registered in NPPES</small></span></div>
            <div><Sparkles /><span><strong>Verified specialty</strong><small>Current specialty supported by public web evidence</small></span></div>
          </section>

          <section className="action-panel">
            <div className="action-copy">
              <h3>Run NPPES and ChatGPT verification</h3>
              <p>NPPES supplies the NPI and registered taxonomy. If exact name searches fail, ChatGPT researches alternate names and possible NPIs, then every candidate is checked directly in NPPES. Email and license key stay in this browser.</p>
              <span className={`ai-config-status ${aiConfigured ? "ready" : "missing"}`}>
                {aiConfigurationChecked ? aiConfigured ? "OpenAI API key detected" : "Add OPENAI_API_KEY to activate ChatGPT verification" : "Checking ChatGPT configuration…"}
              </span>
            </div>
            <div className="button-row">
              {isRunning ? (
                <Button variant="outline" onClick={stopActiveRun}><CircleStop /> Stop</Button>
              ) : (
                <>
                  <Button variant="outline" disabled={isAiRunning || !hasRequiredFields || !eligibleRows.length} onClick={() => void runEnrichment(25)}>
                    <SearchCheck /> Test 25 rows
                  </Button>
                  <Button disabled={isAiRunning || !hasRequiredFields || !eligibleRows.length} onClick={() => void runEnrichment()}>
                    <Stethoscope /> Enrich all eligible
                  </Button>
                </>
              )}
            </div>
            {(isRunning || isAiRunning || runTotal > 0) && (
              <div className="progress-wrap"><Progress value={progress} /><span>{processed.toLocaleString()} / {runTotal.toLocaleString()}</span></div>
            )}
          </section>

          <section className="results-panel">
            <div className="results-head">
              <div>
                <h3>Contact queue</h3>
                <p>{filteredRows.length.toLocaleString()} shown from {eligibleRows.length.toLocaleString()} contacts under review. {manualCount.toLocaleString()} added manually · {excludedCount.toLocaleString()} removed.</p>
              </div>
              <div className="button-row">
                {lastExcluded !== null && <Button variant="ghost" onClick={undoExclude}><RotateCcw /> Undo remove</Button>}
                <Button variant="outline" onClick={() => setAddDialogOpen(true)}><Plus /> Add person</Button>
                <Button disabled={!usContactRows.length} onClick={() => downloadExcelWorkbook(usContactRows, exportRows)}>
                  <Download /> Download Excel · {usContactRows.length.toLocaleString()} US contacts / {exportRows.length.toLocaleString()} doctors
                </Button>
              </div>
            </div>

            <div className="results-tools">
              <label className="search-control">
                <span className="sr-only">Search contacts</span>
                <Search aria-hidden="true" />
                <Input value={searchQuery} onChange={(event) => setSearchQuery(event.target.value)} placeholder="Search name, organization, state, NPI, or specialty" />
              </label>
              <Select value={statusFilter} onValueChange={setStatusFilter}>
                <SelectTrigger className="w-[180px] bg-white"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="all">All statuses</SelectItem>
                  <SelectItem value="queued">Queued</SelectItem>
                  <SelectItem value="verified">Confirmed physicians</SelectItem>
                  <SelectItem value="not-physician">Not physicians</SelectItem>
                  <SelectItem value="country-review">Country needs review</SelectItem>
                  <SelectItem value="specialty-review">Specialty needs review</SelectItem>
                  <SelectItem value="review">NPI/identity needs review</SelectItem>
                  <SelectItem value="manual">Added manually</SelectItem>
                </SelectContent>
              </Select>
              <Select value={pageSize} onValueChange={setPageSize}>
                <SelectTrigger className="w-[145px] bg-white"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="100">100 per page</SelectItem>
                  <SelectItem value="250">250 per page</SelectItem>
                  <SelectItem value="500">500 per page</SelectItem>
                  <SelectItem value="all">Show all</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {visibleRows.length ? (
              <>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Contact</TableHead><TableHead>Country</TableHead><TableHead>NPI match</TableHead><TableHead>U.S. doctor?</TableHead><TableHead>Claimed specialty</TableHead><TableHead>NPI primary specialty</TableHead><TableHead>AI specialty verification</TableHead><TableHead>Evidence</TableHead><TableHead className="text-right">Actions</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {visibleRows.map((row) => (
                      <TableRow key={row.__index}>
                        <TableCell>
                          <div className="contact-cell">
                            <strong>{row.__displayName}{row.__manualAdded && <span className="manual-tag">Manual</span>}</strong>
                            <span>{textValue(row, fields.company) || "No organization"}{row.__state ? ` · ${row.__state}` : ""}</span>
                            <span className="mt-2"><StatusPill status={row.__status} /></span>
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="decision-cell">
                            <DecisionPill kind="country" value={row.__countryDecision} />
                            <small title={row.__countryReason}>{shortQueueReason(row.__countryReason)}</small>
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="decision-cell">
                            <div className="npi-cell">
                              <span className="font-mono text-xs">{String(row["NPI Number"] || row["Suggested NPI"] || "—")}</span>
                              <Button variant="ghost" size="icon-xs" onClick={() => openManualNpi(row)} aria-label={`Edit NPI for ${row.__displayName}`} title="Add or edit NPI"><Pencil /></Button>
                            </div>
                            <small title={String(row["NPI Match Status"] || "Not checked")}>{shortQueueReason(row["NPI Match Status"] || "Not checked")}{row["NPI Match Score"] ? ` · ${row["NPI Match Score"]}` : ""}</small>
                            {row["NPI Registry Name"] && <small><b>Registry name:</b> {String(row["NPI Registry Name"])}</small>}
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="decision-cell">
                            <AnswerPill value={row["US Doctor"]} />
                            <small title={String(row["US Doctor Reason"] || row.__doctorReason)}>{shortQueueReason(row["US Doctor Reason"] || row.__doctorReason)}</small>
                          </div>
                        </TableCell>
                        <TableCell><div className="specialty-cell"><small>{String(row["Claimed Specialty"] || textValue(row, fields.claimedSpecialty) || textValue(row, fields.categoryDetail) || "—")}</small></div></TableCell>
                        <TableCell><div className="specialty-cell"><small>{String(row["NPI Primary Specialty"] || "—")}</small></div></TableCell>
                        <TableCell>
                          <div className="specialty-cell">
                            <AnswerPill value={row["AI Specialty Match"]} />
                            <small><b>Verified:</b> {String(row["Verified Specialty"] || "—")}</small>
                            <small title={String(row["AI Specialty Match Reason"] || "AI specialty comparison has not been run.")}>{shortQueueReason(row["AI Specialty Match Reason"] || "AI specialty comparison has not been run.")}</small>
                            {row.__aiStatus === "checking" && <small><b>AI:</b> Researching public sources…</small>}
                            {row.__aiStatus === "verified" && <small><b>AI:</b> Verified · {row.__aiConfidence} confidence</small>}
                            {row.__aiStatus === "uncertain" && <small><b>AI:</b> More evidence needed · {row.__aiConfidence} confidence</small>}
                            {row.__aiStatus === "error" && <small title={row.__aiReason}><b>AI:</b> {shortQueueReason(row.__aiReason || "Verification failed; retry available")}</small>}
                          </div>
                        </TableCell>
                        <TableCell>
                          <div className="evidence-links evidence-links-stack">
                            {row["NPI Registry URL"] && <a href={String(row["NPI Registry URL"])} target="_blank" rel="noreferrer">NPI Registry <ArrowUpRight /></a>}
                            {String(row.__npiEvidenceUrls || "").split("\n").filter(Boolean).map((url, index) => (
                              <a href={url} target="_blank" rel="noreferrer" key={`${row.__index}-npi-ai-${index}`}>NPI search evidence {index + 1} <ArrowUpRight /></a>
                            ))}
                            {String(row.__aiEvidenceUrls || "").split("\n").filter(Boolean).map((url, index) => (
                              <a href={url} target="_blank" rel="noreferrer" key={`${row.__index}-ai-${index}`}>AI evidence {index + 1} <ArrowUpRight /></a>
                            ))}
                            <a href={linkedinUrl(row, fields)} target="_blank" rel="noreferrer">LinkedIn <ArrowUpRight /></a>
                            <a href={locationUrl(row, fields)} target="_blank" rel="noreferrer">Location search <ArrowUpRight /></a>
                            {row["Manual Evidence URL"] && <a href={String(row["Manual Evidence URL"])} target="_blank" rel="noreferrer">Reviewer evidence <ArrowUpRight /></a>}
                          </div>
                        </TableCell>
                        <TableCell className="text-right">
                          <div className="row-actions">
                            <Button variant="outline" size="sm" onClick={() => openReview(row)}><Pencil /> Review</Button>
                            <Button variant="ghost" size="icon-sm" className="remove-button" onClick={() => excludeRow(row)} aria-label={`Remove ${row.__displayName} from queue and exports`} title="Remove from queue and exports"><X /></Button>
                          </div>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
                <div className="pagination-bar">
                  <span>Rows {((safePage - 1) * pageSizeNumber + 1).toLocaleString()}–{Math.min(safePage * pageSizeNumber, filteredRows.length).toLocaleString()} of {filteredRows.length.toLocaleString()}</span>
                  {pageSize !== "all" && totalPages > 1 && (
                    <Pagination className="mx-0 w-auto">
                      <PaginationContent>
                        <PaginationItem><PaginationPrevious href="#" className={safePage <= 1 ? "pointer-events-none opacity-40" : ""} onClick={(event) => { event.preventDefault(); if (safePage > 1) setPage(safePage - 1); }} /></PaginationItem>
                        <PaginationItem><PaginationLink href="#" isActive onClick={(event) => event.preventDefault()}>{safePage}</PaginationLink></PaginationItem>
                        <PaginationItem><PaginationNext href="#" className={safePage >= totalPages ? "pointer-events-none opacity-40" : ""} onClick={(event) => { event.preventDefault(); if (safePage < totalPages) setPage(safePage + 1); }} /></PaginationItem>
                      </PaginationContent>
                    </Pagination>
                  )}
                  <span>Page {safePage.toLocaleString()} of {totalPages.toLocaleString()}</span>
                </div>
              </>
            ) : (
              <div className="empty-state"><FileSpreadsheet /><p>{eligibleRows.length ? "No contacts match the current filters." : "Upload a file or add a person manually to begin."}</p></div>
            )}
          </section>

          <Dialog open={Boolean(npiDialogRow)} onOpenChange={(open) => { if (!open) { setNpiDialogRow(null); setManualNpiError(""); } }}>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Add or edit NPI</DialogTitle>
                <DialogDescription>Enter an NPI for {npiDialogRow?.__displayName}. Manual entries remain flagged for review and are not treated as CMS-verified.</DialogDescription>
              </DialogHeader>
              <label className="form-field"><span>NPI number</span><Input inputMode="numeric" maxLength={10} value={manualNpi} onChange={(event) => { setManualNpi(event.target.value.replace(/\D/g, "")); setManualNpiError(""); }} placeholder="10-digit NPI" /></label>
              {manualNpiError && <p className="form-error">{manualNpiError}</p>}
              <DialogFooter showCloseButton>
                <Button onClick={saveManualNpi}>Save manual NPI</Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>

          <Dialog open={Boolean(reviewDialogRow)} onOpenChange={(open) => { if (!open) setReviewDialogRow(null); }}>
            <DialogContent className="sm:max-w-2xl">
              <DialogHeader>
                <DialogTitle>Review {reviewDialogRow?.__displayName}</DialogTitle>
                <DialogDescription>Record three separate decisions. An NPI alone is not proof that someone is a doctor.</DialogDescription>
              </DialogHeader>
              <div className="review-summary">
                <div><span>Matched NPI</span><strong>{String(reviewDialogRow?.["NPI Number"] || reviewDialogRow?.["Suggested NPI"] || "Not found")}</strong></div>
                <div><span>Taxonomy</span><strong>{String(reviewDialogRow?.["NPI Primary Specialty"] || "Not available")}</strong></div>
                <div><span>Credential</span><strong>{String(reviewDialogRow?.["NPI Credential"] || "Not available")}</strong></div>
              </div>
              <div className="review-form-grid">
                <label className="form-field">
                  <span>1. Country of professional practice</span>
                  <Select value={reviewDraft.country} onValueChange={(value) => setReviewDraft({ ...reviewDraft, country: value as CountryDecision })}>
                    <SelectTrigger className="bg-white"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {Object.entries(COUNTRY_LABELS).map(([value, label]) => <SelectItem key={value} value={value}>{label}</SelectItem>)}
                    </SelectContent>
                  </Select>
                  <small>{reviewDialogRow?.__countryReason}</small>
                </label>
                <label className="form-field">
                  <span>2. Doctor confirmation</span>
                  <Select value={reviewDraft.doctor} onValueChange={(value) => setReviewDraft({ ...reviewDraft, doctor: value as DoctorDecision })}>
                    <SelectTrigger className="bg-white"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {Object.entries(DOCTOR_LABELS).map(([value, label]) => <SelectItem key={value} value={value}>{label}</SelectItem>)}
                    </SelectContent>
                  </Select>
                  <small>{reviewDialogRow?.__doctorReason}</small>
                </label>
                <label className="form-field form-field-wide">
                  <span>3. Specialty decision</span>
                  <Select value={reviewDraft.specialty} onValueChange={(value) => setReviewDraft({ ...reviewDraft, specialty: value as SpecialtyDecision })}>
                    <SelectTrigger className="bg-white"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {Object.entries(SPECIALTY_LABELS).map(([value, label]) => <SelectItem key={value} value={value}>{label}</SelectItem>)}
                    </SelectContent>
                  </Select>
                  <small>Claimed: {String(reviewDialogRow?.["Claimed Specialty"] || (reviewDialogRow ? textValue(reviewDialogRow, fields.claimedSpecialty) || textValue(reviewDialogRow, fields.categoryDetail) : "") || "—")} · Registered: {String(reviewDialogRow?.["Registered NPI Specialty"] || reviewDialogRow?.["NPI Primary Specialty"] || "—")}</small>
                </label>
                <label className="form-field form-field-wide"><span>Evidence URL</span><Input type="url" value={reviewDraft.evidenceUrl} onChange={(event) => setReviewDraft({ ...reviewDraft, evidenceUrl: event.target.value })} placeholder="Institutional profile or other source" /></label>
                <label className="form-field form-field-wide"><span>Reviewer notes</span><Textarea value={reviewDraft.notes} onChange={(event) => setReviewDraft({ ...reviewDraft, notes: event.target.value })} placeholder="Explain any override, identity clue, or specialty evidence." /></label>
              </div>
              <DialogFooter showCloseButton>
                <Button onClick={saveReview}><Check /> Save review</Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>

          <Dialog open={addDialogOpen} onOpenChange={(open) => { setAddDialogOpen(open); if (!open) setNewContactError(""); }}>
            <DialogContent className="sm:max-w-2xl">
              <DialogHeader>
                <DialogTitle>Add a person</DialogTitle>
                <DialogDescription>Add a licensed contact to the working queue. Leave country or state blank when it still needs research.</DialogDescription>
              </DialogHeader>
              <div className="contact-form-grid">
                <label className="form-field"><span>First name *</span><Input value={newContact.firstName} onChange={(event) => setNewContact({ ...newContact, firstName: event.target.value })} /></label>
                <label className="form-field"><span>Last name *</span><Input value={newContact.lastName} onChange={(event) => setNewContact({ ...newContact, lastName: event.target.value })} /></label>
                <label className="form-field"><span>Email</span><Input type="email" value={newContact.email} onChange={(event) => setNewContact({ ...newContact, email: event.target.value })} /></label>
                <label className="form-field"><span>Organization</span><Input value={newContact.company} onChange={(event) => setNewContact({ ...newContact, company: event.target.value })} /></label>
                <label className="form-field"><span>Country</span><Input value={newContact.country} onChange={(event) => setNewContact({ ...newContact, country: event.target.value })} placeholder="USA or leave blank for research" /></label>
                <label className="form-field"><span>U.S. state</span><Input maxLength={2} value={newContact.state} onChange={(event) => setNewContact({ ...newContact, state: event.target.value.toUpperCase() })} placeholder="NY or leave blank" /></label>
                <label className="form-field"><span>License key *</span><Input value={newContact.license} onChange={(event) => setNewContact({ ...newContact, license: event.target.value })} /></label>
                <label className="form-field"><span>Claimed specialty</span><Input value={newContact.specialty} onChange={(event) => setNewContact({ ...newContact, specialty: event.target.value })} placeholder="Medical Oncology" /></label>
                <label className="form-field"><span>NPI, if known</span><Input inputMode="numeric" maxLength={10} value={newContact.npi} onChange={(event) => setNewContact({ ...newContact, npi: event.target.value.replace(/\D/g, "") })} /></label>
                <label className="form-field form-field-wide"><span>LinkedIn or institutional profile</span><Input type="url" value={newContact.linkedIn} onChange={(event) => setNewContact({ ...newContact, linkedIn: event.target.value })} placeholder="https://…" /></label>
              </div>
              {newContactError && <p className="form-error">{newContactError}</p>}
              <DialogFooter showCloseButton>
                <Button onClick={addContact}><Plus /> Add to queue</Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>

          <section className="method-grid">
            <div><span>1</span><h4>Scope</h4><p>Include licensed U.S. contacts and blank-country contacts that require research.</p></div>
            <div><span>2</span><h4>Country</h4><p>Use the source value or a strong NPPES practice-address match. Ambiguous results remain reviewable.</p></div>
            <div><span>3</span><h4>NPI + doctor</h4><p>Match the identity, then confirm an active individual NPI in the physician taxonomy family, whose codes begin with 20.</p></div>
            <div><span>4</span><h4>Specialty</h4><p>Compare the claimed focus with NPPES taxonomy and review institutional or LinkedIn evidence.</p></div>
            <div><span>5</span><h4>Export</h4><p>Download one cleaned CSV containing the included contacts and their useful enrichment fields.</p></div>
          </section>
        </section>
      </div>
    </main>
  );
}

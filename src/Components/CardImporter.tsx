import  { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  Upload,
  FileSpreadsheet,
  Loader2,
  AlertTriangle,
  RotateCcw,
  ChevronLeft,
  ChevronRight,
  Package,
  Download,
  UserPlus,
} from "lucide-react";

// ---- design tokens ----
const ink = "#14181B";
const inkSoft = "#8A9089";
const paper = "#F1E8D6";
const paperLine = "#C9BB98";
const paperDim = "#E4D8BE";
const brass = "#B8862B";
const brassBright = "#D4A445";
const rust = "#9A4B3E";

// ---- types ----
type ContactField = "firstName" | "lastName" | "workEmail" | "personalEmail" | "location" | "phone";

interface Contact {
  id: number;
  firstName?: string;
  lastName?: string;
  workEmail?: string;
  personalEmail?: string;
  location?: string;
  phone?: string;
  valid: boolean;
}

type Status = "idle" | "parsing" | "parsed" | "error" | "zipping";

const HEADER_MAP: Record<string, ContactField> = {
  firstname: "firstName",
  first: "firstName",
  lastname: "lastName",
  last: "lastName",
  workemail: "workEmail",
  officeemail: "workEmail",
  businessemail: "workEmail",
  personalemail: "personalEmail",
  privateemail: "personalEmail",
  homeemail: "personalEmail",
  email: "personalEmail",
  location: "location",
  city: "location",
  address: "location",
  number: "phone",
  phone: "phone",
  phonenumber: "phone",
  mobile: "phone",
  contactnumber: "phone",
  cell: "phone",
};

function useGoogleFonts(): void {
  useEffect(() => {
    const id = "cci-fonts";
    if (document.getElementById(id)) return;
    const link = document.createElement("link");
    link.id = id;
    link.rel = "stylesheet";
    link.href =
      "https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,600;1,9..144,500&family=IBM+Plex+Mono:wght@400;500;600&display=swap";
    document.head.appendChild(link);
  }, []);
}

function normalizeHeader(h: unknown): string {
  return String(h ?? "").toLowerCase().replace(/[^a-z]/g, "");
}

function mapRow(row: Record<string, unknown>): Omit<Contact, "id" | "valid"> {
  const out: Omit<Contact, "id" | "valid"> = {};
  for (const [rawKey, val] of Object.entries(row)) {
    const field = HEADER_MAP[normalizeHeader(rawKey)];
    if (!field) continue;
    const str = String(val ?? "").trim();
    if (!str) continue;
    // don't let a generic "email" column overwrite one we already matched more specifically
    if (field === "personalEmail" && out.personalEmail) continue;
    out[field] = str;
  }
  return out;
}

function parseWorkbook(arrayBuffer: ArrayBuffer): { sheetName: string; contacts: Contact[] } {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
  const contacts: Contact[] = rows.map((r, i) => {
    const mapped = mapRow(r);
    return { id: i, ...mapped, valid: Boolean(mapped.firstName && mapped.lastName) };
  });
  return { sheetName, contacts };
}

function initials(c: Contact): string {
  const a = (c.firstName ?? "?")[0] ?? "?";
  const b = (c.lastName ?? "")[0] ?? "";
  return (a + b).toUpperCase();
}

// escape characters vCard reserves (comma, semicolon, backslash, newline)
function vcardEscape(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/,/g, "\\,").replace(/;/g, "\\;").replace(/\n/g, "\\n");
}

function buildVCard(c: Contact): string {
  const first = vcardEscape(c.firstName ?? "");
  const last = vcardEscape(c.lastName ?? "");
  const lines = [
    "BEGIN:VCARD",
    "VERSION:3.0",
    `N:${last};${first};;;`,
    `FN:${[first, last].filter(Boolean).join(" ")}`,
  ];
  if (c.workEmail) lines.push(`EMAIL;TYPE=WORK:${vcardEscape(c.workEmail)}`);
  if (c.personalEmail) lines.push(`EMAIL;TYPE=HOME:${vcardEscape(c.personalEmail)}`);
  if (c.phone) lines.push(`TEL;TYPE=CELL:${vcardEscape(c.phone)}`);
  if (c.location) lines.push(`ADR;TYPE=HOME:;;${vcardEscape(c.location)};;;;`);
  lines.push("END:VCARD", "");
  return lines.join("\r\n");
}

function safeFileName(c: Contact): string {
  const base = `${c.firstName ?? "contact"}-${c.lastName ?? c.id}`;
  return base.replace(/[^a-z0-9-_]+/gi, "_");
}

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function downloadSingleContact(c: Contact): void {
  const blob = new Blob([buildVCard(c)], { type: "text/vcard;charset=utf-8" });
  triggerDownload(blob, `${safeFileName(c)}.vcf`);
}

// Share the vCard to the native OS chooser when file sharing is available.
// Browsers without it fall back to opening the vCard directly.
async function addToDeviceContacts(c: Contact): Promise<void> {
  const vcard = buildVCard(c);
  const filename = `${safeFileName(c)}.vcf`;
  const file = new File([vcard], filename, { type: "text/vcard;charset=utf-8" });

  const isAppleMobile = /iPhone|iPad|iPod/.test(navigator.userAgent);
  if (!isAppleMobile && navigator.canShare?.({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: `${c.firstName} ${c.lastName}` });
      return;
    } catch (error) {
      if (error instanceof DOMException && error.name === "AbortError") return;
    }
  }

  const blob = new Blob([vcard], { type: "text/vcard;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const win = window.open(url, "_blank");
  if (!win) {
    // popup blocked — fall back to navigating the current tab
    window.location.href = url;
  }
  // give the browser/OS time to hand the blob off before we revoke it
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

async function downloadAllAsZip(contacts: Contact[]): Promise<void> {
  const zip = new JSZip();
  const used = new Set<string>();
  contacts
    .filter((c) => c.valid)
    .forEach((c) => {
      let name = `${safeFileName(c)}.vcf`;
      let n = 2;
      while (used.has(name)) {
        name = `${safeFileName(c)}-${n}.vcf`;
        n += 1;
      }
      used.add(name);
      zip.file(name, buildVCard(c));
    });
  const blob = await zip.generateAsync({ type: "blob" });
  triggerDownload(blob, "contacts.zip");
}

export default function ContactCardImporter() {
  useGoogleFonts();

  const [status, setStatus] = useState<Status>("idle");
  const [fileName, setFileName] = useState<string>("");
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [sheetName, setSheetName] = useState<string>("");
  const [dragOver, setDragOver] = useState<boolean>(false);
  const [parseError, setParseError] = useState<string>("");
  const inputRef = useRef<HTMLInputElement>(null);
  const scrollerRef = useRef<HTMLDivElement>(null);

  const validCount = contacts.filter((c) => c.valid).length;
  const invalidCount = contacts.length - validCount;

  const handleFile = useCallback((file: File | undefined | null): void => {
    if (!file) return;
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      setParseError("That doesn't look like a spreadsheet. Use .xlsx, .xls, or .csv.");
      setStatus("error");
      return;
    }
    setFileName(file.name);
    setStatus("parsing");
    setParseError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        const { sheetName, contacts } = parseWorkbook(buffer);
        if (contacts.length === 0) {
          setParseError("This sheet came back empty. Check that row 1 has column headers.");
          setStatus("error");
          return;
        }
        setSheetName(sheetName);
        setContacts(contacts);
        setStatus("parsed");
      } catch {
        setParseError("Couldn't read that file. It may be corrupted or password-protected.");
        setStatus("error");
      }
    };
    reader.onerror = () => {
      setParseError("Couldn't read that file from disk. Try again.");
      setStatus("error");
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const reset = (): void => {
    setStatus("idle");
    setFileName("");
    setContacts([]);
    setParseError("");
    if (inputRef.current) inputRef.current.value = "";
  };

  const handleDownloadAll = async (): Promise<void> => {
    setStatus("zipping");
    try {
      await downloadAllAsZip(contacts);
    } finally {
      setStatus("parsed");
    }
  };

  const scrollBy = (dir: number): void => {
    const el = scrollerRef.current;
    if (!el) return;
    const first = el.firstElementChild as HTMLElement | null;
    const cardWidth = first ? first.offsetWidth + 16 : 280;
    el.scrollBy({ left: dir * cardWidth, behavior: "smooth" });
  };

  return (
    <div
      className="min-h-screen w-full flex flex-col items-center px-4 py-10 sm:py-14"
      style={{ background: ink, fontFamily: "'IBM Plex Mono', monospace" }}
    >
      <div className="w-full max-w-xl">
        {/* header */}
        <div className="mb-8 sm:mb-10">
          <div
            className="inline-block px-3 py-1 text-xs tracking-wide mb-5 rotate-[-1deg]"
            style={{ background: brass, color: ink }}
          >
            import / contacts.xlsx → cards.zip
          </div>
          <h1
            className="text-3xl sm:text-4xl leading-tight"
            style={{ color: paper, fontFamily: "'Fraunces', serif", fontWeight: 600 }}
          >
            Turn a spreadsheet into a stack of{" "}
            <span style={{ fontStyle: "italic", fontWeight: 500, color: brassBright }}>shareable</span>{" "}
            contact cards.
          </h1>
          <p className="mt-3 text-sm leading-relaxed" style={{ color: inkSoft }}>
            Drop in a list of people. Everything happens on this device — we build one vCard per
            row and let you download it alone or as one zip.
          </p>
        </div>

        {/* upload zone */}
        {status !== "parsed" && status !== "zipping" && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragOver(true);
            }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragOver(false);
              handleFile(e.dataTransfer.files?.[0]);
            }}
            className="relative"
          >
            <div className="absolute -top-3 left-6 px-2 py-0.5 text-xs" style={{ background: paperDim, color: ink }}>
              spreadsheet in
            </div>
            <label
              htmlFor="file-input"
              className="flex flex-col items-center justify-center text-center cursor-pointer px-6 py-14 sm:py-16 transition-transform"
              style={{
                background: paper,
                border: `2px dashed ${dragOver ? brass : paperLine}`,
                transform: dragOver ? "scale(1.01) rotate(0deg)" : "rotate(-0.3deg)",
              }}
            >
              <div className="w-4 h-4 rounded-full mb-6" style={{ background: ink, boxShadow: `0 0 0 3px ${brass}` }} />
              {status === "parsing" ? (
                <Loader2 className="w-7 h-7 animate-spin mb-3" style={{ color: ink }} />
              ) : (
                <Upload className="w-7 h-7 mb-3" style={{ color: ink }} />
              )}
              <p style={{ color: ink, fontFamily: "'Fraunces', serif", fontSize: "1.05rem" }}>
                {status === "parsing" ? "Reading your file…" : "Tap to choose, or drop your file here"}
              </p>
              <p className="mt-2 text-xs" style={{ color: "#6B5F45" }}>
                .xlsx · .xls · .csv
              </p>
              <input
                id="file-input"
                ref={inputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => handleFile(e.target.files?.[0])}
              />
            </label>
          </div>
        )}

        {status === "error" && (
          <div
            className="mt-4 flex items-start gap-3 px-4 py-3 text-sm"
            style={{ background: "#241A17", border: `1px solid ${rust}`, color: "#E8C9C1" }}
          >
            <AlertTriangle className="w-4 h-4 mt-0.5 shrink-0" style={{ color: rust }} />
            <div className="flex-1">{parseError}</div>
            <button onClick={reset} className="text-xs underline shrink-0" style={{ color: "#E8C9C1" }}>
              try again
            </button>
          </div>
        )}

        {(status === "parsed" || status === "zipping") && (
          <div>
            <div className="flex items-center justify-between px-4 py-3 mb-5" style={{ background: paperDim, color: ink }}>
              <div className="flex items-center gap-2 text-sm min-w-0">
                <FileSpreadsheet className="w-4 h-4 shrink-0" />
                <span className="truncate">{fileName}</span>
              </div>
              <button onClick={reset} className="flex items-center gap-1 text-xs shrink-0 ml-3" style={{ color: "#6B5F45" }}>
                <RotateCcw className="w-3 h-3" /> replace
              </button>
            </div>

            <div className="flex items-center justify-between mb-3 text-xs" style={{ color: inkSoft }}>
              <span>
                {validCount} contact{validCount === 1 ? "" : "s"} ready
                {invalidCount > 0 ? ` · ${invalidCount} skipped (missing name)` : ""}
              </span>
              <span>{sheetName}</span>
            </div>

            <p className="mb-3 text-xs leading-relaxed" style={{ color: "#6B5F45" }}>
              "Add to contacts" opens the contact import screen on iPhone/iPad. On supported
              Android browsers it opens the share sheet; choose Contacts to import. Other browsers
              open or download the .vcf file.
            </p>

            <div className="relative">
              <div ref={scrollerRef} className="flex gap-4 overflow-x-auto pb-3 snap-x snap-mandatory" style={{ scrollbarWidth: "none" }}>
                {contacts.map((c) => (
                  <div
                    key={c.id}
                    className="snap-start shrink-0 w-64 sm:w-72 px-5 py-5 relative"
                    style={{ background: c.valid ? paper : "#3A2E28", border: `1px solid ${c.valid ? paperLine : rust}` }}
                  >
                    <div
                      className="absolute top-3 right-3 w-3 h-3 rounded-full"
                      style={{ background: ink, boxShadow: `0 0 0 2px ${c.valid ? paperDim : "#5A4038"}` }}
                    />
                    <div
                      className="w-10 h-10 flex items-center justify-center text-sm mb-4"
                      style={{ background: c.valid ? brass : rust, color: paper, fontFamily: "'Fraunces', serif" }}
                    >
                      {initials(c)}
                    </div>
                    <p className="text-lg leading-snug mb-3" style={{ color: c.valid ? ink : "#E8C9C1", fontFamily: "'Fraunces', serif" }}>
                      {c.firstName || "—"} {c.lastName || ""}
                      {!c.valid && (
                        <span className="block text-xs mt-1" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                          missing first/last name
                        </span>
                      )}
                    </p>
                    <dl className="space-y-1.5 text-xs mb-4" style={{ color: c.valid ? "#5C5138" : "#C9A99E" }}>
                      {c.workEmail && (
                        <div className="flex gap-2">
                          <dt className="shrink-0 w-14 opacity-70">work</dt>
                          <dd className="truncate">{c.workEmail}</dd>
                        </div>
                      )}
                      {c.personalEmail && (
                        <div className="flex gap-2">
                          <dt className="shrink-0 w-14 opacity-70">personal</dt>
                          <dd className="truncate">{c.personalEmail}</dd>
                        </div>
                      )}
                      {c.phone && (
                        <div className="flex gap-2">
                          <dt className="shrink-0 w-14 opacity-70">phone</dt>
                          <dd className="truncate">{c.phone}</dd>
                        </div>
                      )}
                      {c.location && (
                        <div className="flex gap-2">
                          <dt className="shrink-0 w-14 opacity-70">location</dt>
                          <dd className="truncate">{c.location}</dd>
                        </div>
                      )}
                      {!c.workEmail && !c.personalEmail && !c.phone && !c.location && (
                        <div className="opacity-60">no other fields found</div>
                      )}
                    </dl>
                    {c.valid && (
                      <div className="flex items-center gap-3">
                        <button
                          onClick={() => addToDeviceContacts(c)}
                          className="flex items-center gap-1.5 text-xs px-3 py-1.5"
                          style={{ background: ink, color: paper }}
                        >
                          <UserPlus className="w-3 h-3" /> add to contacts
                        </button>
                        <button
                          onClick={() => downloadSingleContact(c)}
                          className="flex items-center gap-1 text-xs underline"
                          style={{ color: "#6B5F45" }}
                        >
                          <Download className="w-3 h-3" /> .vcf
                        </button>
                      </div>
                    )}
                  </div>
                ))}
              </div>

              {contacts.length > 1 && (
                <div className="hidden sm:flex items-center justify-between mt-1">
                  <button onClick={() => scrollBy(-1)} className="w-8 h-8 flex items-center justify-center" style={{ background: paperDim, color: ink }} aria-label="Previous card">
                    <ChevronLeft className="w-4 h-4" />
                  </button>
                  <span className="text-xs" style={{ color: inkSoft }}>
                    swipe or use arrows
                  </span>
                  <button onClick={() => scrollBy(1)} className="w-8 h-8 flex items-center justify-center" style={{ background: paperDim, color: ink }} aria-label="Next card">
                    <ChevronRight className="w-4 h-4" />
                  </button>
                </div>
              )}
            </div>

            <div className="mt-6">
              <button
                onClick={handleDownloadAll}
                disabled={validCount === 0 || status === "zipping"}
                className="w-full flex items-center justify-center gap-2 py-4 text-sm transition-opacity"
                style={{
                  background: validCount === 0 ? "#5C5138" : brass,
                  color: ink,
                  opacity: status === "zipping" ? 0.8 : 1,
                  cursor: validCount === 0 ? "not-allowed" : "pointer",
                }}
              >
                {status === "zipping" ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" /> building your zip…
                  </>
                ) : (
                  <>
                    <Package className="w-4 h-4" /> download all {validCount} as .zip
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        <p className="mt-10 text-xs text-center" style={{ color: "#4A5049" }}>
          first name · last name · work email · personal email · location · number
        </p>
      </div>
    </div>
  );
}
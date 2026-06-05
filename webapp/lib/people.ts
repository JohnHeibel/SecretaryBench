// A curated roster of standard people so the same roles are spelled the same way across every
// storyline. The cast is still stored per-node (the corpus model is unchanged) — picking from the
// roster just materializes a standard key + readable name into THIS node's cast. From/To/Cc are
// inbox presentation only and never graded, so this is purely an authoring-consistency aid.

export interface RosterPerson { key: string; name: string; }

// Keys are in the cast-key format (UPPER_SNAKE). Order is the order shown in the picker.
export const STANDARD_ROSTER: RosterPerson[] = [
  { key: "CEO", name: "you (the boss)" },
  { key: "EA", name: "Executive Assistant" },
  { key: "CHIEF_OF_STAFF", name: "Chief of Staff" },
  { key: "CFO", name: "Chief Financial Officer" },
  { key: "COO", name: "Chief Operating Officer" },
  { key: "CTO", name: "Chief Technology Officer" },
  { key: "GC", name: "General Counsel" },
  { key: "BOARD_CHAIR", name: "Board Chair" },
  { key: "VP_SALES", name: "VP Sales" },
  { key: "VP_ENG", name: "VP Engineering" },
  { key: "VP_PRODUCT", name: "VP Product" },
  { key: "VP_MARKETING", name: "VP Marketing" },
  { key: "HR", name: "Head of HR" },
  { key: "IR", name: "Investor Relations" },
  { key: "CLIENT", name: "Client Contact" },
  { key: "VENDOR", name: "Vendor Contact" },
  { key: "PARTNER", name: "Partner Contact" },
  { key: "RECRUITER", name: "Recruiter" },
  // Extended roles used in the Project Atlas example and common in other storylines
  { key: "CORPDEV", name: "Corporate Development" },
  { key: "BOARD", name: "Board office" },
  { key: "DESIGN", name: "Design Lead" },
  { key: "BIZDEV", name: "Business Development" },
  { key: "LEGAL", name: "Legal" },
  { key: "FINANCE", name: "Finance" },
  { key: "PEOPLE", name: "People Ops" },
  { key: "OPS", name: "Operations" },
  { key: "COMMS", name: "Communications / PR" },
];

export const MAX_PERSON_KEY = 24;     // cast key (the identifier referenced by From/To/Cc) length cap
export const MAX_PERSON_NAME = 40;    // display-name length cap (room for "Name (Role)")

// Normalize a cast key to the standard format: UPPER_SNAKE, no spaces/punctuation, length-capped.
// So a key is always a clean identifier and the same person reads identically everywhere.
export function normalizeKey(raw: string): string {
  return raw.toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, MAX_PERSON_KEY);
}

// Collapse runs of whitespace, trim, and cap length, so the same person isn't stored five
// slightly-different ways ("Sarah  Chen ", "sarah chen", …).
export function normalizeName(raw: string): string {
  return raw.replace(/\s+/g, " ").trim().slice(0, MAX_PERSON_NAME);
}

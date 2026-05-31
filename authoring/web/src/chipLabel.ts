import type { Chip, ChipBase } from "./types";

export const WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
export const WEEKDAY_LONG: Record<string, string> = {
  MON: "Monday", TUE: "Tuesday", WED: "Wednesday", THU: "Thursday",
  FRI: "Friday", SAT: "Saturday", SUN: "Sunday",
};
export const UNITS = [
  { value: "days", label: "days" },
  { value: "business-days", label: "business days" },
  { value: "weeks", label: "weeks" },
  { value: "months", label: "months" },
  { value: "years", label: "years" },
];
export const ORDINALS: (number | "last")[] = [1, 2, 3, 4, 5, "last"];

export const BASE_LABELS: Record<string, string> = {
  serve: "the day this email arrives",
  anchor: "a saved date",
  next_weekday: "next weekday",
  this_weekday: "this week's weekday",
  nth_weekday: "Nth weekday of a month",
  day_of_month: "a day of the month",
  month: "a whole month",
  week_of: "the week of…",
  raw: "advanced (raw token)",
};

function ord(n: number | "last" | undefined): string {
  if (n === "last") return "last";
  const map: Record<number, string> = { 1: "1st", 2: "2nd", 3: "3rd", 4: "4th", 5: "5th" };
  return map[n ?? 1] || `${n}th`;
}

function monthRef(off: number | undefined): string {
  const o = off ?? 0;
  if (o === 0) return "this month";
  if (o === 1) return "next month";
  if (o === -1) return "last month";
  return o > 0 ? `in ${o} months` : `${-o} months ago`;
}

function timeLabel(t?: { hour: number; minute: number } | null): string {
  if (!t) return "";
  const h = t.hour % 12 || 12;
  const ap = t.hour < 12 ? "AM" : "PM";
  const mm = t.minute === 0 ? "" : `:${String(t.minute).padStart(2, "0")}`;
  return ` · ${h}${mm} ${ap}`;
}

function baseLabel(b: ChipBase): string {
  switch (b.kind) {
    case "serve": return "when it arrives";
    case "anchor": return `📌 ${b.name || "?"}`;
    case "next_weekday": return `next ${WEEKDAY_LONG[b.weekday || "MON"]}`;
    case "this_weekday": return `this ${WEEKDAY_LONG[b.weekday || "MON"]}`;
    case "nth_weekday": return `${ord(b.n)} ${WEEKDAY_LONG[b.weekday || "MON"]} of ${monthRef(b.month_offset)}`;
    case "day_of_month": return `the ${ord(b.day || 1)} of ${monthRef(b.month_offset)}`;
    case "month": return monthRef(b.month_offset);
    case "week_of": return `week of ${b.inner ? baseLabel(b.inner.base) : "?"}`;
    case "raw": return b.token || "raw";
  }
}

export function chipLabel(chip: Chip): string {
  let s = baseLabel(chip.base);
  if (chip.offset && chip.base.kind !== "raw") {
    const u = UNITS.find((x) => x.value === chip.offset!.unit)?.label || chip.offset.unit;
    const a = chip.offset.amount;
    s = a >= 0 ? `${a} ${u} after ${s}` : `${-a} ${u} before ${s}`;
  }
  s += timeLabel(chip.time);
  return s;
}

export function emptyChip(kind: ChipBase["kind"] = "next_weekday"): Chip {
  const base: ChipBase = { kind };
  if (kind === "next_weekday" || kind === "this_weekday") base.weekday = "THU";
  if (kind === "nth_weekday") { base.n = 1; base.weekday = "MON"; base.month_offset = 0; }
  if (kind === "day_of_month") { base.day = 1; base.month_offset = 0; }
  if (kind === "month") base.month_offset = 0;
  if (kind === "anchor") base.name = "";
  if (kind === "week_of") base.inner = emptyChip("next_weekday");
  if (kind === "raw") base.token = "serve";
  return { base, offset: null, time: null };
}

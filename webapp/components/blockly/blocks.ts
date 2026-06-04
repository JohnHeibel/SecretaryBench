// Custom Blockly blocks for the SecretaryBench date grammar, plus a manual
// block -> grammar-string compiler. We traverse the blocks ourselves (rather than
// using Blockly's generator framework) so the output is version-proof and obviously
// correct against ANSWER_KEY_GRAMMAR.md §3. The string is then handed to the REAL
// Python resolver for the truth — these blocks only make valid strings easy to build.
/* eslint-disable @typescript-eslint/no-explicit-any */

import { WEEKDAYS, NTH, UNITS } from "@/lib/grammar";

let currentAnchors: string[] = [];
export function setAnchors(list: string[]) {
  currentAnchors = list;
}

const EXPR = "Expr";
const MONTHREF = "Monthref";

export function defineGrammarBlocks(Blockly: any) {
  const wd = WEEKDAYS.map((d) => [d.toLowerCase(), d]);
  const units = UNITS.map((u) => [u.label, u.value]);
  const nth = NTH.map((n) => [n, n]);

  Blockly.defineBlocksWithJsonArray([
    { type: "tk_serve", message0: "serve (today)", output: EXPR, colour: 210, tooltip: "The day this email is delivered." },
    {
      type: "tk_anchor", message0: "anchor @ %1", output: EXPR, colour: 285,
      args0: [{ type: "field_dropdown", name: "ANCHOR", options: () => (currentAnchors.length ? currentAnchors.map((a) => [a, a]) : [["(no anchors yet)", ""]]) }],
      tooltip: "A date another email published with {!name = ...}.",
    },
    {
      type: "tk_offset", message0: "%1 %2 %3 %4", output: EXPR, colour: 160,
      args0: [
        { type: "input_value", name: "EXPR", check: EXPR },
        { type: "field_dropdown", name: "SIGN", options: [["+ plus", "+"], ["− minus", "-"]] },
        { type: "field_number", name: "N", value: 1, min: 0, precision: 1 },
        { type: "field_dropdown", name: "UNIT", options: units },
      ],
      inputsInline: true, tooltip: "Shift a date by N days / business days / weeks / months / years.",
    },
    {
      type: "tk_next", message0: "next %1 %2", output: EXPR, colour: 30,
      args0: [{ type: "field_dropdown", name: "WD", options: wd }, { type: "input_value", name: "FROM", check: EXPR }],
      inputsInline: true, tooltip: "The next weekday strictly after serve, or after the optional 'from' date.",
    },
    {
      type: "tk_this", message0: "this %1 %2", output: EXPR, colour: 30,
      args0: [{ type: "field_dropdown", name: "WD", options: wd }, { type: "input_value", name: "FROM", check: EXPR }],
      inputsInline: true, tooltip: "The weekday in serve's (or the 'from' date's) Mon–Sun week.",
    },
    {
      type: "tk_nth", message0: "the %1 %2 of %3", output: EXPR, colour: 45,
      args0: [{ type: "field_dropdown", name: "N", options: nth }, { type: "field_dropdown", name: "WD", options: wd }, { type: "input_value", name: "MONTH", check: MONTHREF }],
      inputsInline: true, tooltip: "Nth weekday of a month, e.g. the 3rd Friday of next month.",
    },
    {
      type: "tk_dom", message0: "day %1 of %2", output: EXPR, colour: 45,
      args0: [{ type: "field_number", name: "D", value: 1, min: 1, max: 31, precision: 1 }, { type: "input_value", name: "MONTH", check: MONTHREF }],
      inputsInline: true, tooltip: "The Dth day-of-month of a month.",
    },
    {
      type: "tk_week_of", message0: "the week of %1", output: EXPR, colour: 120,
      args0: [{ type: "input_value", name: "EXPR", check: EXPR }],
      inputsInline: true, tooltip: "An INTERVAL: the Mon–Sun week containing a date.",
    },
    {
      type: "tk_month", message0: "the whole month %1", output: EXPR, colour: 120,
      args0: [{ type: "input_value", name: "MONTH", check: MONTHREF }],
      inputsInline: true, tooltip: "An INTERVAL spanning a whole month.",
    },
    {
      type: "tk_monthref", message0: "%1 %2", output: MONTHREF, colour: 260,
      args0: [{ type: "field_dropdown", name: "SIGN", options: [["this month", "0"], ["months ahead +", "+"], ["months back −", "-"]] }, { type: "field_number", name: "N", value: 1, min: 0, precision: 1 }],
      inputsInline: true, tooltip: "Which month, relative to serve's month.",
    },
  ]);
}

// --- block -> grammar string ----------------------------------------------

function child(block: any, name: string): any {
  const input = block.getInput(name);
  return input?.connection?.targetBlock() ?? null;
}

// Returns the grammar string, or null if a required slot is empty (incomplete build).
export function blockToExpr(block: any): string | null {
  if (!block) return null;
  switch (block.type) {
    case "tk_serve":
      return "serve";
    case "tk_anchor": {
      const a = block.getFieldValue("ANCHOR");
      return a ? `@${a}` : null;
    }
    case "tk_offset": {
      const base = blockToExpr(child(block, "EXPR"));
      if (base === null) return null;
      return `${base}${block.getFieldValue("SIGN")}${block.getFieldValue("N")}${block.getFieldValue("UNIT")}`;
    }
    case "tk_next":
    case "tk_this": {
      const kw = block.type === "tk_next" ? "next" : "this";
      const wd = block.getFieldValue("WD");
      const from = child(block, "FROM");
      if (!from) return `${kw}:${wd}`;
      const fe = blockToExpr(from);
      return fe === null ? null : `${kw}:${wd} from ${fe}`;
    }
    case "tk_nth": {
      const m = blockToExpr(child(block, "MONTH"));
      if (m === null) return null;
      return `nth:${block.getFieldValue("N")},${block.getFieldValue("WD")},${m}`;
    }
    case "tk_dom": {
      const m = blockToExpr(child(block, "MONTH"));
      if (m === null) return null;
      return `dom:${block.getFieldValue("D")},${m}`;
    }
    case "tk_week_of": {
      const e = blockToExpr(child(block, "EXPR"));
      return e === null ? null : `week_of:(${e})`;
    }
    case "tk_month": {
      const m = blockToExpr(child(block, "MONTH"));
      return m === null ? null : `month:${m}`;
    }
    case "tk_monthref": {
      const sign = block.getFieldValue("SIGN");
      if (sign === "0") return "0m";
      return `${sign}${block.getFieldValue("N")}m`;
    }
    default:
      return null;
  }
}

// The expression root: a top block that produces an Expr and isn't plugged in.
export function compileWorkspace(workspace: any): string | null {
  const tops = workspace.getTopBlocks(true).filter((b: any) => b.outputConnection && !b.outputConnection.targetBlock());
  const root = tops.find((b: any) => b.type !== "tk_monthref");
  return root ? blockToExpr(root) : null;
}

// Toolbox offered in the block editor flyout.
export const TOOLBOX = {
  kind: "flyoutToolbox",
  contents: [
    { kind: "block", type: "tk_serve" },
    { kind: "block", type: "tk_anchor" },
    { kind: "block", type: "tk_offset" },
    { kind: "block", type: "tk_next" },
    { kind: "block", type: "tk_this" },
    { kind: "block", type: "tk_nth" },
    { kind: "block", type: "tk_dom" },
    { kind: "block", type: "tk_week_of" },
    { kind: "block", type: "tk_month" },
    { kind: "block", type: "tk_monthref" },
  ],
};

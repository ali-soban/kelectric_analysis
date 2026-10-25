// Checks parse.js against ke_update.py on every bill in $KE_BILLS_DIR (default ~/Desktop/ke_bills).
// Run from web/: npm install (once), then node test_parse.mjs
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { execFileSync } from "child_process";
import fs from "fs";
import os from "os";
import path from "path";
import { pdfText, parseBill, BillError } from "./parse.js";

const dir = process.env.KE_BILLS_DIR || path.join(os.homedir(), "Desktop", "ke_bills");
const py = JSON.parse(execFileSync("python", ["-c", `
import json, sys; sys.path.insert(0, ".."); import ke_update as k
out = {}
for p in sorted(k.BILLS.glob("[0-9][0-9][0-9][0-9].pdf")):
    try: out[p.stem] = k.parse_bill(p)
    except k.BillError as e: out[p.stem] = None
print(json.dumps(out, default=str))`], { encoding: "utf8" }));

const PRE_SOLAR = ["2308", "2309"];  // regular bills before net metering
let fails = 0;
const fail = msg => { fails++; console.log("FAIL", msg); };

for (const [bill, want] of Object.entries(py)) {
  const t = await pdfText(pdfjs, new Uint8Array(fs.readFileSync(path.join(dir, bill + ".pdf"))));
  let got;
  try { got = parseBill(t); } catch (e) {
    if (!(e instanceof BillError)) throw e;
    const kind = PRE_SOLAR.includes(bill) ? "regular" : "old";
    if (want || e.kind !== kind) fail(`${bill}: ${e.kind} "${e.message}", python ${want ? "parsed it" : "skipped it"}, expected kind ${kind}`);
    continue;
  }
  if (!want) { fail(`${bill}: parsed but python raised`); continue; }
  for (const [k, v] of Object.entries(want)) {
    const g = got[k];
    const same = typeof v === "number" ? Math.abs(v - g) < 1e-9 : JSON.stringify(v) === JSON.stringify(g);
    if (!same) fail(`${bill}.${k}: js ${JSON.stringify(g)} python ${JSON.stringify(v)}`);
  }
}
console.log(fails ? `${fails} failures` : `OK: ${Object.keys(py).length} bills match ke_update.py`);
process.exit(fails ? 1 : 0);

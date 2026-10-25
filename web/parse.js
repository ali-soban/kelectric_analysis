// Browser port of parse_bill() / charge_breakdown() / derive() from ke_update.py.
// Keep the two in step: web/test_parse.mjs checks this file gives the same results as the Python.

export const MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(" ");

export class BillError extends Error {
  // kind: "other" (not a KE bill), "regular" (no net metering), "old" (pre Aug-25 layout), "bad"
  constructor(message, kind = "bad") { super(message); this.kind = kind; }
}

const num = s => parseFloat(s.replace(/,/g, ""));
const halfUp = x => x >= 0 ? Math.floor(x + 0.5) : -Math.floor(-x + 0.5);
const ceil2 = x => Math.ceil(Math.round(x * 100) / 100);  // math.ceil(round(x, 2))
export const label = ym => `${MONTHS[+ym.slice(2) - 1]}-${ym.slice(0, 2)}`;

/** Page text as lines, grouped like pdfplumber's extract_text() so the same regexes work. */
export async function pdfText(pdfjs, data) {
  const doc = await pdfjs.getDocument({ data, verbosity: 0 }).promise;
  const out = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const items = (await (await doc.getPage(p)).getTextContent()).items
      .filter(i => i.str.trim() && !i.transform[1] && !i.transform[2])  // rotated text = chart labels
      .map(i => ({ s: i.str, x: i.transform[4], top: i.transform[5] + i.height, w: i.width }))
      .sort((a, b) => b.top - a.top || a.x - b.x);
    const lines = [];
    for (const it of items) {
      const line = lines.find(l => Math.abs(l.top - it.top) <= 3);
      line ? line.items.push(it) : lines.push({ top: it.top, items: [it] });
    }
    for (const { items: li } of lines) {
      li.sort((a, b) => a.x - b.x);
      let s = "", end = null;
      for (const it of li) { s += (end != null && it.x - end > 1 ? " " : "") + it.s; end = it.x + it.w; }
      out.push(s.replace(/\s+/g, " ").trim());
    }
  }
  await doc.destroy();
  return out.join("\n");
}

/** '2609' from the header line ': Sep-26 : 07-Sep-2026' (new layout only). */
export function billMonth(t) {
  const m = t.match(/^: (\w{3})-(\d{2}) : \d/m);
  return m && MONTHS.includes(m[1]) ? m[2] + String(MONTHS.indexOf(m[1]) + 1).padStart(2, "0") : null;
}

export const accountNo = t => t.match(/Account No\.? ?(\d{10,})/)?.[1] ?? null;

/** One new-layout (Aug-25 onward) net-metering bill. Throws BillError rather than guessing. */
export function parseBill(t) {
  if (!/ke\.com\.pk|K-Electric|Consumer No/i.test(t)) throw new BillError("not a K-Electric bill", "other");
  if (!/Import Energy|Net Metering Invoice/.test(t)) throw new BillError("regular bill, no net metering on it", "regular");
  const ym = billMonth(t);
  if (!ym || !t.includes("Variable Off Peak")) throw new BillError("older bill layout (before Aug-2025), not supported yet", "old");

  const find = (re, what) => {
    const m = t.match(re);
    if (!m) throw new BillError(`could not find ${what}`);
    return m;
  };
  const meter = name => parseFloat(find(new RegExp(`^${name} [\\d.]+ [\\d.]+ ([\\d.]+)`, "m"), name)[1]);
  const money = (name, re = name) => num(find(new RegExp(`^${re} (-?[\\d,]+(?:\\.\\d+)?)`, "m"), name)[1]);

  const off = meter("Energy - Off Peak"), peak = meter("Energy - Peak");
  const exported = meter("Import Energy Off Peak") + meter("Import Energy Peak");

  // One line per tariff in force; a mid-period change adds a "(New)" line.
  const varOff = [...t.matchAll(/^Variable Off Peak(?: \(New\))? ([\d.]+) ([\d.]+) /gm)];
  const varPeak = [...t.matchAll(/^Variable Peak(?: \(New\))? ([\d.]+) ([\d.]+) /gm)];
  if (!varOff.length || !varPeak.length) throw new BillError("could not find tariff lines");
  if (Math.abs(varOff.reduce((s, m) => s + parseFloat(m[1]), 0) - off) > 0.05 * varOff.length)
    throw new BillError("off-peak billed units do not match the meter reading");

  const billed = money("Current Bill Amount");
  const benefit = money("Total Net Metering Benefit");
  const opening = money("Opening Balance");
  const prev = money("Previous Dues", "Previous (?:Recoverable )?Dues");
  const net = prev + billed - benefit;  // what this bill leaves: + to pay, - credit

  const outstanding = t.match(/^Outstanding Balance (-?[\d,]+\.\d+)/m);
  const printed = t.match(/^Amount Payable within Due Date ([\d,]+)/m);
  let payable, credit;
  if (outstanding) {
    [payable, credit] = [0, halfUp(-num(outstanding[1]))];
    if (Math.abs(num(outstanding[1]) - net) > 1)
      throw new BillError(`outstanding balance ${outstanding[1]} != computed ${net.toFixed(2)}`);
  } else if (printed) {
    [payable, credit] = [Math.trunc(num(printed[1])), 0];
    if (Math.abs(payable - ceil2(net)) > 1)
      throw new BillError(`printed payable ${printed[1]} != computed ${net.toFixed(2)}`);
  } else if (net < -1) {
    [payable, credit] = [0, halfUp(-net)];
  } else if (net <= 1) {
    // Downloaded after paying: the payment already sits in Previous Dues, so
    // rebuild what was due from the credit (if any) carried into this bill.
    [payable, credit] = [ceil2(Math.min(0, opening) + billed - benefit), 0];
  } else {
    [payable, credit] = [ceil2(net), 0];
  }

  const ppp = t.match(/NAPP Off-Peak ([\d.,]+) ([\d.]+) /);
  const head = find(/^Description (\w{3}-\d{2}) (\w{3}-\d{2}) (\w{3}-\d{2})$/m, "carry-forward history").slice(1);
  const tot = find(/^Total ([\d,.]+) ([\d,.]+) ([\d,.]+)$/m, "carry-forward totals");
  if (!head.includes(label(ym))) throw new BillError("carry-forward history does not include this month");
  const cf = ppp ? 0 : num(tot[head.indexOf(label(ym)) + 1]);

  const [d, mon, y] = find(/Reading Date: (\d{2})-(\w{3})-(\d{2})/, "reading date").slice(1);
  return {
    breakdown: chargeBreakdown(t, billed),
    bill: ym, source: "PDF", account: accountNo(t),
    reading: `20${y}-${String(MONTHS.indexOf(mon) + 1).padStart(2, "0")}-${d}`,
    E: exported, F: off, G: peak,
    I: parseFloat(varOff.at(-1)[2]), J: parseFloat(varPeak.at(-1)[2]),
    L: cf, M: payable, N: billed, Q: credit, R: benefit,
    S: ppp ? num(ppp[1]) : null, T: ppp ? parseFloat(ppp[2]) : null,
  };
}

const CHARGE_GROUPS = [["Variable Off Peak", "Off-peak energy"], ["Variable Peak", "Peak energy"],
  ["Fixed Charges", "Fixed charges"], ["FCA", "Fuel adjustment (FCA)"],
  ["Uniform Quarterly", "Quarterly adjustment"], ["Additional Surcharge", "PHL surcharge"],
  ["Electricity Duty", "Electricity duty"], ["Sales Tax", "Sales tax"]];

/** Bill charges grouped by type; null unless they add up to the Current Bill Amount. */
function chargeBreakdown(t, billed) {
  const sec = t.match(/Units Rate Amount\n(.*?)\nElectricity Charges for current month/s);
  if (!sec) return null;
  const groups = {};
  for (const line of sec[1].split("\n")) {
    const amt = line.match(/(-?[\d,]+\.\d{2})$/);
    if (!amt || line.startsWith("Electricity Charges ") || line.startsWith("Taxes and Duties")) continue;  // subtotals
    const name = CHARGE_GROUPS.find(([p]) => line.startsWith(p))?.[1]
      ?? line.slice(0, amt.index).replace(/[\s\d.,:-]+$/, "").trim();
    groups[name] = Math.round(((groups[name] ?? 0) + num(amt[1])) * 100) / 100;
  }
  return Math.abs(Object.values(groups).reduce((a, b) => a + b, 0) - billed) <= 1 ? groups : null;
}

/** Same maths as derive() in ke_update.py; C (solar production) may be null. */
export function derive(row) {
  const { C: c, E: e, F: f, G: g, N: n, M: m } = row;
  const k = f + g ? n / (f + g) : null;
  const d = c != null ? c - e : null;
  const o = d != null && k != null ? n + d * k : null;
  return { ...row, D: d, H: e - f, K: k, O: o, P: o != null ? o - m : null };
}

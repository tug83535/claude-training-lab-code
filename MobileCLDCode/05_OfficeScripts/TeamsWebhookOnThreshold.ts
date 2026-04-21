/**
 * Office Script: Post a Teams notification when any monitored cell crosses its threshold.
 *
 * WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
 * ---------------------------------------
 * Office Scripts run in Excel on the web *and* can be triggered by Power Automate
 * on a schedule. Native Excel conditional formatting can colour a cell but cannot
 * post outbound HTTP to Teams. This lets a finance workbook in SharePoint
 * alert a Teams channel without any desktop VBA.
 *
 * HOW IT WORKS
 * ------------
 *  1. Sheet "Watchers" columns: Name | CellRef | Operator | Threshold | LastAlert
 *  2. Script iterates every row, compares the current value to the threshold,
 *     and posts an Adaptive Card to the Teams webhook URL stored in a named range.
 *  3. Paired with a Power Automate scheduled flow (every 15 minutes on weekdays),
 *     this becomes a continuous monitor.
 *
 * RUN FROM: Excel on the web -> Automate tab -> New Script, paste, save.
 */

async function main(workbook: ExcelScript.Workbook): Promise<void> {
  const watcherSheet = workbook.getWorksheet("Watchers");
  if (!watcherSheet) {
    console.log("No 'Watchers' sheet found.");
    return;
  }

  const webhookUrl = getNamedValue(workbook, "TeamsWebhookUrl");
  if (!webhookUrl) {
    console.log("Named range 'TeamsWebhookUrl' not set.");
    return;
  }

  const used = watcherSheet.getUsedRange();
  const values = used.getValues();            // [ [headers], [row], ... ]
  const headers = values[0] as string[];

  const idx = {
    name:      headers.indexOf("Name"),
    cellRef:   headers.indexOf("CellRef"),
    op:        headers.indexOf("Operator"),
    threshold: headers.indexOf("Threshold"),
    lastAlert: headers.indexOf("LastAlert"),
  };

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const cellRef = row[idx.cellRef] as string;
    const op      = row[idx.op] as string;
    const thresh  = Number(row[idx.threshold]);
    if (!cellRef || !op) continue;

    const current = resolveCell(workbook, cellRef);
    if (current === null) continue;

    const breached = evaluate(current, op, thresh);
    if (!breached) continue;

    const name = row[idx.name] as string;
    const msg = `**${name}** ${op} ${thresh}. Current value: **${current}**`;
    await postAdaptiveCard(webhookUrl, name, msg, current, thresh, op);

    // Record the alert time
    const cell = used.getCell(r, idx.lastAlert);
    cell.setValue(new Date().toISOString().replace("T", " ").slice(0, 16));
  }
}

function resolveCell(wb: ExcelScript.Workbook, ref: string): number | null {
  // Accepts "Sheet!A1" or "A1"
  let sheet: ExcelScript.Worksheet | undefined;
  let addr: string = ref;
  if (ref.includes("!")) {
    const parts = ref.split("!");
    sheet = wb.getWorksheet(parts[0].replace(/'/g, ""));
    addr = parts[1];
  } else {
    sheet = wb.getActiveWorksheet();
  }
  if (!sheet) return null;
  const v = sheet.getRange(addr).getValue();
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function evaluate(v: number, op: string, t: number): boolean {
  switch (op) {
    case ">":  return v > t;
    case "<":  return v < t;
    case ">=": return v >= t;
    case "<=": return v <= t;
    case "=":  return v === t;
    case "<>": return v !== t;
    default:   return false;
  }
}

function getNamedValue(wb: ExcelScript.Workbook, name: string): string {
  const ni = wb.getNamedItem(name);
  if (!ni) return "";
  const range = ni.getRange();
  const v = range ? range.getValue() : ni.getValue();
  return v ? String(v) : "";
}

async function postAdaptiveCard(
  url: string,
  title: string,
  summary: string,
  currentValue: number,
  threshold: number,
  op: string
): Promise<void> {
  const card = {
    type: "MessageCard",
    context: "http://schema.org/extensions",
    themeColor: "D13438",
    summary: title,
    title: `Threshold Breach: ${title}`,
    sections: [
      {
        activityTitle: summary,
        facts: [
          { name: "Value",     value: String(currentValue) },
          { name: "Threshold", value: String(threshold) },
          { name: "Operator",  value: op },
          { name: "At",        value: new Date().toISOString() },
        ],
      },
    ],
  };

  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(card),
  });
  if (!resp.ok) {
    console.log(`Teams post failed: HTTP ${resp.status}`);
  }
}

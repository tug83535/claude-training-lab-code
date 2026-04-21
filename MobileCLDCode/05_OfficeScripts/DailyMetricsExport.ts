/**
 * Office Script: Daily metrics export -> CSV -> SharePoint library.
 *
 * WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
 * ---------------------------------------
 * Excel AutoSave writes the current file, not a dated snapshot. OneDrive
 * version history keeps copies but indexed by timestamp, not by the
 * metrics you care about. Power Automate's "Save to SharePoint" step
 * needs file *content*; this script converts the live sheet's numbers
 * into CSV bytes it can send downstream.
 *
 * PAIR WITH: a scheduled Power Automate flow that runs this script at
 * 18:00, takes the returned string, and saves it as
 *   /sites/Finance/Shared Documents/DailyMetrics/metrics_YYYY-MM-DD.csv
 */

function main(
  workbook: ExcelScript.Workbook,
  sheetName: string = "Dashboard",
  rangeAddress: string = "A1:L40"
): string {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw `Sheet '${sheetName}' not found.`;
  }
  const range = sheet.getRange(rangeAddress);
  const values = range.getValues();

  // Pick up numberFormats so currency and dates survive the CSV round-trip.
  const formats = range.getNumberFormats();

  const rows: string[] = [];
  const date = new Date().toISOString().slice(0, 10);
  rows.push(`# snapshot_date,${date}`);

  for (let r = 0; r < values.length; r++) {
    const row: string[] = [];
    for (let c = 0; c < values[r].length; c++) {
      let v = values[r][c];
      const fmt = formats[r][c];
      if (v instanceof Date) {
        v = v.toISOString().slice(0, 10);
      } else if (typeof v === "number" && fmt && fmt.includes("%")) {
        v = (v * 100).toFixed(2) + "%";
      } else if (typeof v === "number" && fmt && (fmt.includes("$") || fmt.includes(","))) {
        v = v.toFixed(2);
      }
      row.push(csvEscape(String(v ?? "")));
    }
    rows.push(row.join(","));
  }
  return rows.join("\n");
}

function csvEscape(s: string): string {
  if (/[,"\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

/**
 * Office Script: Bulk-format a freshly imported data range.
 *
 * WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
 * ---------------------------------------
 * Paste from Snowflake / a REST API / a CSV into Excel online, and the
 * numbers come in as text, dates come in ISO-8601, and currency loses $
 * signs. Manually re-formatting every week is mind-numbing. This script
 * reads header names and applies the right format automatically:
 *
 *   - "*_date" / "*_at"      -> yyyy-mm-dd
 *   - "*_pct" / "*_rate"     -> 0.00%
 *   - "*_amount" / "*_usd"   -> $#,##0.00
 *   - "*_count" / "qty*"     -> #,##0
 *   - "id" / "*_id"          -> General, left aligned, text-safe
 *
 * Also does:
 *   - Turns every column into a filter/sort-ready Table
 *   - Freezes the header row
 *   - Applies iPipeline brand colours to the header band
 *   - Adds banded rows
 */

function main(workbook: ExcelScript.Workbook): void {
  const sheet = workbook.getActiveWorksheet();
  const used = sheet.getUsedRange();
  if (!used) return;

  const headers = (used.getRow(0).getValues()[0] as string[]).map(h => String(h || "").trim());
  const nRows = used.getRowCount();
  const nCols = used.getColumnCount();

  // 1. Apply a brand table (creates it if none exists)
  const existing = sheet.getTables()[0];
  if (existing) existing.delete();
  const table = workbook.addTable(used, true);
  table.setPredefinedTableStyle("TableStyleMedium2"); // close enough to brand blue

  // 2. Format each column based on the header name
  for (let c = 0; c < nCols; c++) {
    const h = headers[c].toLowerCase();
    const col = sheet.getRangeByIndexes(1, c, nRows - 1, 1);
    const fmt = pickFormat(h);
    if (fmt) col.setNumberFormat(fmt);

    if (/(^id$|_id$)/.test(h)) col.setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    if (/(amount|usd|total|revenue|cost|mrr|arr)/.test(h)) {
      col.setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    }
  }

  // 3. Style the header band with iPipeline Blue
  const headerRange = used.getRow(0);
  headerRange.getFormat().getFill().setColor("#0B4779");
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFont().setBold(true);

  // 4. Freeze header + autofit columns
  sheet.getFreezePanes().freezeRows(1);
  for (let c = 0; c < nCols; c++) {
    sheet.getRangeByIndexes(0, c, 1, 1).getFormat().autofitColumns();
  }

  console.log(`Formatted ${nRows} rows x ${nCols} columns`);
}

function pickFormat(header: string): string | null {
  if (/_(date|at|on)$/.test(header))                return "yyyy-mm-dd";
  if (/(_pct$|_rate$|percentage)/.test(header))      return "0.00%";
  if (/(_amount$|_usd$|revenue|cost|mrr|arr|total)/.test(header)) return "$#,##0.00";
  if (/(^qty$|_count$|customers|users|seats)/.test(header))       return "#,##0";
  if (/(_ms$|latency)/.test(header))                 return "#,##0 \"ms\"";
  if (/(size_bytes|bytes)/.test(header))             return "#,##0 \"B\"";
  return null;
}

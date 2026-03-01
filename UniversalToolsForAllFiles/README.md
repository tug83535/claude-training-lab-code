# Universal Tools For All Files

**Status:** Future Backlog — Not yet packaged

---

## What This Folder Is For

Most VBA modules in this project are built specifically for the
`KeystoneBenefitTech_PL_Model.xlsm` file. They look for exact sheet names,
row positions, and named ranges that only exist in that workbook.

However, a small set of tools are truly universal — they work on any
Excel file regardless of how it is structured. This folder will eventually
hold a standalone Excel Add-In (.xlam) that coworkers can install once
and use across all of their own Excel files.

**This is a future task. It is NOT required for the demo.**

---

## Tools Planned for This Package

The following subs will be extracted and packaged here:

| Module | Sub Name | What It Does |
|--------|----------|--------------|
| modDataSanitizer | RunFullSanitize | Cleans text-stored numbers and floating-point noise. Never touches dates, names, or ID columns. |
| modDataSanitizer | PreviewSanitizeChanges | Dry-run preview before any changes are made. |
| modDataGuards | FindNegativeAmounts | Highlights any negative numbers in red across the whole file. |
| modDataGuards | FindZeroAmounts | Highlights any zero values in yellow across the whole file. |
| modDataGuards | FindSuspiciousRoundNumbers | Flags suspiciously round numbers (e.g., exact $10,000) in orange. |
| modAuditTools | FindExternalLinks | Scans for formulas and hyperlinks pointing to external files. |
| modAuditTools | AuditHiddenSheets | Lists every hidden and very-hidden sheet in the workbook. |
| modSearch | CrossSheetSearch | Searches every sheet for a keyword and highlights matches. |

---

## How Sharing Will Work (Future)

1. The subs above will be copied into a new standalone workbook.
2. That workbook will be saved as `KBT_UniversalTools.xlam` (Excel Add-In format).
3. Coworkers install it once:
   - File → Options → Add-Ins → Manage: Excel Add-Ins → Go → Browse
   - Select `KBT_UniversalTools.xlam` → OK
4. A new ribbon tab or macro shortcut becomes available in every Excel file they open.
5. A step-by-step install guide for non-technical staff will be written at that time.

---

## When to Do This

After the CFO/CEO demo is complete and the main P&L model is delivered.
This is a nice-to-have for coworkers who want the tools on their own files.
It does not affect the demo in any way.

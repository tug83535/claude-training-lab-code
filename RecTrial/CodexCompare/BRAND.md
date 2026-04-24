# iPipeline Brand Styling Guide

All user-facing visual output — Excel sheet styling, PDF reports, video thumbnails, training guide formatting, chart colors, headers — **must** follow these rules. This includes file covers, splash screens, command centers, dashboards, PDF exports, and slide decks.

The goal: anyone looking at a screen should immediately feel "this is iPipeline quality."

---

## Colors

### Primary
| Name | Hex | RGB | Usage |
|---|---|---|---|
| **iPipeline Blue** | `#0B4779` | `11, 71, 121` | Primary headers, top banners, main CTA buttons, chart primary series |

### Secondary
| Name | Hex | RGB | Usage |
|---|---|---|---|
| **Navy** | `#112E51` | `17, 46, 81` | Totals rows, footers, dark text on light backgrounds, secondary headers |
| **Innovation Blue** | `#4B9BCB` | `75, 155, 203` | Hyperlinks, highlights, secondary chart series, accents |

### Accents
| Name | Hex | RGB | Usage |
|---|---|---|---|
| **Lime Green** | `#BFF18C` | `191, 241, 140` | "Pass" / positive indicators, success status, favorable variances |
| **Aqua** | `#2BCCD3` | `43, 204, 211` | Info callouts, tips, informational banners |

### Neutrals
| Name | Hex | RGB | Usage |
|---|---|---|---|
| **Arctic White** | `#F9F9F9` | `249, 249, 249` | Page / sheet backgrounds, alternating row color (lighter of two) |
| **Charcoal** | `#161616` | `22, 22, 22` | Body text, dark text |

### Semantic (Pass/Fail)
- **Pass / OK / Favorable:** Lime Green `#BFF18C`
- **Fail / Warning / Unfavorable:** use a clearly contrasting red-orange. Pick one shade and stick with it across the project (suggestion: `#D64545`). Do not invent per-feature red tones.
- **Neutral / Info:** Aqua `#2BCCD3`

---

## Typography

**Font family: Arial only.** No Calibri, no Times, no Comic Sans, no custom fonts that may not exist on coworker machines.

| Element | Weight | Size range |
|---|---|---|
| Primary headings (H1 on a sheet / cover page) | Arial Bold | 18–24 pt |
| Sub-headings (H2) | Arial Narrow or Arial Bold | 14–16 pt |
| Section labels | Arial Bold | 11–12 pt |
| Body text | Arial Regular | 10–11 pt |
| Data table body | Arial Regular | 9–10 pt |
| Footnotes / captions | Arial Regular | 8–9 pt |

Cell alignment: left for labels, right for numbers, center for headers. Numbers always formatted with consistent decimal places and thousands separators.

---

## Visual Composition Rules

### Sheet / Report Layout
- Every report sheet starts with a **branded header block**: company name + report title + date range + run timestamp.
- Headers use iPipeline Blue background with white text.
- First data row = "header row" styled boldly with Navy text on light Innovation-Blue-tinted background.
- Alternating row shading for tables: Arctic White and a very light Innovation Blue (~10% tint).
- Totals rows: Navy background, white bold text.
- Borders: thin light gray (`#D0D0D0` range) — no thick black borders.

### Excel Number Formatting
- Currency: `$#,##0;($#,##0);"-"` for whole dollars; `$#,##0.00` when cents matter.
- Percentages: `0.0%` (one decimal) for variances; `0%` for share-of-total.
- Dates: `mmm-yy` for monthly labels, `yyyy-mm-dd` for logs/timestamps.

### Charts
- First series uses iPipeline Blue.
- Second series uses Innovation Blue.
- Third series uses Aqua.
- Reference/target lines use Navy dashed.
- No 3D effects. No rainbow color palettes. No default Excel color themes.
- Chart background: Arctic White. Chart title: Arial Bold, 14 pt, Navy.

### PDF / Print
- Portrait orientation unless data demands landscape.
- Header on every page: iPipeline logo area (text placeholder if no image), report title, page X of Y.
- Footer: "Confidential — iPipeline Finance & Accounting" with date/time.
- Margins: 0.5" on all sides.

---

## Logo / Mark

No logo image is provided in this repo. Use a **text treatment** everywhere a logo would go:

```
iPipeline
Finance & Accounting
```

Styling: "iPipeline" in iPipeline Blue, Arial Bold 20pt. "Finance & Accounting" below it in Navy, Arial Regular 10pt.

If you want to support an actual logo file later, reference it via a relative path — never hardcode a drive letter.

---

## Voice & Tone (for Guides, Videos, UI Text)

- Plain English. No jargon unless explained on first use.
- Conversational but professional. You're briefing a smart colleague, not a junior hire and not a board meeting.
- Active voice. Short sentences.
- Avoid exclamation points, emoji, and cutesy language — this is Finance.
- Never say "simply" or "just" when describing a step. It's condescending.

## Banned in All Brand Output

- Comic Sans (obviously)
- Default Excel "Office" theme colors (the pastel blue-orange-green palette)
- Rainbow gradients
- 3D chart effects
- WordArt
- Emoji in Excel cells, PDFs, or official guides (code is fine)
- "Lorem ipsum" placeholder text in anything shipping
- Screenshots with other companies' names/data visible
- Personal email addresses in examples (use `firstname.lastname@ipipeline.com` pattern)

---
name: ipipeline-brand-styling
description: "Apply iPipeline's official corporate brand identity and style guidelines to any output. Use this skill whenever the user asks to style, format, brand, or apply corporate identity to documents, presentations, spreadsheets, PDFs, or web content using iPipeline standards. Trigger on keywords like: branding, corporate identity, visual identity, iPipeline style, brand colors, company styling, on-brand, brand guidelines, visual formatting, visual design, iPipeline template."
---

# iPipeline Brand Styling Skill

Use this skill to apply iPipeline's official 2024 brand guidelines to any output — PowerPoint, Word, PDF, Excel, or HTML/web content.

---

## Brand Colors

### Primary Color
| Name | Hex | RGB |
|------|-----|-----|
| iPipeline Blue | `#0B4779` | R:11 G:71 B:121 |

### Secondary Colors
| Name | Hex | RGB |
|------|-----|-----|
| Navy Blue | `#112E51` | R:17 G:46 B:81 |
| Innovation Blue | `#4B9BCB` | R:75 G:155 B:203 |

### Accent Colors
| Name | Hex | RGB |
|------|-----|-----|
| Lime Green | `#BFF18C` | R:191 G:241 B:140 |
| Aqua | `#2BCCD3` | R:43 G:204 B:211 |

### Neutral Colors
| Name | Hex | RGB |
|------|-----|-----|
| Arctic White | `#F9F9F9` | R:249 G:249 B:249 |
| Charcoal | `#161616` | R:22 G:22 B:22 |

---

## Typography

All fonts are from the Arial family — no custom font installation required.

| Role | Font | Style |
|------|------|-------|
| Headings (large) | Arial Bold | Bold or Bold Italic |
| Subheadings | Arial Narrow | Bold or Bold Italic |
| Body text | Arial | Regular or Italic |

- **Never** use decorative or non-Arial fonts
- Heading text should use iPipeline Blue (`#0B4779`) or Navy Blue (`#112E51`) where possible
- Body text uses Charcoal (`#161616`) on light backgrounds, Arctic White (`#F9F9F9`) on dark

---

## Color Application Guidelines

### Backgrounds
- **Primary slide/page backgrounds**: Arctic White (`#F9F9F9`) or iPipeline Blue (`#0B4779`)
- **Dark sections/headers**: Navy Blue (`#112E51`) or Charcoal (`#161616`)
- **Avoid** using accent colors (Lime Green, Aqua) as full backgrounds

### Text on Backgrounds
- Light background -> Charcoal (`#161616`) or iPipeline Blue (`#0B4779`) text
- Dark background -> Arctic White (`#F9F9F9`) text

### Accents and Highlights
- Use Innovation Blue (`#4B9BCB`) for secondary visual elements, borders, dividers, icons
- Cycle through Lime Green -> Aqua for accent details (charts, callouts, icons)
- Avoid overusing accent colors — they should draw the eye to key information only

### Non-text Shapes and Elements
- Primary shapes: iPipeline Blue (`#0B4779`) or Navy Blue (`#112E51`)
- Accent shapes/icons: cycle through Innovation Blue, Lime Green, Aqua
- Charts: use the full color palette in order — iPipeline Blue, Innovation Blue, Aqua, Lime Green, Navy Blue

---

## Per-Output-Type Guidance

### PowerPoint (.pptx)
- Read `/mnt/skills/public/pptx/SKILL.md` before starting
- Title slide: iPipeline Blue (`#0B4779`) background, Arctic White title text (Arial Bold), Innovation Blue subtitle text
- Content slides: Arctic White background, Charcoal body text, iPipeline Blue headings
- Section dividers: Navy Blue background, Arctic White text
- Charts/graphs: use brand palette in order
- Borders/lines: Innovation Blue (`#4B9BCB`)

### Word Documents (.docx)
- Read `/mnt/skills/public/docx/SKILL.md` before starting
- Heading 1: Arial Bold, iPipeline Blue (`#0B4779`), 18pt+
- Heading 2: Arial Narrow Bold, Navy Blue (`#112E51`), 14pt
- Body: Arial Regular, Charcoal (`#161616`), 11pt
- Accent lines/dividers: Innovation Blue
- Table headers: iPipeline Blue background, Arctic White text

### PDFs
- Read `/mnt/skills/public/pdf/SKILL.md` before starting
- Follow same typography and color rules as Word documents
- Cover page: iPipeline Blue background, Arctic White headline, Innovation Blue subheading

### Excel / Spreadsheets (.xlsx)
- Read `/mnt/skills/public/xlsx/SKILL.md` before starting
- Header rows: iPipeline Blue (`#0B4779`) fill, Arctic White text, Arial Bold
- Alternating rows: Arctic White and Light Gray (use `#F0F0EE` as a soft neutral)
- Totals/summary rows: Navy Blue (`#112E51`) fill, Arctic White text
- Charts: brand palette in order
- Accent highlights: Lime Green or Aqua for callout cells

### HTML / Web Content
- Background: Arctic White (`#F9F9F9`)
- Primary headings: `font-family: Arial, sans-serif; font-weight: bold; color: #0B4779`
- Subheadings: `font-family: Arial Narrow, Arial, sans-serif; font-weight: bold; color: #112E51`
- Body text: `font-family: Arial, sans-serif; color: #161616`
- Links / interactive: `#4B9BCB` (Innovation Blue)
- Buttons (primary): `background: #0B4779; color: #F9F9F9`
- Buttons (secondary): `background: #4B9BCB; color: #F9F9F9`
- Accent/highlight bar: `#2BCCD3` (Aqua) or `#BFF18C` (Lime Green)
- Borders/dividers: `#4B9BCB`

---

## Logo Usage Rules (reference only — do not generate logo)

- Give the logo clear spacing — do not crowd it with other elements
- Use the logo on white or light neutral backgrounds whenever possible
- If placing on a colored background, use the negative (white) logo version
- Never stretch, rotate, or distort the logo
- Never add drop shadows, borders, or effects to the logo

---

## Quick Reference Palette (for code use)

```python
# python-pptx / ReportLab RGB values
IPIPELINE_BLUE    = (11,  71,  121)   # #0B4779 - Primary
NAVY_BLUE         = (17,  46,  81)    # #112E51 - Secondary
INNOVATION_BLUE   = (75,  155, 203)   # #4B9BCB - Secondary
LIME_GREEN        = (191, 241, 140)   # #BFF18C - Accent
AQUA              = (43,  204, 211)   # #2BCCD3 - Accent
ARCTIC_WHITE      = (249, 249, 249)   # #F9F9F9 - Neutral
CHARCOAL          = (22,  22,  22)    # #161616 - Neutral
```

```css
/* CSS variables */
--color-primary:     #0B4779;
--color-navy:        #112E51;
--color-innovation:  #4B9BCB;
--color-lime:        #BFF18C;
--color-aqua:        #2BCCD3;
--color-white:       #F9F9F9;
--color-charcoal:    #161616;
```

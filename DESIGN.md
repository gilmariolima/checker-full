---
name: Conferência de Caixa
description: Painel de conferência de PIX entre extratos bancários (BB, C6) e planilhas de agentes
colors:
  paper: "#f4f1e9"
  paper-line: "#ded8ca"
  card: "#fffdf8"
  ink: "#201d17"
  ink-soft: "#5c5648"
  ink-faint: "#756e5f"
  stamp: "#8a1f3d"
  stamp-deep: "#601530"
  success-ink: "#2f6b3d"
  warning-ink: "#92590c"
  error-ink: "#a3271f"
typography:
  display:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "1.5rem"
    fontWeight: 600
    lineHeight: 1.2
    letterSpacing: "-0.025em"
  headline:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "1.5rem"
    fontWeight: 600
    lineHeight: 1.25
  title:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "1.05rem"
    fontWeight: 600
    lineHeight: 1.3
  body:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "0.85rem"
    fontWeight: 400
    lineHeight: 1.55
  data:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "0.82rem"
    fontWeight: 400
    lineHeight: 1.4
  label:
    fontFamily: "IBM Plex Sans, Segoe UI, sans-serif"
    fontSize: "0.7rem"
    fontWeight: 700
    letterSpacing: "0.08em"
rounded:
  sm: "6px"
  md: "10px"
  lg: "16px"
spacing:
  sm: "8px"
  md: "14px"
  lg: "20px"
  xl: "28px"
components:
  button-primary:
    backgroundColor: "{colors.stamp}"
    textColor: "{colors.paper}"
    rounded: "{rounded.sm}"
    padding: "10px 18px"
  button-primary-hover:
    backgroundColor: "{colors.stamp-deep}"
  button-quiet:
    backgroundColor: "{colors.card}"
    textColor: "{colors.ink}"
    rounded: "{rounded.sm}"
    padding: "10px 18px"
  chip:
    backgroundColor: "{colors.card}"
    rounded: "{rounded.md}"
    padding: "16px 18px"
  card-agent:
    backgroundColor: "{colors.card}"
    rounded: "{rounded.md}"
    padding: "16px 18px"
---

# Design System: Conferência de Caixa

## Overview

**Creative North Star: "Fita de Caixa Registradora" (Register Tape)**

This is a redesign, replacing the earlier "Trust Ledger" fintech-SaaS world (navy header, blue→purple gradient, pill buttons) with a world literal to the product's own name: a "Conferência de Caixa" (cash register reconciliation) rendered as an actual register tape. The page reads as paper, not glass: warm off-white receipt stock, ink-black type, a single stamp-red accent used the way a rubber "CONFERIDO" stamp is used — sparingly, with authority, never as a gradient wash.

Every rectangle in the previous world was a pill or a heavily rounded card; this world cuts paper instead — small, consistent corner radii (6/10/16px), dashed borders standing in for perforation lines, and a literal torn/perforated edge under the header where the "tape" separates from the machine. Status (conferido/falta-PDF/falta-Excel) is still carried by the same three-color vocabulary as before, but rendered as colored text and section grouping rather than thick side borders, which added nothing the text color didn't already say.

**Key Characteristics:**
- Paper-and-ink material world: warm off-white ground, near-black ink text, one stamp-red accent — no gradients anywhere
- Monospace (IBM Plex Sans) for every heading, label, and data value — the "printed" register-tape voice; IBM Plex Sans carries prose only
- Small, consistent corner radii and dashed "perforation" borders replace the old pill/gradient language entirely
- Sector identity reads through a small text-labeled badge (SUP/VALE/CANOA/TOP); status reads through text color and the section it's grouped under — never a thick colored side-border, and never an unlabeled color-only dot

## Colors

Two-material palette: paper and ink, plus one stamp-red accent that carries all brand energy. Status colors are muted, "stamped-ink" versions of green/amber/red rather than bright SaaS pastel.

### Primary
- **Stamp** (#8a1f3d): The one accent — primary CTA, brand mark, focus rings, section accents. Used the way a rubber stamp's ink is used: deliberate, occasional, never a wash.

### Neutral
- **Ink** (#201d17): Header background and primary text — near-black, warm, not blue-black.
- **Ink Soft** (#5c5648): Secondary text, meta lines, subtitles.
- **Ink Faint** (#756e5f): Tertiary text, hints, placeholders.
- **Paper** (#f4f1e9): Page background — warm off-white register-tape stock.
- **Paper Line** (#ded8ca): Borders, dashed perforation lines, dividers.
- **Card** (#fffdf8): Elevated surfaces — panels, cards, entries, inputs. Barely lighter than Paper; elevation here is a border, not a shadow-driven lift.

### Status (semantic, not decorative)
- **Success Ink** (#2f6b3d): Conferido / matched state.
- **Warning Ink** (#92590c): Faltando no PDF.
- **Error Ink** (#a3271f): Faltando no Excel, destructive actions.

### Named Rules
**The One Stamp Rule.** #8a1f3d is the only saturated color used for brand/interactive purposes. It never appears as a gradient, wash, or background fill larger than a button or icon.

**The No-Gradient Rule.** This world has zero gradients, replacing the prior "Trust Ledger" world's blue→purple CTA gradient entirely. Flat ink or flat stamp-red only.

## Typography

**Heading/Data Font:** IBM Plex Sans (with Segoe UI, sans-serif fallback)
**Body Font:** IBM Plex Sans (with Segoe UI, sans-serif fallback)

**Character:** One sans-serif family throughout, with weight and size establishing hierarchy. Financial values use tabular numerals.

### Hierarchy
- **Display** (600, 1.5rem desktop / 1.1rem mobile, Plex Sans, sentence case): Page title in the header.
- **Headline** (600, 1.2rem, Plex Sans): Section titles ("Resultado da análise").
- **Title** (600, 1.1rem panel / .9rem agent, Plex Sans): Panel sub-headings, agent card names.
- **Body** (400, 0.85rem, Plex Sans): Prose — field hints, header subtitle.
- **Data** (400, 0.82rem, Plex Sans): Every monetary value, date, time, agent meta line.
- **Label** (700, 0.7rem, Plex Sans, letter-spacing 0.08em, uppercase): Small tags — field "PDF"/"XLSX" markers, badges.

### Named Rules
**The Tabular-Numbers Rule.** Currency, totals and progress use IBM Plex Sans with tabular numerals.

## Layout

Unchanged from the prior world at the structural level: centered single column, `width: min(1180px, calc(100% - 40px))`, upload panel above results. No eyebrows and no decorative section numbers — they add no real wayfinding value on a single-page tool. A dashed `border-top` on the action bar reads as a tear line rather than a plain divider.

## Elevation & Depth

This world drops shadow-driven elevation almost entirely. Cards and panels are flat paper with a 1px border; the only shadow in the system (`--shadow-sm`) appears on hover of the panel/agent-card, faint and warm-tinted (`rgba(32,29,23,...)`, never blue-black). Depth reads through border and paper-tone contrast, not lift.

### Named Rules
**The Flat-Paper Rule.** Nothing floats. A card sitting on the page looks like paper sitting on paper — bordered, not elevated — until the user's cursor proves it's interactive.

## Shapes

Consistent radii: 6px (inputs, badges, buttons, entries), 10px (cards, agent cards), 16px (outer panel; 10px on mobile). No pills anywhere — the prior world's fully-rounded buttons and badges are replaced by short rectangles that read as cut/stamped paper, not soft UI chrome. Dashed borders (`1px dashed`) mark every "perforation": the upload dropzones, the manual-base box, the export button, the action-bar divider.

### Named Rules
**The Cut-Paper Rule.** Every corner resolves to 6, 10, or 16px. A pill anywhere in this system is a regression to the old world, not a valid exception.

## Components

Register-tape components: rectangular, ink-on-paper, with a stamp-red accent reserved for the one primary action and for status accents.

### Buttons
- **Shape:** Short rectangle, 6px radius, 1px border, minimum height 44px.
- **Primary:** Flat Stamp fill, paper text, no gradient, no shadow lift — `:active` presses to `scale(0.97)` per Emil Kowalski's press-feedback principle, `transition: transform 140ms cubic-bezier(0.23,1,0.32,1)` (transform only, not `all`).
- **Quiet (secondary):** Card background, paper-line border, ink text.
- **Export:** Dashed border (perforation motif) instead of a filled style — visually signals "tear this off," distinct from the two in-app actions.

### Chips (status summary)
- **Style:** Card background, thin 2px colored top rule (not a side border — deliberately avoids the "thick side-tab" look the prior world had), circular icon badge in flat status color.
- **State:** Read-only counters, no selected state.

### Cards / Containers
- **Corner Style:** 10px (agent cards), 16px (outer panel), 6px (entries).
- **Background:** Card on Paper.
- **Elevation:** Border only at rest; faint shadow on hover (see Elevation & Depth).
- **Sector identity:** A small solid-color text badge (`.sector-badge`, e.g. "SUP", "VALE", "CANOA", "TOP") before the agent name — replaces both the original 6px colored left-border and an interim unlabeled color-swatch dot, which real usage showed wasn't self-explanatory. The badge is legible on its own, no legend required, and doubles as a scan target in a long collapsed list.
- **Status identity (entries):** Carried by text color alone (`.text-success`/`.text-warning`/entry's own red `<strong>`) plus which of the three sections (Conferidos/Faltando no PDF/Faltando no Excel) the entry sits under. A per-entry color dot was tried and removed — it duplicated information the text color and section heading already gave, adding clutter without adding clarity.

### Inputs / Fields
- **Style:** 1px paper-line border, 6px radius; native file inputs use Plex Sans and a paper background with a wine-colored selector label.
- **Focus:** Border shifts to Stamp plus a soft `rgba(138,31,61,.14)` glow ring.

### Bank Badge
Small rectangular tag (6px radius, Plex Sans, 700 weight, uppercase-adjacent). BB renders as a warm gold tag; C6 renders as ink-on-paper (near-black background, paper text) — same high-contrast pairing as before, now sharp-cornered instead of pill-shaped for shape-scale consistency.

### Tear Edge (signature element)
A zigzag torn-paper strip (`.tear-edge`, pure CSS triangle pattern) sits between the ink header and the paper body — the single most literal expression of the "register tape" concept, appearing once, at the one seam in the layout where it makes structural sense.

## Do's and Don'ts

### Do:
- **Do** render currency, totals, and progress in IBM Plex Sans — this is the detail that sells the "printed tape" concept.
- **Do** keep the Stamp accent (#8a1f3d) flat — no gradients, no tints used as backgrounds.
- **Do** resolve every corner to 6/10/16px; no pills.
- **Do** use dashed borders specifically to signal a "perforation" (upload zones and dividers) — not as a generic input style.

### Don't:
- **Don't** reintroduce a side-border color accent, or an unlabeled color-only dot, on cards or entries — use a text-labeled badge (like `.sector-badge` or `.bank-badge`) when a color code needs to be scannable without a legend.
- **Don't** mix in the prior world's blue/purple tokens; they are fully retired, not a fallback.
- **Don't** add shadow-driven elevation at rest; this world's depth comes from paper/ink contrast and borders.
- **Don't** use emoji as inline icons in the live UI — Bootstrap Icons is the icon system throughout; the exported PDF summary uses plain text labels instead of icons, since icon fonts are unreliable inside the html2canvas capture path.


## Refinement ? 2026-09-10

The user explicitly chose to retain beige and wine, refining the current identity. This is Operate mode: uploading statements, reviewing agent/sector matches, and exporting results. Keep the paper motif, dark header, receipt edge, Plex Sans typography and Gilmario Lima attribution.

The updated surface uses a lighter beige ground, brighter paper panels, quieter semantic backgrounds, native file controls and a stronger spacing hierarchy. The primary action sits left; export sits right on desktop. File hints and agent metadata use Plex Sans for readability. Sector badges spell out Suporte Online, Vale Viagens, Top Viagens and Canoa in soft semantic colors; do not abbreviate them back to SUP/VALE/TOP.

### Additional semantic surfaces
- PDF / Suporte Online: #f4e8e9 with wine text (#8a1f3d / #842540).
- Excel / Top Viagens / success: #e9eee5 with green text (#2f6b3d / #326140).
- Vale Viagens: #e9edf5 with #254777.
- Canoa: #f4e9ef with #893661.
- Warning icon: #f5ecd9 with #92590c.
- Error icon: #f8e9e4 with #a3271f.

### States and accessibility
- Empty results show an instruction to select files, without fabricated data.
- While processing, file controls and start/clear actions are disabled; the loading message has status semantics.
- Export stays disabled until agent results exist. Clearing resets both counts and currency totals.
- Agent headers expose button semantics, Enter/Space operation, expansion state and a chevron.
- Wine keyboard focus, skip link, 44px action targets and reduced-motion support apply throughout.
- At 800px uploads stack. At 520px the primary action occupies a full row, summary counters remain in three compact columns, and manual inputs stack.

Validation used synthetic API responses in a local browser at 1440px and 390px. Empty, processing, error, successful results, sector names, keyboard expansion, PDF download, manual unmark and clearing passed. No horizontal overflow or JavaScript errors occurred. The mechanical detector ran in its degraded regex mode and reported advisory design-token differences; it did not provide a computed contrast audit.


## Sans-serif typography ? 2026-09-10

At the user?s request, the entire site and PDF export now use IBM Plex Sans, with Segoe UI and generic sans-serif fallbacks. This is an available sans-serif alternative, not the proprietary Claude font. The shared CSS token is `--font-ui`. The monospace font download has been removed. Beige, wine, spacing and interactions remain unchanged.

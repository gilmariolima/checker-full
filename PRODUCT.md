# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Gilmario Lima — sole user. He runs the tool himself at the end of a day or shift to reconcile PIX transactions received against Excel control sheets tracked per agent/sector (e.g. Suporte Online, Vale Viagens, Top Viagens, Canoa). Not built for external clients or a team of operators today.

## Product Purpose

Replace manual, line-by-line reconciliation between bank PDF statements and Excel control sheets with an automated matching pass. The win Gilmario confirmed is speed + precision of the match itself (fuzzy name matching, value tolerance, time-window scoring) — not report generation as the primary value, though a PDF export exists as an output.

## Positioning

Unifies two structurally different bank PDF export formats (Banco do Brasil, C6 Bank) into one reconciliation pipeline against agent-grouped Excel data, and does the name/value/time fuzzy matching a manual spreadsheet comparison can't do quickly or reliably.

## Operating Context

- Input: one or more bank PDF statements (BB or C6) + one or more Excel control sheets, uploaded together per reconciliation run.
- Excel sheets are organized by agent/sector blocks (rows starting with "AGENTE" set the context for following rows).
- Output: three buckets — conferidos (matched), faltando no PDF (in Excel, no PDF match), faltando no Excel (in PDF, unused) — reviewable per agent, with manual override to mark/unmark a match, and a PDF summary export.
- Used ad hoc per reconciliation session, not a standing dashboard.

## Capabilities and Constraints

- Only Banco do Brasil and C6 Bank PDF formats are supported; no other banks today (confirmed — do not assume broader bank support).
- Stateless by design: no database, single-request processing, nothing persisted server-side between runs.
- Statement/spreadsheet contents are sensitive and must not be sent to third parties or logged persistently — confirmed as a binding constraint. **Known gap against this**: `servidor.py`'s `detalhe_bb` currently writes the full extracted PDF text to `pdf_debug.txt` on disk on every run, and both BB/C6 parsers print each parsed name/value/time to stdout for debugging. This predates the confirmed privacy constraint and should be treated as tech debt to flag, not silently normal.
- PT-BR only; no i18n need identified.

## Brand Commitments

Attributed to "Gilmario Lima" in the UI footer, PDF export footer, and page `<meta name="author">`. Keep this attribution in any redesign.

## Evidence on Hand

No screenshots, user research, or existing DESIGN.md. Visual state must be read from the current `frontend/` implementation (Bootstrap 5 + Bootstrap Icons + Inter font), which is mid-refactor as of 2026-08-22 (uncommitted changes to `leitor-extratos.html`, `style.css`, `servidor.py`).

## Product Principles

1. Matching speed and precision come before report polish — the core value is not making the user re-verify the machine's work.
2. Never widen scope to banks/formats that aren't BB or C6 without the user explicitly asking.
3. Treat uploaded statement/spreadsheet content as sensitive: no new persistence, logging, or outbound calls involving it.
4. This is a single-operator tool, not a multi-tenant product — design for one confident daily user, not onboarding or broad audiences.

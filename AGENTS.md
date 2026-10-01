# Project brief: Excel-driven questionnaire review prototype

> This file is the standing context for AI assistants working on this project.
> It is automatically loaded by Posit Assistant; other tools may need to be
> pointed at it explicitly.

## Purpose

A **local browser-only prototype** for reviewing survey/questionnaire answers. It tests three ideas:

1. A full questionnaire can be **generated from metadata stored in an Excel workbook** (.xlsx) rather than hand-built HTML.
2. Standalone, grouped (side-by-side), and sector-repeated questions render correctly from that metadata.
3. **Explicit answer-level versioning + comments + review statuses** provide a workable alternative to Microsoft Word tracked changes for a Respondent ↔ Researcher review workflow.

It is explicitly a prototype: no backend, database, authentication, npm project, or build process. All data lives in the browser's localStorage.

## How to run

Open this folder (`WebApp_test`) in VS Code, install the **Live Server** extension, and serve `frontend/index.html` (right-click → "Open with Live Server"). No build step. Note: `index.html` loads SheetJS from a CDN (`cdn.sheetjs.com/xlsx-0.20.3`), so an internet connection is needed on first load; `frontend/README.md` documents how to run fully offline by storing `xlsx.full.min.js` locally.

## Repository layout

```
frontend/
  index.html    page skeleton, three <dialog> modals (history, comment, import)
  styles.css    plain CSS, restrained professional design, a11y contrast
  app.js        all logic (~1,800 lines, no framework, plain JS + DOM APIs)
  README.md     run instructions, Excel format spec, data model, manual test plan
q_specifications/
  CMP_q_structure.xlsx   the user's real specification workbook (imported at runtime
                         via the file picker; not read from disk by the app)
```

## Excel specification format

The workbook must contain three worksheets, matched **case-insensitively by name** (column headings likewise case-insensitive, surrounding spaces ignored):

- **`questions`**: `question_id` (unique, required), `section`, `subsection`, `question_text`, `question_type` (`text` | `longtext` | `integer` | `decimal` | `boolean` | `date`, case-insensitive), `group_id` (optional), `repeat_by` (blank or `sector`), `required` (TRUE/FALSE/Yes/No/1/0), `help_text`, `display_order` (numeric; missing values sort last in workbook row order).
- **`question_groups`**: `group_id` (unique), `group_text` (shown once above the group), `layout` (only `side_by_side` supported).
- **`sectors`**: `NACE_code` (unique, non-empty), `sector_name`.

Placeholders `{sector_name}` and `{NACE_code}` are resolved in question/group/help text; a `help_text` that is exactly `sector_name` displays the sector name.

On import the workbook is parsed (SheetJS) and **validated first**; a report of errors (blocking) and warnings (non-blocking) is shown in a dialog, and nothing replaces existing data until the user clicks "Use this specification" (with a confirm if response data exists).

## Data model (single localStorage key: `excel-questionnaire-proto-v1`)

One JSON object with deliberately separated structures:

```jsonc
{
  "spec": { "meta": {fileName, importedAt, source, fingerprint}, "questions": [...], "groups": [...], "sectors": [...] },
  "role": "respondent" | "researcher",
  "answers":  { "<fieldInstanceId>": value },
  "versions": { "<fieldInstanceId>": [ /* append-only, immutable */ ] },
  "comments": { "<reviewUnitId>": [ ... ] },
  "statuses": { "<reviewUnitId>": "draft|submitted|revision_requested|revised|accepted" }
}
```

Core concepts:

- **Review unit** = one standalone question instance OR one whole group instance (including each sector-specific clone). A unit owns exactly **one status and one comment thread**. A group never gets per-field comment sections.
- **Field instance id** = `question_id`, or `question_id__NACE_code` when sector-repeated; a field inside a repeated group also carries `group_instance_id` = `group_id__NACE_code`.
- **Versions are keyed by field instance id**; comments/statuses by review unit id. This separation is central: saving a unit appends new immutable versions only for fields whose value actually changed; comments and status changes never create versions; identical saves create nothing.
- Each version stores: `version_id`, `base_question_id`, `question_instance_id`, `group_instance_id`, `NACE_code`, `sector_name`, `version_number`, typed `value`, `author_role`, `timestamp`, `change_note`.
- **Instances are derived, not stored**: `generateInstances(spec)` deterministically rebuilds them from the spec on every render; they are included in the JSON export.

## app.js architecture (all plain functions, organized in numbered sections)

1. **Constants** (columns, types, statuses, roles).
2. **DOM helpers** — `el(tag, options, children)` and `clearNode`. **All rendering uses createElement/textContent; workbook or answer content must never go through innerHTML** (XSS-safe by construction).
3. **Utilities** — timestamp formatting, `required`/`display_order` parsing, spec fingerprint (djb2 hash), placeholder resolution/stripping.
4. **Built-in sample specification** (labelled demonstration data) with seeded answers, a two-version trail on a longtext field, a researcher revision request, comments, and one Accepted item. Lets the UI be tested without an Excel file.
5. **Excel parsing** — `parseWorkbook()` + per-sheet row readers; headers normalized case-insensitively; Excel row numbers retained for validation messages.
6. **Validation** — `validateParsedSpec()` returning `{errors, warnings}` per the rules in README; errors block rendering entirely.
7. **Instance generation** — `generateInstances(spec)`: questions sorted by `display_order`; groups emitted **once**, at the position of their lowest-display_order member (later members skipped via an `emittedGroups` set); `repeat_by = sector` clones one instance per sector row.
8. **State/persistence** — module-level `state`, `loadState()`/`persistState()`, `clearResponseData()`.
9. **Rendering** — sections/subsections drawn only on change; status summary chips; status filter.
10. **Unit rendering** — `renderUnit` (regular cards), `renderField`, `buildFieldInput` (type-appropriate controls), `workflowButtons` (shared), `renderSectorMatrix`/`renderMatrixRow` (compact grid: question text once, one row per sector: name left, control + small buttons right — each row still an independent review unit), `renderCommentThread`.
11. **Editing/versioning** — `handleSaveUnit` (diffs per field; required check deferred to Submit), `createAnswerVersion` (the **only** place versions are written; append-only), `handleSubmitUnit`, accept/reopen.
12. **History dialog** — versions in reverse chronological order, current marked; `wordDiff` (LCS) comparison for text/longtext rendered safely; plain previous-vs-current for numbers/boolean/date.
13. **Comments** — immutable, chronological, role-styled; required comment when a researcher requests a revision.
14. **Workflow** — Respondent edits/saves/submits (blocked when Accepted; required fields must be populated to submit; boolean "No" counts as answered). Researcher requests revision (→ Revision requested), accepts (→ Accepted, fields read-only), reopens (→ Submitted). Researcher never edits answers.
15. **Import/export/resets** — validation-gated import dialog, sample loader, JSON export (spec, instances, answers, versions, comments, statuses, timestamp), reset-responses (keeps spec) and reset-all.

**Event handling**: one delegated click listener on `#questionnaire` dispatches on `button.dataset.action` (`save`, `submit`, `request-revision`, `accept`, `reopen`, `comment`, `history`, `toggle-expand`); dialog forms have their own submit handlers. View-state like expanded longtext boxes lives in a memory-only `Set`, not localStorage.

## Conventions and guardrails for any future change

- Keep it plain HTML/CSS/JS; no frameworks, no build step, no backend.
- Never insert workbook/answer/user content via `innerHTML`; use the `el()` helper.
- Never mutate, reorder, or delete stored versions — only push.
- Keep comments/statuses unit-level and versions field-level; don't collapse the two.
- Accessible markup: labels/`for`, fieldsets+legends for groups, `aria-pressed` role switcher, `aria-expanded` toggles, visible focus states, contrast-checked status badges.
- No `grep`/`sed`/`node` in this environment's shell — verification is done with R scripts via the code-execution tool (bracket balance, ID cross-checks, action-handler coverage).
- `frontend/README.md` is the user-facing reference (Excel format, test plan) — update it when behaviour changes.

## Current state and known limitations

All requested features work end-to-end, including: import + validation, sample data, role switching, append-only versioning, word-level diff, comments, status workflow with accept/reopen, sector matrix rendering, longtext expand/collapse toggle, and JSON export. Known limitations (documented in README): importing a new spec clears responses (with warning); JSON export only (no import); researchers cannot edit answers; diff is word-level; data is single-browser and not secure/production-ready.

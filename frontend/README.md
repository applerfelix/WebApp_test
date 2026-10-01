# Excel-driven questionnaire review — local prototype

A browser-based prototype that generates a questionnaire from an Excel
specification workbook, with answer-level versioning, comments and
review statuses as an alternative to Microsoft Word tracked changes.

Everything runs in the browser. There is **no backend, no database, no
authentication, no build step and no npm project**. All data is stored
in the browser's `localStorage`.

---

## Running the prototype (VS Code + Live Server)

1. Open this folder in VS Code (**File → Open Folder**).
2. Install the **Live Server** extension by Ritwick Dey (publisher:
   *ritwickdey.liveserver*) if you do not have it already
   (Extensions view → search "Live Server" → Install).
3. Open `index.html`, right-click in the editor and choose
   **"Open with Live Server"** — or click **"Go Live"** in the blue
   status bar at the bottom of the window.
4. Your browser opens the prototype (typically
   `http://127.0.0.1:5500`).

Opening `index.html` directly from the file system usually works too,
but Live Server is the supported way to run the prototype.

### Internet connection requirement

`index.html` loads **SheetJS** (the .xlsx reader) from a CDN:

```html
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
```

An **internet connection is required** the first time the page is
loaded (the browser caches the script afterwards).

### Running fully offline (optional)

1. Download the SheetJS standalone script from:
   `https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js`
2. Save it **into this folder** as `xlsx.full.min.js`.
3. In `index.html`, replace the CDN `<script>` tag above with:

   ```html
   <script src="xlsx.full.min.js"></script>
   ```

The prototype then works with no internet connection at all.

---

## Excel workbook format

The workbook must contain three worksheets (or sheets containing Excel
tables), matched **case-insensitively** by name. Column headings are
matched **case-insensitively**, ignoring leading/trailing spaces.

### 1. `questions`

| Column | Meaning |
|---|---|
| `question_id` | **Required.** Unique, non-empty machine-readable id. |
| `section` | Main section heading. Rendered when it changes. |
| `subsection` | Optional subsection heading. |
| `question_text` | Question text. May be empty for grouped questions (shared text comes from `group_text`). Supports `{sector_name}` and `{NACE_code}` placeholders. |
| `question_type` | **Required.** One of: `text`, `longtext`, `integer`, `decimal`, `boolean`, `date` (case-insensitive). |
| `group_id` | Optional. If set, the question is a field inside that group. All fields of a group form one review unit. |
| `repeat_by` | Blank (render once) or `sector` (repeat once per row of `sectors`). Case-insensitive. |
| `required` | `TRUE/FALSE`, `true/false`, `Yes/No`, `yes/no` or `1/0`. Required fields must be populated before a unit can be submitted (drafts may be incomplete; for boolean fields "No" counts as populated). |
| `help_text` | Optional supporting text. `{sector_name}`, `{NACE_code}` placeholders are replaced; a value that is exactly `sector_name` displays the current sector name. |
| `display_order` | Numeric sort order. Missing values sort after valid ones, preserving workbook row order. Non-numeric values are validation errors. |

### 2. `question_groups`

| Column | Meaning |
|---|---|
| `group_id` | **Required.** Unique group identifier referenced by `questions.group_id`. |
| `group_text` | **Required.** Shared text displayed once above the group's fields. |
| `layout` | `side_by_side` puts the group's fields in columns on wide screens (stacked on narrow screens). Unknown or blank layouts stack vertically. |

Group rules:

- A group is rendered **once**, at the position of its member with the
  lowest `display_order`; its fields are ordered by their own
  `display_order`.
- The group stays in the section/subsection of its first member; a
  warning is shown if members conflict.
- If **all** members have `repeat_by = sector`, the whole group is
  repeated once per sector (`group_id__NACE_code`), each with its own
  status and comment thread.
- Inconsistent `repeat_by` values inside one group are a validation
  **error**.

### 3. `sectors`

| Column | Meaning |
|---|---|
| `NACE_code` | **Required.** Unique, non-empty NACE code. |
| `sector_name` | **Required.** Sector display name. |

A question with `repeat_by = sector` produces one instance per sector
row, with instance id `question_id__NACE_code` (e.g.
`active_in_sector__61`).

**Compact rendering:** a standalone sector-repeated question is shown as
ONE card — the question text appears once (sector placeholders removed,
so prefer wording without placeholders) — with one row per sector:
sector name and NACE code on the left, the answer control, status badge
and compact workflow buttons on the right. Each row remains an
independent review unit with its own status, comment thread and
field-level version history.

---

## Review model

- **Review unit** = one standalone question instance, or one complete
  group instance (including every sector-specific instance). Each unit
  has **one status and one comment thread**.
- **Field-level versioning**: every field of a unit has its own
  immutable, append-only version history. Saving a unit creates new
  versions only for fields whose values changed; identical values,
  comments and status changes never create versions.
- **Statuses**: Draft → Submitted → Revision requested → Revised →
  Accepted (plus a Reopen action from Accepted).
- **Roles**: switch with the buttons at the top.
  - *Respondent*: edits and saves answers (unless the unit is
    Accepted), submits for review, revises after a revision request,
    adds comments.
  - *Researcher*: reviews answers and histories, adds comments,
    requests revisions (a comment is required), accepts items, reopens
    accepted items. Researchers do not edit respondent answers.
- **Version comparison**: word-level diff for `text`/`longtext`
  (removed words struck through, added words highlighted);
  previous/current side by side for `integer`, `decimal`, `boolean`
  and `date`.

---

## Data storage and its limitations

Everything is stored in a single namespaced localStorage key:

```
excel-questionnaire-proto-v1
```

The stored object separates: the imported specification (with metadata
and fingerprint), current answers, immutable answer versions, comments,
review statuses and the active role. Generated question instances are
derived from the specification on every load (deterministically) and
are included in the JSON export.

**localStorage is NOT secure and NOT production-ready:**

- Any script on the page (or anyone with access to the browser) can
  read and change it.
- Data never leaves the browser; it is not shared between browsers,
  machines or users, and clearing browser data deletes it.
- There is no backup; use **Export all data (JSON)** to keep a copy.

---

## Toolbar buttons

| Button | Effect |
|---|---|
| Import Excel specification | Select an `.xlsx` file; it is parsed and validated first, and a summary is shown. Nothing is applied until you click "Use this specification". Applying a new specification clears existing responses (you are warned first). |
| Load sample specification | Loads the built-in **demonstration data** (sample questions, groups, sectors, versions, comments and statuses) so the interface can be tested without an Excel file. |
| Reload saved specification | Re-reads everything from localStorage and re-renders. |
| Export all data (JSON) | Downloads specification, generated instances, current answers, all answer versions, comments, statuses and an export timestamp as a JSON file. |
| Reset responses and history | Clears answers, versions, comments and statuses; keeps the specification. |
| Reset all prototype data | Clears everything, including the specification. |

---

## Manual test plan

Run with Live Server. After each step check the notice line at the top
and the status summary.

1. **Import a valid workbook**: click "Load sample specification"
   first (or import a valid workbook), confirm the summary line shows
   questions, groups, sectors and review-unit counts.
2. **Standalone questions**: verify the `text`, `integer` and `boolean`
   standalone questions each render with the correct input type, help
   text, required marker and their own status badge and comment thread.
3. **Side-by-side group renders once**: the "Please describe the
   remote-working policy" group appears once, with its `group_text`
   shown once and both fields (boolean + longtext) side by side on a
   wide screen (stack them by narrowing the window). No second copy of
   the group appears anywhere.
4. **Sector expansion**: `active_in_sector` renders as ONE compact card
   with one row per sector (61, 62, 63) — question text once, sector
   names on the left, Yes/No controls on the right — where each row
   keeps its own status, history and comment thread (verify by adding
   comments to two different sectors, or by filtering by status).
5. **Switch roles**: click Respondent/Researcher; the active button
   highlights, and the available actions per unit change (respondent:
   save/submit; researcher: request revision/accept/reopen).
6. **Save a new answer version**: as Respondent, edit
   `employees_total`, add a change note, save; the current-answer line
   shows "version 2" and "View history" shows two versions with the
   newest marked Current.
7. **Unchanged answer creates no version**: save again without changing
   anything; the notice says no changes were saved and no new version
   appears in the history.
8. **Request a revision**: as Researcher, click "Request revision" on
   a Submitted unit; leaving the comment empty is rejected, saving adds
   a Researcher comment and sets the status to Revision requested.
9. **Group-level comment**: add a comment on the group unit; it appears
   once under the group (not under each field), alongside the sample
   researcher/respondent comments.
10. **Submit a revised answer**: as Respondent, change the longtext
    field in the group and save; the status becomes Revised, then
    "Submit for review" sets it to Submitted. (The sample description
    field already contains the version 1 → version 2 revision trail.)
11. **Compare two text versions**: open "View history" on the policy
    description field, click "Compare with previous version"; removed
    and added words are highlighted distinctly.
12. **Accept and reopen**: as Researcher, Accept a Submitted unit (the
    fields become read-only for the respondent, sample unit `org_name`
    shows this state), then use "Reopen" to make it editable again.
13. **Export all data**: click "Export all data (JSON)" and check the
    downloaded file contains the specification, generated instances,
    answers, all versions, comments, statuses and a timestamp.
14. **Persistence**: press F5; everything (specification, answers,
    versions, comments, statuses, role) is still there.
15. **Invalid workbook**: import an `.xlsx` that is missing a sheet or a
    column, has a duplicate `question_id`, an unknown `question_type`,
    or inconsistent `repeat_by` within a group; the import dialog lists
    errors and the questionnaire is **not** updated.

---

## Known limitations (by design for this stage)

- Applying a new specification clears existing responses (you are
  warned and asked to confirm first).
- Only JSON export exists; there is no JSON import.
- Researchers cannot edit answers at all in this prototype.
- The diff is word-level; line breaks in long text are normalised in
  the comparison view.
- One specification, one browser: no multi-user collaboration.

'use strict';

/* =====================================================================
   Excel-driven questionnaire prototype.
   Plain JavaScript, no frameworks, no backend. All data lives in
   localStorage under one namespaced key (see STORAGE_KEY).

   Data model (one JSON object in localStorage):
   {
     spec: {                                    // imported specification
       meta: { fileName, importedAt, source, fingerprint },
       questions: [...], groups: [...], sectors: [...]
     },
     role: 'respondent' | 'researcher',         // active role
     answers:  { [fieldInstanceId]: value },    // current values
     versions: { [fieldInstanceId]: [...] },    // append-only history
     comments: { [reviewUnitId]: [...] },       // separate from versions
     statuses: { [reviewUnitId]: status }
   }

   Key separation of concerns:
   - ANSWER VERSIONS are keyed by FIELD instance id and are immutable:
     createAnswerVersion() only pushes; nothing in this file ever
     mutates, removes or rewrites a stored version.
   - COMMENTS and REVIEW STATUSES are keyed by REVIEW UNIT id (a
     standalone question instance, or a whole group instance). One
     group shares one thread and one status; its fields each have
     their own version history.
   - Question INSTANCES are derived from the spec on every render,
     never stored, and are included in the JSON export.

   All rendering uses DOM APIs (createElement/textContent). Workbook
   and answer content never passes through innerHTML.
   ===================================================================== */

/* ---------------------------------------------------------------------
   1. Constants
   ------------------------------------------------------------------- */

const STORAGE_KEY = 'excel-questionnaire-proto-v1';

const QUESTION_TYPES = ['text', 'longtext', 'integer', 'decimal', 'boolean', 'date'];
const REPEAT_BY_VALUES = ['', 'sector'];
const KNOWN_LAYOUTS = ['side_by_side'];

const STATUS_ORDER = ['draft', 'submitted', 'revision_requested', 'revised', 'accepted'];
const STATUS_LABELS = {
  draft: 'Draft',
  submitted: 'Submitted',
  revision_requested: 'Revision requested',
  revised: 'Revised',
  accepted: 'Accepted',
};

const ROLES = { respondent: 'Respondent', researcher: 'Researcher' };

/* Expected columns per worksheet. Column names are matched
   case-insensitively, ignoring surrounding spaces. */
const QUESTION_COLUMNS = [
  'question_id', 'section', 'subsection', 'question_text', 'question_type',
  'group_id', 'repeat_by', 'required', 'help_text', 'display_order',
];
const GROUP_COLUMNS = ['group_id', 'group_text', 'layout'];
const SECTOR_COLUMNS = ['NACE_code', 'sector_name'];

/* A missing column here is a validation ERROR; missing optional
   columns only produce warnings or defaults. */
const REQUIRED_COLUMNS = {
  questions: ['question_id', 'question_type'],
  question_groups: ['group_id', 'group_text'],
  sectors: ['NACE_code', 'sector_name'],
};

/* ---------------------------------------------------------------------
   2. DOM helpers (safe rendering)
   ------------------------------------------------------------------- */

/**
 * Create an element with options and children.
 * All untrusted content goes in via `text` (textContent), never
 * through innerHTML, so workbook and answer values cannot inject HTML.
 */
function el(tag, options, children) {
  const node = document.createElement(tag);
  if (options) {
    if (options.class) node.className = options.class;
    if (options.text !== undefined && options.text !== null) {
      node.textContent = String(options.text);
    }
    if (options.id) node.id = options.id;
    if (options.attrs) {
      for (const [name, value] of Object.entries(options.attrs)) {
        node.setAttribute(name, value);
      }
    }
    if (options.props) {
      for (const [name, value] of Object.entries(options.props)) {
        node[name] = value;
      }
    }
    if (options.dataset) {
      for (const [name, value] of Object.entries(options.dataset)) {
        node.dataset[name] = value;
      }
    }
    if (options.onClick) node.addEventListener('click', options.onClick);
  }
  if (children) {
    for (const child of children) {
      if (child) node.append(child);
    }
  }
  return node;
}

function clearNode(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

/* ---------------------------------------------------------------------
   3. Utilities
   ------------------------------------------------------------------- */

function formatTimestamp(iso) {
  return new Date(iso).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
}

/** Accepted forms: TRUE/FALSE, true/false, Yes/No, yes/no, 1/0. */
function parseRequiredValue(value) {
  if (value === true) return true;
  if (value === false) return false;
  if (typeof value === 'number') return value === 1;
  const text = String(value ?? '').trim().toLowerCase();
  if (['true', 'yes', '1'].includes(text)) return true;
  if (['false', 'no', '0'].includes(text)) return false;
  return false;
}

/**
 * display_order parsing:
 * - blank/missing  -> null (sort after questions that have a value)
 * - valid number   -> the number
 * - anything else  -> NaN (present but unusable -> validation error)
 */
function parseDisplayOrder(value) {
  if (value === '' || value === null || value === undefined) return null;
  const number = typeof value === 'number' ? value : Number(String(value).trim());
  return Number.isFinite(number) ? number : NaN;
}

/** Small deterministic hash (djb2) used as the specification fingerprint. */
function computeFingerprint(obj) {
  const text = JSON.stringify(obj);
  let hash = 5381;
  for (let i = 0; i < text.length; i++) {
    hash = ((hash << 5) + hash + text.charCodeAt(i)) >>> 0;
  }
  return hash.toString(16);
}

/**
 * Resolve placeholders in question_text, group_text and help_text:
 * - {sector_name} and {NACE_code} are replaced with the current sector.
 * - A help_text that is exactly "sector_name" displays the sector name
 *   itself (kept for compatibility with older specification files).
 */
function resolvePlaceholders(text, sector) {
  if (text === null || text === undefined) return '';
  const value = String(text);
  if (!sector) return value;
  if (value.trim() === 'sector_name') return sector.sector_name;
  return value
    .replaceAll('{sector_name}', sector.sector_name)
    .replaceAll('{NACE_code}', sector.NACE_code);
}

/** Human-readable display of a stored answer value. */
function displayAnswerValue(value) {
  if (value === '' || value === null || value === undefined) return '(empty)';
  return String(value);
}

/**
 * For the compact sector-matrix card the shared question text is shown
 * ONCE, so sector placeholders are removed and leftover wording is
 * tidied up (extra spaces, space before punctuation). For best results
 * write sector-repeated question text without placeholders.
 */
function stripSectorPlaceholders(text) {
  return String(text ?? '')
    .replaceAll('{sector_name}', '')
    .replaceAll('{NACE_code}', '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([?.!:;,])/g, '$1')
    .trim();
}

/* ---------------------------------------------------------------------
   4. Built-in sample specification (demonstration data)
   ------------------------------------------------------------------- */

/** Normalised sample spec, same shape as an imported workbook.
   The sample is labelled as demonstration data wherever it is shown. */
function buildSampleSpec() {
  const questions = [
    {
      question_id: 'org_name',
      section: 'About the organisation',
      subsection: '',
      question_text: 'What is the name of your organisation?',
      question_type: 'text',
      group_id: '',
      repeat_by: '',
      required: true,
      help_text: 'The registered legal name.',
      display_order: 1,
    },
    {
      question_id: 'employees_total',
      section: 'About the organisation',
      subsection: '',
      question_text: 'How many employees does your organisation have?',
      question_type: 'integer',
      group_id: '',
      repeat_by: '',
      required: true,
      help_text: 'Total headcount at 31 December 2025.',
      display_order: 2,
    },
    {
      question_id: 'has_remote_work_policy',
      section: 'Remote working',
      subsection: 'General',
      question_text: 'Does the organisation have a remote-working policy?',
      question_type: 'boolean',
      group_id: '',
      repeat_by: '',
      required: false,
      help_text: '',
      display_order: 3,
    },
    {
      question_id: 'remote_policy_exists',
      section: 'Remote working',
      subsection: 'Policy details',
      question_text: 'Has the policy been formally approved?',
      question_type: 'boolean',
      group_id: 'remote_policy_details',
      repeat_by: '',
      required: true,
      help_text: '',
      display_order: 4,
    },
    {
      question_id: 'remote_policy_description',
      section: 'Remote working',
      subsection: 'Policy details',
      question_text: 'Policy description',
      question_type: 'longtext',
      group_id: 'remote_policy_details',
      repeat_by: '',
      required: true,
      help_text: 'Describe the main provisions of the policy.',
      display_order: 5,
    },
    {
      question_id: 'active_in_sector',
      section: 'Remote working',
      subsection: 'By sector',
      question_text: 'Is the organisation active in the following sectors?',
      question_type: 'boolean',
      group_id: '',
      repeat_by: 'sector',
      required: false,
      help_text: 'sector_name',
      display_order: 6,
    },
  ];

  const groups = [
    {
      group_id: 'remote_policy_details',
      group_text: 'Please describe the remote-working policy',
      layout: 'side_by_side',
    },
  ];

  const sectors = [
    { NACE_code: '61', sector_name: 'Telecommunications' },
    { NACE_code: '62', sector_name: 'Software development' },
    { NACE_code: '63', sector_name: 'Information services' },
  ];

  return {
    meta: {
      fileName: 'Built-in sample specification (demonstration data)',
      source: 'sample',
      importedAt: '2026-10-01T14:00:00',
      fingerprint: computeFingerprint({ questions, groups, sectors }),
    },
    questions,
    groups,
    sectors,
  };
}

function makeSampleVersion(opts) {
  const questionInstanceId = opts.NACE
    ? `${opts.questionId}__${opts.NACE}`
    : opts.questionId;
  return {
    version_id: `${questionInstanceId}__v${opts.number}__sample`,
    base_question_id: opts.questionId,
    question_instance_id: questionInstanceId,
    group_instance_id: opts.groupInstanceId || null,
    NACE_code: opts.NACE || null,
    sector_name: opts.sectorName || null,
    version_number: opts.number,
    value: opts.value,
    author_role: opts.role,
    timestamp: opts.timestamp,
    change_note: opts.note || '',
  };
}

function makeSampleComment(opts) {
  return {
    comment_id: `${opts.unitId}__c${opts.number}__sample`,
    review_unit_id: opts.unitId,
    author_role: opts.role,
    timestamp: opts.timestamp,
    comment_text: opts.text,
  };
}

/**
 * Seed sample answers, versions, comments and statuses. Includes:
 * - two versions of the policy description (revision request flow),
 * - a researcher revision request comment,
 * - a revised respondent answer,
 * - comment threads and one Accepted item (org_name).
 */
function seedSampleResponses() {
  const descriptionV1 =
    'Employees may work from home up to two days per week with approval ' +
    'from their line manager. Requests are considered on a case-by-case basis.';
  const descriptionV2 =
    'Employees may work from home up to two days per week with approval ' +
    'from their line manager. Final approval authority rests with the head ' +
    'of department. Requests are considered on a case-by-case basis and the ' +
    'policy applies to all permanent staff members. The policy took effect ' +
    'on 1 January 2026.';

  state.answers = {
    org_name: 'Acme Research Institute',
    employees_total: 245,
    has_remote_work_policy: 'Yes',
    remote_policy_exists: 'Yes',
    remote_policy_description: descriptionV2,
    active_in_sector__61: 'Yes',
    active_in_sector__62: 'Yes',
    active_in_sector__63: 'No',
  };

  state.versions = {
    org_name: [
      makeSampleVersion({ questionId: 'org_name', number: 1, value: 'Acme Research Institute', role: 'Respondent', timestamp: '2026-09-14T09:12:00' }),
    ],
    employees_total: [
      makeSampleVersion({ questionId: 'employees_total', number: 1, value: 245, role: 'Respondent', timestamp: '2026-09-14T09:20:00' }),
    ],
    has_remote_work_policy: [
      makeSampleVersion({ questionId: 'has_remote_work_policy', number: 1, value: 'Yes', role: 'Respondent', timestamp: '2026-09-14T09:25:00' }),
    ],
    remote_policy_exists: [
      makeSampleVersion({ questionId: 'remote_policy_exists', groupInstanceId: 'remote_policy_details', number: 1, value: 'Yes', role: 'Respondent', timestamp: '2026-09-14T09:45:00' }),
    ],
    remote_policy_description: [
      makeSampleVersion({ questionId: 'remote_policy_description', groupInstanceId: 'remote_policy_details', number: 1, value: descriptionV1, role: 'Respondent', timestamp: '2026-09-14T09:50:00' }),
      makeSampleVersion({ questionId: 'remote_policy_description', groupInstanceId: 'remote_policy_details', number: 2, value: descriptionV2, role: 'Respondent', timestamp: '2026-09-18T14:05:00', note: 'Revised answer following the reviewer comment.' }),
    ],
    active_in_sector__61: [
      makeSampleVersion({ questionId: 'active_in_sector', NACE: '61', sectorName: 'Telecommunications', number: 1, value: 'Yes', role: 'Respondent', timestamp: '2026-09-14T10:00:00' }),
    ],
    active_in_sector__62: [
      makeSampleVersion({ questionId: 'active_in_sector', NACE: '62', sectorName: 'Software development', number: 1, value: 'Yes', role: 'Respondent', timestamp: '2026-09-14T10:00:00' }),
    ],
    active_in_sector__63: [
      makeSampleVersion({ questionId: 'active_in_sector', NACE: '63', sectorName: 'Information services', number: 1, value: 'No', role: 'Respondent', timestamp: '2026-09-14T10:00:00' }),
    ],
  };

  state.comments = {
    has_remote_work_policy: [
      makeSampleComment({ unitId: 'has_remote_work_policy', number: 1, role: 'Researcher', text: 'Please confirm whether external contractors are also covered by this policy.', timestamp: '2026-09-16T11:00:00' }),
    ],
    remote_policy_details: [
      makeSampleComment({ unitId: 'remote_policy_details', number: 1, role: 'Researcher', text: 'Please clarify who has final authority to approve requests, whether the policy covers all staff groups, and state the effective date of the policy.', timestamp: '2026-09-16T11:30:00' }),
      makeSampleComment({ unitId: 'remote_policy_details', number: 2, role: 'Respondent', text: 'The description has been revised to cover these points.', timestamp: '2026-09-18T14:10:00' }),
    ],
    active_in_sector__62: [
      makeSampleComment({ unitId: 'active_in_sector__62', number: 1, role: 'Researcher', text: 'Please double-check activity in this sector.', timestamp: '2026-09-17T09:00:00' }),
    ],
  };

  state.statuses = {
    org_name: 'accepted',
    employees_total: 'submitted',
    has_remote_work_policy: 'revision_requested',
    remote_policy_details: 'revised',
    active_in_sector__61: 'submitted',
    active_in_sector__62: 'submitted',
    active_in_sector__63: 'draft',
  };
}

/* ---------------------------------------------------------------------
   5. Excel parsing (SheetJS)
   ------------------------------------------------------------------- */

/** Find a worksheet by name, case-insensitively, ignoring spaces. */
function findSheet(workbook, wantedName) {
  const target = wantedName.trim().toLowerCase();
  for (const name of workbook.SheetNames) {
    if (name.trim().toLowerCase() === target) return workbook.Sheets[name];
  }
  return null;
}

/**
 * Read a worksheet into row objects keyed by the canonical column
 * names. Header cells are matched case-insensitively after trimming
 * surrounding spaces. Returns null when the sheet does not exist.
 */
function readSheetRows(workbook, sheetName, columns) {
  const sheet = findSheet(workbook, sheetName);
  if (!sheet) return null;

  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: '' });
  if (!matrix.length) {
    return { rows: [], missingColumns: columns.slice() };
  }

  const headers = matrix[0].map((h) => String(h ?? '').trim().toLowerCase());
  const columnIndex = {};
  for (const column of columns) {
    columnIndex[column] = headers.indexOf(column.toLowerCase());
  }

  const rows = [];
  for (let i = 1; i < matrix.length; i++) {
    const raw = matrix[i];
    if (!raw || raw.every((cell) => cell === '' || cell === null || cell === undefined)) continue;
    const row = { _excelRow: i + 1 }; // matrix row i is Excel row i + 1
    for (const column of columns) {
      const index = columnIndex[column];
      row[column] = index === -1 ? '' : (raw[index] ?? '');
    }
    rows.push(row);
  }

  const missingColumns = columns.filter((column) => columnIndex[column] === -1);
  return { rows, missingColumns };
}

function normalizeQuestionRow(row) {
  return {
    _excelRow: row._excelRow,
    _displayOrderRaw: String(row.display_order ?? ''),
    question_id: String(row.question_id ?? '').trim(),
    section: String(row.section ?? '').trim(),
    subsection: String(row.subsection ?? '').trim(),
    question_text: String(row.question_text ?? '').trim(),
    question_type: String(row.question_type ?? '').trim().toLowerCase(),
    group_id: String(row.group_id ?? '').trim(),
    repeat_by: String(row.repeat_by ?? '').trim().toLowerCase(),
    required: parseRequiredValue(row.required),
    help_text: String(row.help_text ?? '').trim(),
    display_order: parseDisplayOrder(row.display_order),
  };
}

function normalizeGroupRow(row) {
  return {
    _excelRow: row._excelRow,
    group_id: String(row.group_id ?? '').trim(),
    group_text: String(row.group_text ?? '').trim(),
    layout: String(row.layout ?? '').trim().toLowerCase(),
  };
}

function normalizeSectorRow(row) {
  return {
    _excelRow: row._excelRow,
    NACE_code: String(row.NACE_code ?? '').trim(),
    sector_name: String(row.sector_name ?? '').trim(),
  };
}

/** Parse an .xlsx ArrayBuffer into normalised questions/groups/sectors. */
function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const questionsSheet = readSheetRows(workbook, 'questions', QUESTION_COLUMNS);
  const groupsSheet = readSheetRows(workbook, 'question_groups', GROUP_COLUMNS);
  const sectorsSheet = readSheetRows(workbook, 'sectors', SECTOR_COLUMNS);

  return {
    presence: {
      questions: !!questionsSheet,
      question_groups: !!groupsSheet,
      sectors: !!sectorsSheet,
    },
    missingColumns: {
      questions: questionsSheet ? questionsSheet.missingColumns : [],
      question_groups: groupsSheet ? groupsSheet.missingColumns : [],
      sectors: sectorsSheet ? sectorsSheet.missingColumns : [],
    },
    questions: questionsSheet ? questionsSheet.rows.map(normalizeQuestionRow) : [],
    groups: groupsSheet ? groupsSheet.rows.map(normalizeGroupRow) : [],
    sectors: sectorsSheet ? sectorsSheet.rows.map(normalizeSectorRow) : [],
  };
}

/* ---------------------------------------------------------------------
   6. Specification validation
   ------------------------------------------------------------------- */

/**
 * Validate a parsed workbook. Returns { errors, warnings }.
 * Errors block rendering; warnings do not.
 */
function validateParsedSpec(parsed) {
  const errors = [];
  const warnings = [];

  const usesSectors = parsed.questions.some((q) => q.repeat_by === 'sector');

  /* --- Worksheet presence and required columns --------------------- */
  if (!parsed.presence.questions) {
    errors.push("Required worksheet 'questions' is missing.");
  } else {
    for (const column of REQUIRED_COLUMNS.questions) {
      if (parsed.missingColumns.questions.includes(column)) {
        errors.push(`Worksheet 'questions' is missing required column '${column}'.`);
      }
    }
  }
  if (!parsed.presence.question_groups) {
    warnings.push("Worksheet 'question_groups' was not found; grouped questions are unavailable.");
  } else {
    for (const column of REQUIRED_COLUMNS.question_groups) {
      if (parsed.missingColumns.question_groups.includes(column)) {
        errors.push(`Worksheet 'question_groups' is missing required column '${column}'.`);
      }
    }
  }
  if (!parsed.presence.sectors && !usesSectors) {
    warnings.push("Worksheet 'sectors' was not found; sector-repeated questions cannot be expanded.");
  } else if (parsed.presence.sectors) {
    for (const column of REQUIRED_COLUMNS.sectors) {
      if (parsed.missingColumns.sectors.includes(column)) {
        errors.push(`Worksheet 'sectors' is missing required column '${column}'.`);
      }
    }
  }

  /* --- Question rows ------------------------------------------------ */
  const questionIdCounts = new Map();
  for (const q of parsed.questions) {
    if (!q.question_id) {
      errors.push(`questions, Excel row ${q._excelRow}: question_id is empty.`);
    } else {
      questionIdCounts.set(q.question_id, (questionIdCounts.get(q.question_id) || 0) + 1);
    }
    if (!q.question_type) {
      errors.push(`questions, Excel row ${q._excelRow}: question_type is empty.`);
    } else if (!QUESTION_TYPES.includes(q.question_type)) {
      errors.push(`questions, Excel row ${q._excelRow}: unknown question_type '${q.question_type}'. Supported: ${QUESTION_TYPES.join(', ')}.`);
    }
    if (!REPEAT_BY_VALUES.includes(q.repeat_by)) {
      errors.push(`questions, Excel row ${q._excelRow}: unsupported repeat_by '${q.repeat_by}'. Supported: (blank) or 'sector'.`);
    }
    if (Number.isNaN(q.display_order)) {
      errors.push(`questions, Excel row ${q._excelRow}: display_order '${q._displayOrderRaw}' is not a valid number.`);
    }
    if (!q.group_id && !q.question_text) {
      warnings.push(`questions, Excel row ${q._excelRow}: standalone question '${q.question_id}' has no question_text.`);
    }
    if (q.group_id && !q.question_text) {
      warnings.push(`questions, Excel row ${q._excelRow}: grouped field '${q.question_id}' has no field-level question_text; a generated label will be used.`);
    }
    if (!q.help_text) {
      warnings.push(`questions, Excel row ${q._excelRow}: help_text is empty for '${q.question_id}'.`);
    }
  }
  for (const [id, count] of questionIdCounts) {
    if (count > 1) errors.push(`questions: duplicate question_id '${id}' (${count} definitions).`);
  }

  /* --- Sectors -------------------------------------------------------- */
  if (usesSectors && parsed.sectors.length === 0) {
    errors.push("repeat_by = 'sector' is used, but no sectors are available (the 'sectors' worksheet is missing or empty).");
  }
  const naceCounts = new Map();
  for (const s of parsed.sectors) {
    if (!s.NACE_code) {
      errors.push(`sectors, Excel row ${s._excelRow}: NACE_code is empty.`);
    } else {
      naceCounts.set(s.NACE_code, (naceCounts.get(s.NACE_code) || 0) + 1);
    }
    if (!s.sector_name) {
      warnings.push(`sectors, Excel row ${s._excelRow}: sector_name is empty for NACE code '${s.NACE_code}'.`);
    }
  }
  for (const [code, count] of naceCounts) {
    if (count > 1) errors.push(`sectors: duplicate NACE_code '${code}' (${count} rows).`);
  }

  /* --- Groups --------------------------------------------------------- */
  const groupIdCounts = new Map();
  for (const g of parsed.groups) {
    if (!g.group_id) {
      errors.push(`question_groups, Excel row ${g._excelRow}: group_id is empty.`);
    } else {
      groupIdCounts.set(g.group_id, (groupIdCounts.get(g.group_id) || 0) + 1);
    }
    if (!g.group_text) {
      warnings.push(`question_groups: group '${g.group_id}' has no group_text.`);
    }
    if (g.layout && !KNOWN_LAYOUTS.includes(g.layout)) {
      warnings.push(`question_groups: group '${g.group_id}' uses unknown layout '${g.layout}'; fields will be stacked.`);
    }
  }
  for (const [id, count] of groupIdCounts) {
    if (count > 1) errors.push(`question_groups: duplicate group_id '${id}' (${count} rows).`);
  }

  /* --- Group references and membership -------------------------------- */
  const knownGroupIds = new Set(parsed.groups.map((g) => g.group_id));
  const questionsByGroup = new Map();
  for (const q of parsed.questions) {
    if (!q.group_id) continue;
    if (!knownGroupIds.has(q.group_id)) {
      errors.push(`questions, Excel row ${q._excelRow}: question '${q.question_id}' references group_id '${q.group_id}', which does not exist in question_groups.`);
      continue;
    }
    if (!questionsByGroup.has(q.group_id)) questionsByGroup.set(q.group_id, []);
    questionsByGroup.get(q.group_id).push(q);
  }
  for (const [groupId, members] of questionsByGroup) {
    const sortedMembers = sortQuestionsByDisplayOrder(members);
    const first = sortedMembers[0];
    for (const m of members) {
      if (m.section !== first.section || m.subsection !== first.subsection) {
        warnings.push(`question group '${groupId}': field '${m.question_id}' has a different section/subsection from the first field; the group stays in the section of its first field ('${first.section}').`);
      }
    }
    // Inconsistent repeat_by within one group is never silently resolved.
    const repeatValues = new Set(members.map((m) => m.repeat_by));
    if (repeatValues.size > 1) {
      const listed = [...repeatValues].map((v) => v || '(blank)').join(', ');
      errors.push(`question group '${groupId}': inconsistent repeat_by values (${listed}). All fields in a group must use the same repeat_by.`);
    }
    if (members.length === 1) {
      warnings.push(`question group '${groupId}' contains only one field.`);
    }
  }

  return { errors, warnings };
}

/* ---------------------------------------------------------------------
   7. Instance generation
   ------------------------------------------------------------------- */

/**
 * Sort questions by display_order. Questions without a display_order
 * go after all questions that have one, preserving workbook row order.
 */
function sortQuestionsByDisplayOrder(questions) {
  return [...questions].sort((a, b) => {
    const ao = a.display_order ?? Infinity;
    const bo = b.display_order ?? Infinity;
    if (ao !== bo) return ao - bo;
    // Workbook row order breaks ties. The built-in sample has no
    // _excelRow (it is never displayed in validation messages).
    return (a._excelRow ?? 0) - (b._excelRow ?? 0);
  });
}

/** One answerable field, with sector context resolved. */
function buildField(question, sector, groupInstanceId) {
  // Field instance id: question_id, or question_id__NACE_code when
  // the question is repeated by sector.
  const questionInstanceId = sector
    ? `${question.question_id}__${sector.NACE_code}`
    : question.question_id;

  const hasOwnLabelText = Boolean(question.question_text);
  return {
    questionInstanceId,
    baseQuestionId: question.question_id,
    groupInstanceId: groupInstanceId || null,
    NACE_code: sector ? sector.NACE_code : null,
    sector_name: sector ? sector.sector_name : null,
    questionType: question.question_type,
    required: question.required,
    // For grouped fields without question_text, generate an accessible
    // but visually quiet label from the type and the question id.
    fieldLabel: hasOwnLabelText
      ? resolvePlaceholders(question.question_text, sector)
      : `${question.question_type} answer (${question.question_id})`,
    fieldLabelSubtle: !hasOwnLabelText,
    helpText: resolvePlaceholders(question.help_text, sector),
  };
}

/** A standalone question becomes one review unit (one field). */
function buildQuestionInstance(question, sector) {
  return {
    id: sector ? `${question.question_id}__${sector.NACE_code}` : question.question_id,
    kind: 'question',
    baseId: question.question_id,
    section: question.section,
    subsection: question.subsection,
    heading: resolvePlaceholders(question.question_text || `Question ${question.question_id}`, sector),
    NACE_code: sector ? sector.NACE_code : null,
    sector_name: sector ? sector.sector_name : null,
    fields: [buildField(question, sector, null)],
  };
}

/**
 * A group becomes one review unit containing several fields. The unit
 * carries ONE status and ONE comment thread; each field keeps its own
 * answers and version history (keyed by questionInstanceId).
 */
function buildGroupInstance(group, members, sector) {
  const groupInstanceId = sector
    ? `${group.group_id}__${sector.NACE_code}`
    : group.group_id;
  return {
    id: groupInstanceId,
    kind: 'group',
    baseId: group.group_id,
    layout: group.layout,
    section: members[0].section,
    subsection: members[0].subsection,
    heading: resolvePlaceholders(group.group_text || `Group ${group.group_id}`, sector),
    NACE_code: sector ? sector.NACE_code : null,
    sector_name: sector ? sector.sector_name : null,
    fields: members.map((m) => buildField(m, sector, groupInstanceId)),
  };
}

/**
 * Generate all review units from the specification.
 *
 * Rendering-order logic for groups:
 * 1. All questions are sorted by display_order.
 * 2. Walking that order, a standalone question is emitted directly.
 * 3. When a group's FIRST member (its lowest display_order member) is
 *    reached, the COMPLETE group is emitted at that position — this is
 *    what determines the group's position in the questionnaire.
 * 4. Later members of the same group are skipped (emittedGroups set),
 *    so every group renders exactly once.
 * 5. Fields inside the group are ordered by their own display_order.
 *
 * Sector repeats: a question or group with repeat_by = 'sector' is
 * cloned once per row of the sectors table, with the pattern
 * question_id__NACE_code / group_id__NACE_code. Each clone is an
 * independent review unit with its own status and comment thread.
 */
function generateInstances(spec) {
  const questions = sortQuestionsByDisplayOrder(spec.questions);
  const groupsById = new Map(spec.groups.map((g) => [g.group_id, g]));
  const instances = [];
  const emittedGroups = new Set();

  for (const question of questions) {
    if (question.group_id && groupsById.has(question.group_id)) {
      if (emittedGroups.has(question.group_id)) continue; // group already emitted
      emittedGroups.add(question.group_id);

      const group = groupsById.get(question.group_id);
      const members = questions.filter((m) => m.group_id === question.group_id); // already sorted
      // Validation guarantees all members share the same repeat_by.
      const repeated = members[0].repeat_by === 'sector';
      const sectorList = repeated ? spec.sectors : [null];
      for (const sector of sectorList) {
        instances.push(buildGroupInstance(group, members, sector));
      }
    } else if (question.repeat_by === 'sector') {
      for (const sector of spec.sectors) {
        instances.push(buildQuestionInstance(question, sector));
      }
    } else {
      instances.push(buildQuestionInstance(question, null));
    }
  }
  return instances;
}

/* ---------------------------------------------------------------------
   8. State and persistence
   ------------------------------------------------------------------- */

const EMPTY_STATE = () => ({
  spec: null,
  role: 'respondent',
  answers: {},
  versions: {},
  comments: {},
  statuses: {},
});

let state = EMPTY_STATE();
let instances = [];      // derived from state.spec on every render
let statusFilter = 'all';
let pendingImportSpec = null;

// Field ids of longtext answer boxes the user has expanded. View
// preference only — kept in memory, not questionnaire data.
const expandedLongtextFields = new Set();

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return {
        spec: parsed.spec || null,
        role: parsed.role === 'researcher' ? 'researcher' : 'respondent',
        answers: parsed.answers || {},
        versions: parsed.versions || {},
        comments: parsed.comments || {},
        statuses: parsed.statuses || {},
      };
    }
  } catch (err) {
    console.warn('Could not read saved data; starting fresh.', err);
  }
  return EMPTY_STATE();
}

function persistState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

/** Clear answers, versions, comments and statuses (statuses become Draft). */
function clearResponseData() {
  state.answers = {};
  state.versions = {};
  state.comments = {};
  state.statuses = {};
}

function hasExistingResponseData() {
  return (
    Object.keys(state.answers).length > 0 ||
    Object.keys(state.versions).length > 0 ||
    Object.keys(state.comments).length > 0 ||
    Object.keys(state.statuses).length > 0
  );
}

function currentAnswer(fieldInstanceId) {
  return state.answers[fieldInstanceId] ?? '';
}

function statusOf(reviewUnitId) {
  return state.statuses[reviewUnitId] || 'draft';
}

/** Current version of a field = last element of its append-only array. */
function currentVersionOf(fieldInstanceId) {
  const history = state.versions[fieldInstanceId];
  return history && history.length ? history[history.length - 1] : null;
}

/* ---------------------------------------------------------------------
   9. App rendering
   ------------------------------------------------------------------- */

function renderApp() {
  instances = state.spec ? generateInstances(state.spec) : [];
  renderRoleSwitcher();
  renderSpecInfo();
  renderSummary();
  renderQuestionnaire();
  updateToolbarState();
}

function renderRoleSwitcher() {
  document.getElementById('role-respondent').setAttribute(
    'aria-pressed',
    String(state.role === 'respondent')
  );
  document.getElementById('role-researcher').setAttribute(
    'aria-pressed',
    String(state.role === 'researcher')
  );
}

function renderSpecInfo() {
  const info = document.getElementById('spec-info');
  if (!state.spec) {
    info.textContent = 'No specification loaded. Import an .xlsx workbook or load the built-in sample to explore the interface.';
    return;
  }
  const meta = state.spec.meta;
  const sourceLabel = meta.source === 'sample'
    ? 'built-in sample (demonstration data)'
    : 'Excel workbook';
  info.textContent =
    `Specification: ${meta.fileName} — ${sourceLabel} · imported ${formatTimestamp(meta.importedAt)}` +
    ` · fingerprint ${meta.fingerprint} · ${state.spec.questions.length} questions, ` +
    `${state.spec.groups.length} groups, ${state.spec.sectors.length} sectors, ${instances.length} review units.`;
}

function renderSummary() {
  const counts = Object.fromEntries(STATUS_ORDER.map((s) => [s, 0]));
  for (const inst of instances) counts[statusOf(inst.id)]++;
  document.getElementById('status-summary').replaceChildren(
    ...STATUS_ORDER.map((status) =>
      el('span', { class: 'summary-chip' }, [
        el('span', { class: `status-badge status-${status}`, text: STATUS_LABELS[status] }),
        el('span', { class: 'summary-count', text: counts[status] }),
      ])
    )
  );

  // Form-level submit is a Respondent action, only when a spec is loaded.
  document.getElementById('btn-form-submit').hidden =
    !(state.role === 'respondent' && state.spec);
}

function renderQuestionnaire() {
  const main = document.getElementById('questionnaire');
  clearNode(main);
  if (!state.spec) {
    main.append(buildEmptyState());
    return;
  }
  const filtered = instances.filter(
    (inst) => statusFilter === 'all' || statusOf(inst.id) === statusFilter
  );
  if (!filtered.length) {
    main.append(el('p', { class: 'muted', text: 'No review units match the current status filter.' }));
    return;
  }

  // Section and subsection headings appear only when the value changes;
  // a new section resets the subsection tracker.
  let lastSection = null;
  let lastSubsection = null;

  // Sector-repeated questions: every sector instance of one base
  // question is drawn inside ONE compact matrix card, so the base id
  // is tracked to make sure the card is rendered only once.
  const renderedMatrixBases = new Set();

  for (const inst of filtered) {
    if (inst.section && inst.section !== lastSection) {
      main.append(el('h2', { class: 'section-heading', text: inst.section }));
      lastSection = inst.section;
      lastSubsection = null;
    }
    if (inst.subsection && inst.subsection !== lastSubsection) {
      main.append(el('h3', { class: 'subsection-heading', text: inst.subsection }));
      lastSubsection = inst.subsection;
    }
    if (isSectorQuestionInstance(inst)) {
      if (renderedMatrixBases.has(inst.baseId)) continue; // row already shown
      renderedMatrixBases.add(inst.baseId);
      const rows = filtered.filter(
        (i) => i.baseId === inst.baseId && isSectorQuestionInstance(i)
      );
      main.append(renderSectorMatrix(rows));
      continue;
    }
    main.append(renderUnit(inst));
  }
}

function buildEmptyState() {
  return el('div', { class: 'empty-state' }, [
    el('h2', { text: 'No specification loaded' }),
    el('p', { text: 'This prototype generates a questionnaire from an Excel workbook containing three worksheets: questions, question_groups and sectors.' }),
    el('ol', {}, [
      el('li', { text: 'Prepare a workbook (.xlsx) with the sheets and columns described in README.md.' }),
      el('li', { text: 'Use "Import Excel specification" in the toolbar to select the file. The workbook is parsed and validated first; nothing is replaced until you confirm.' }),
      el('li', { text: 'Alternatively, click the button below to explore the interface with built-in demonstration data.' }),
    ]),
    el('button', { class: 'primary', text: 'Load sample specification', onClick: loadSample }),
  ]);
}

/* ---------------------------------------------------------------------
   10. Review unit rendering
   ------------------------------------------------------------------- */

function renderUnit(inst) {
  const status = statusOf(inst.id);
  // Respondent edit rule: editable unless the unit is Accepted.
  // Researchers review; they do not edit respondent answers.
  const editable = state.role === 'respondent' && status !== 'accepted';

  const card = el('article', { class: 'question-card', dataset: { unitId: inst.id } });

  const header = el('div', { class: 'question-head' });
  const title = el('h3', { class: 'question-title', text: inst.heading });
  // Standalone questions carry the required indicator on the heading:
  // their field label is rendered visually hidden (see renderField) so
  // the question text is not shown twice.
  if (inst.kind === 'question' && inst.fields[0].required) {
    title.append(el('span', { class: 'required-mark', attrs: { 'aria-hidden': 'true' }, text: ' *' }));
  }
  header.append(title);
  if (inst.NACE_code) {
    header.append(el('span', { class: 'sector-badge', text: `${inst.sector_name} (NACE ${inst.NACE_code})` }));
  }
  header.append(el('span', { class: `status-badge status-${status}`, text: STATUS_LABELS[status] }));
  card.append(header);

  // Grouped units wrap their fields in a fieldset (accessibility);
  // side_by_side groups use a responsive grid (see styles.css).
  const fieldsClass = inst.kind === 'group' && inst.layout === 'side_by_side'
    ? 'group-fields side-by-side'
    : 'group-fields';
  const fieldsContainer =
    inst.kind === 'group'
      ? el('fieldset', { class: fieldsClass }, [el('legend', { class: 'visually-hidden', text: inst.heading })])
      : el('div', { class: fieldsClass });
  for (const field of inst.fields) {
    fieldsContainer.append(renderField(inst, field, editable));
  }
  card.append(fieldsContainer);

  card.append(renderUnitActions(inst, status));
  card.append(renderCommentThread(inst));
  return card;
}

function renderField(inst, field, editable) {
  const wrapper = el('div', { class: 'form-field' });
  const inputId = `input-${field.questionInstanceId}`;

  // Standalone questions already show their text as the card heading, so
  // their field label stays accessible but visually hidden. Grouped
  // fields keep a visible field-level label below the shared group text.
  const labelClass = inst.kind === 'question'
    ? 'field-label visually-hidden'
    : (field.fieldLabelSubtle ? 'field-label subtle' : 'field-label');
  const label = el('label', {
    class: labelClass,
    attrs: { for: inputId },
    text: field.fieldLabel,
  });
  if (field.required) {
    label.append(el('span', { class: 'required-mark', attrs: { 'aria-hidden': 'true' }, text: ' *' }));
  }
  wrapper.append(label, buildFieldInput(field, editable));
  if (field.questionType === 'longtext') {
    wrapper.append(buildExpandButton(inst, field));
    wrapper.append(buildCommentSelectionButton(inst, field));
  }

  if (field.helpText) {
    wrapper.append(el('p', { class: 'help-text', text: field.helpText }));
  }

  // Current answer and its version number, always visible.
  const version = currentVersionOf(field.questionInstanceId);
  const currentLine = el('p', { class: 'current-answer-line' });
  if (version) {
    currentLine.append(
      el('span', {
        class: 'current-answer-label',
        text: `Current answer (version ${version.version_number}, ${version.author_role}, ${formatTimestamp(version.timestamp)}): `,
      }),
      el('span', { class: 'current-answer-value', text: displayAnswerValue(version.value) })
    );
  } else {
    currentLine.append(el('span', { class: 'muted', text: 'No answer saved yet.' }));
  }
  wrapper.append(currentLine);

  wrapper.append(el('button', {
    class: 'secondary small',
    text: 'View history',
    attrs: { 'aria-label': `View answer history for ${field.fieldLabel}` },
    dataset: { action: 'history', unitId: inst.id, fieldId: field.questionInstanceId },
  }));
  return wrapper;
}

/** Build the input control matching the question type. */
function buildFieldInput(field, editable) {
  const value = currentAnswer(field.questionInstanceId);
  const inputId = `input-${field.questionInstanceId}`;
  const attrs = {};
  const props = { value: value === null || value === undefined ? '' : value };
  let tag = 'input';
  let type = 'text';

  switch (field.questionType) {
    case 'longtext': tag = 'textarea'; break;
    case 'integer': type = 'number'; attrs.step = '1'; break;
    case 'decimal': type = 'number'; attrs.step = 'any'; break;
    case 'date': type = 'date'; break;
    case 'boolean': tag = 'select'; break;
    default: type = 'text';
  }
  if (field.required) attrs['aria-required'] = 'true';
  if (!editable) props.disabled = true;

  if (tag === 'textarea') {
    // Longtext boxes start collapsed (2 rows) to keep the page compact
    // and can be expanded with the toggle button next to them.
    const isExpanded = expandedLongtextFields.has(field.questionInstanceId);
    return el('textarea', {
      class: isExpanded ? 'input longtext expanded' : 'input longtext',
      id: inputId,
      attrs: Object.assign({ rows: '2' }, attrs),
      props, dataset: { fieldId: field.questionInstanceId },
    });
  }
  if (tag === 'select') {
    return el('select', { class: 'input', id: inputId, attrs, props, dataset: { fieldId: field.questionInstanceId } }, [
      el('option', { props: { value: '' }, text: '— select —' }),
      el('option', { props: { value: 'Yes', selected: value === 'Yes' }, text: 'Yes' }),
      el('option', { props: { value: 'No', selected: value === 'No' }, text: 'No' }),
    ]);
  }
  return el('input', {
    class: 'input', id: inputId,
    attrs: Object.assign({ type }, attrs),
    props, dataset: { fieldId: field.questionInstanceId },
  });
}

/**
 * Expand/collapse toggle for longtext answer boxes. Toggling happens in
 * place (see the 'toggle-expand' case in handleQuestionnaireClick), so
 * the text being typed and the scroll position are not disturbed.
 */
function buildExpandButton(inst, field) {
  const fieldId = field.questionInstanceId;
  const isExpanded = expandedLongtextFields.has(fieldId);
  return el('button', {
    class: 'secondary small expand-toggle',
    text: isExpanded ? 'Collapse' : 'Expand',
    attrs: {
      'aria-expanded': String(isExpanded),
      'aria-controls': `input-${fieldId}`,
    },
    dataset: { action: 'toggle-expand', unitId: inst.id, fieldId },
  });
}

/**
 * "Comment on selection" button for longtext fields. It opens the comment
 * dialog anchored to whatever text is selected in the answer box.
 */
function buildCommentSelectionButton(inst, field) {
  const fieldId = field.questionInstanceId;
  return el('button', {
    class: 'secondary small comment-selection-toggle',
    text: 'Comment on selection',
    attrs: {
      'aria-controls': `input-${fieldId}`,
      title: 'Select part of the answer text, then click to comment on that selection',
    },
    dataset: { action: 'comment-selection', unitId: inst.id, fieldId },
  });
}

/**
 * Workflow buttons shared by regular cards and compact sector rows.
 * They carry the same data-action values, so the single delegated click
 * handler (handleQuestionnaireClick) dispatches both layouts.
 */
function workflowButtons(inst, status, options) {
  const small = options && options.small ? ' small' : '';
  const make = (classes, text, action, title) =>
    el('button', {
      class: classes + small,
      text,
      attrs: title ? { title } : undefined,
      dataset: { action, unitId: inst.id },
    });

  const buttons = [];
  if (state.role === 'respondent' && status !== 'accepted') {
    // Submission happens at the form level (see handleFormSubmit), not
    // per unit — so each unit only offers a Save button.
    buttons.push(make('primary', 'Save answers', 'save'));
  }
  if (state.role === 'researcher') {
    if (status !== 'accepted' && status !== 'revision_requested') {
      buttons.push(make('secondary', 'Request revision', 'request-revision',
        'Ask the respondent to revise this item; a comment is required'));
    }
    if (status === 'submitted' || status === 'revised') {
      buttons.push(make('primary', 'Accept', 'accept'));
    }
    if (status === 'accepted') {
      buttons.push(make('secondary', 'Reopen', 'reopen',
        'Reopen this item so the respondent can edit again'));
    }
  }
  buttons.push(make('secondary', 'Add comment', 'comment'));
  return buttons;
}

/** Workflow buttons for the unit, plus the respondent change note. */
function renderUnitActions(inst, status) {
  const area = el('div', { class: 'unit-actions' });
  if (state.role === 'respondent' && status !== 'accepted') {
    const noteId = `note-${inst.id}`;
    area.append(
      el('label', { class: 'change-note-label', attrs: { for: noteId }, text: 'Change note (optional)' }),
      el('input', {
        class: 'change-note', id: noteId,
        attrs: { type: 'text', maxlength: '200', placeholder: 'Why the change?' },
      })
    );
  }
  area.append(...workflowButtons(inst, status));
  return area;
}

/** A standalone question instance that is repeated once per sector. */
function isSectorQuestionInstance(inst) {
  return inst.kind === 'question' && inst.NACE_code !== null;
}

/**
 * Render all sector instances of one repeated standalone question as a
 * single compact card:
 * - the question text is shown ONCE at the top (sector placeholders
 *   removed, see stripSectorPlaceholders);
 * - one row per sector shows the sector name on the left and the
 *   answer control, status badge and compact workflow buttons on the
 *   right.
 * Presentation only: each row REMAINS an independent review unit with
 * its own status, its own comment thread and its own field-level
 * version history (keyed by question_id__NACE_code).
 */
function renderSectorMatrix(rows) {
  const first = rows[0];
  const baseQuestion = state.spec.questions.find(
    (q) => q.question_id === first.baseId
  );
  const heading = stripSectorPlaceholders(
    baseQuestion && baseQuestion.question_text
      ? baseQuestion.question_text
      : `Question ${first.baseId}`
  );

  const card = el('article', { class: 'question-card sector-matrix' });
  card.append(el('h3', { class: 'question-title', text: heading }));

  // Shared help text is shown once below the heading. A help_text that
  // only carries the sector placeholder disappears here, because the
  // sector name is already the row label on the left.
  const sharedHelp = baseQuestion ? stripSectorPlaceholders(baseQuestion.help_text) : '';
  if (sharedHelp) {
    card.append(el('p', { class: 'help-text', text: sharedHelp }));
  }

  const table = el('fieldset', { class: 'matrix-table' }, [
    el('legend', { class: 'visually-hidden', text: heading }),
  ]);
  for (const inst of rows) table.append(renderMatrixRow(inst));
  card.append(table);

  // Comment threads stay per sector instance; only sectors that
  // actually have comments show a thread, which keeps the card compact.
  for (const inst of rows) {
    if ((state.comments[inst.id] || []).length) {
      card.append(
        renderCommentThread(inst, `Comments — ${inst.sector_name} (NACE ${inst.NACE_code})`)
      );
    }
  }
  return card;
}

/** One sector row of the compact matrix card. */
function renderMatrixRow(inst) {
  const status = statusOf(inst.id);
  const editable = state.role === 'respondent' && status !== 'accepted';
  const field = inst.fields[0]; // exactly one field per sector instance
  const inputId = `input-${field.questionInstanceId}`;

  const row = el('div', { class: 'matrix-row', dataset: { unitId: inst.id } });

  // Left column: the sector name from the sectors table as the label.
  // Per-row help text is deliberately not repeated here.
  const label = el('label', { class: 'field-label', attrs: { for: inputId }, text: inst.sector_name });
  if (field.required) {
    label.append(el('span', { class: 'required-mark', attrs: { 'aria-hidden': 'true' }, text: ' *' }));
  }
  const version = currentVersionOf(field.questionInstanceId);
  row.append(el('div', { class: 'matrix-label' }, [
    label,
    el('span', {
      class: 'matrix-meta',
      text: `NACE ${inst.NACE_code}` +
        (version ? ` · version ${version.version_number}` : ' · not answered'),
    }),
  ]));

  // Middle column: the answer control (longtext boxes get a toggle).
  const inputCell = el('div', { class: 'matrix-input' }, [buildFieldInput(field, editable)]);
  if (field.questionType === 'longtext') {
    inputCell.append(buildExpandButton(inst, field));
    inputCell.append(buildCommentSelectionButton(inst, field));
  }
  row.append(inputCell);

  // Right column: current status and compact workflow buttons.
  row.append(el('div', { class: 'matrix-actions' }, [
    el('span', { class: `status-badge status-${status}`, text: STATUS_LABELS[status] }),
    ...workflowButtons(inst, status, { small: true }),
    el('button', {
      class: 'secondary small',
      text: 'View history',
      attrs: { 'aria-label': `View answer history for ${inst.sector_name}` },
      dataset: { action: 'history', unitId: inst.id, fieldId: field.questionInstanceId },
    }),
  ]));
  return row;
}

/** One comment thread per review unit (never per field). The heading
    text can be customised (used by the compact sector matrix). */
function renderCommentThread(inst, headingText) {
  const section = el('div', { class: 'comments' });
  section.append(el('h4', { class: 'comments-heading', text: headingText || 'Comments' }));
  const comments = state.comments[inst.id] || [];
  if (!comments.length) {
    section.append(el('p', { class: 'muted small', text: 'No comments yet.' }));
    return section;
  }
  // Chronological order: comments read like a conversation.
  for (const comment of comments) {
    const body = [
      el('p', { class: 'comment-meta', text: `${comment.author_role} · ${formatTimestamp(comment.timestamp)}` }),
    ];
    // Anchored comments quote the selected span and offer a "Locate"
    // button that jumps to it in the answer.
    if (comment.anchor_text) {
      body.push(el('blockquote', { class: 'comment-anchor', text: comment.anchor_text }));
      body.push(el('button', {
        class: 'secondary small locate-anchor',
        text: 'Locate in answer',
        attrs: { 'aria-label': 'Scroll to and highlight the quoted text in the answer' },
        dataset: {
          action: 'locate-anchor',
          unitId: inst.id,
          fieldId: comment.field_id,
          anchorStart: String(comment.anchor_start),
          anchorEnd: String(comment.anchor_end),
        },
      }));
    }
    body.push(el('p', { class: 'comment-text', text: comment.comment_text }));
    section.append(el('div', { class: `comment comment-${comment.author_role.toLowerCase()}` }, body));
  }
  return section;
}

/* ---------------------------------------------------------------------
   11. Editing and version creation
   ------------------------------------------------------------------- */

/** Read and type-check one input control. */
function readFieldInput(input, field) {
  const type = field.questionType;
  if (type === 'boolean') return { ok: true, value: input.value }; // '' | 'Yes' | 'No'
  const text = input.value.trim();
  if (type === 'integer') {
    if (text === '') return { ok: true, value: '' };
    const number = Number(text);
    if (!Number.isFinite(number) || !Number.isInteger(number)) {
      return { ok: false, error: 'enter a whole number.' };
    }
    return { ok: true, value: number };
  }
  if (type === 'decimal') {
    if (text === '') return { ok: true, value: '' };
    const number = Number(text);
    if (!Number.isFinite(number)) return { ok: false, error: 'enter a valid number.' };
    return { ok: true, value: number };
  }
  return { ok: true, value: text }; // text, longtext, date ('' or yyyy-mm-dd)
}

/**
 * Save all fields of a review unit.
 * - Required fields may be empty in a DRAFT (required is enforced at
 *   submission time, see handleFormSubmit).
 * - A new immutable version is appended ONLY for fields whose value
 *   actually changed; identical answers never create versions.
 * - Comments and status changes never happen here.
 */
/**
 * Read a unit's field inputs, validate them, and write answer versions
 * for any field whose value changed. Shared by the per-unit Save button
 * and the form-level Submit (which saves before submitting).
 *
 * Returns { ok, changedCount, error, field }:
 *  - ok=false  → a field failed to parse; `field` is the offending field.
 *  - ok=true   → `changedCount` fields were saved (0 means no changes).
 *
 * This function does NOT persist or re-render; the caller decides.
 */
function saveUnitFields(inst, changeNote) {
  const updates = [];
  for (const field of inst.fields) {
    const input = document.getElementById(`input-${field.questionInstanceId}`);
    const parsed = readFieldInput(input, field);
    if (!parsed.ok) {
      return { ok: false, changedCount: 0, error: parsed.error, field };
    }
    updates.push({ field, value: parsed.value });
  }

  const changed = updates.filter(
    (u) => u.value !== currentAnswer(u.field.questionInstanceId)
  );
  for (const update of changed) {
    state.answers[update.field.questionInstanceId] = update.value;
    createAnswerVersion(update.field, update.value, changeNote);
  }
  return { ok: true, changedCount: changed.length, error: null, field: null };
}

function handleSaveUnit(inst) {
  if (state.role !== 'respondent' || statusOf(inst.id) === 'accepted') return;

  const noteInput = document.getElementById(`note-${inst.id}`);
  const changeNote = noteInput ? noteInput.value.trim() : '';

  const result = saveUnitFields(inst, changeNote);
  if (!result.ok) {
    setNotice(`Cannot save: ${result.field.fieldLabel} — ${result.error}`);
    document.getElementById(`input-${result.field.questionInstanceId}`).focus();
    return;
  }
  if (!result.changedCount) {
    setNotice('No changes to save — the answers are identical to the current versions.');
    return;
  }

  // Respondent saving a changed answer after a revision request moves
  // the unit to Revised.
  if (statusOf(inst.id) === 'revision_requested') {
    state.statuses[inst.id] = 'revised';
  }

  persistState();
  renderApp();
  setNotice(`Saved ${result.changedCount} new answer version${result.changedCount === 1 ? '' : 's'} for "${inst.heading}".`);
}

/**
 * Append an immutable answer version. This is the ONLY place versions
 * are written; the array is append-only and no other code path can
 * modify, reorder or delete stored versions.
 */
function createAnswerVersion(field, value, changeNote) {
  const history = state.versions[field.questionInstanceId] || [];
  const version = {
    version_id: `${field.questionInstanceId}__v${history.length + 1}__${Date.now()}`,
    base_question_id: field.baseQuestionId,
    question_instance_id: field.questionInstanceId,
    group_instance_id: field.groupInstanceId,
    NACE_code: field.NACE_code,
    sector_name: field.sector_name,
    version_number: history.length + 1,
    value,
    author_role: ROLES[state.role],
    timestamp: new Date().toISOString(),
    change_note: changeNote || '',
  };
  history.push(version);
  state.versions[field.questionInstanceId] = history;
}

/**
 * Form-level submit: saves the current values of every editable unit,
 * then submits all of them at once. Blocks (submitting nothing) if any
 * required field across the editable units is still empty. Accepted and
 * already-submitted units are left untouched.
 */
function handleFormSubmit() {
  if (state.role !== 'respondent' || !state.spec) return;

  const editable = instances.filter(
    (i) => statusOf(i.id) === 'draft' || statusOf(i.id) === 'revised'
  );
  if (!editable.length) {
    setNotice('Nothing to submit — all units are already submitted or accepted.');
    return;
  }

  // Save every editable unit's current values first, so submission
  // reflects what is on screen. Abort on the first parse error.
  for (const inst of editable) {
    const result = saveUnitFields(inst, '');
    if (!result.ok) {
      setNotice(`Cannot submit: ${result.field.fieldLabel} — ${result.error}`);
      document.getElementById(`input-${result.field.questionInstanceId}`).focus();
      return;
    }
  }

  // Required-field check across all editable units. For boolean fields,
  // "No" is a valid populated value; only an empty value counts as
  // unanswered.
  const missing = [];
  for (const inst of editable) {
    for (const f of inst.fields) {
      if (f.required && currentAnswer(f.questionInstanceId) === '') {
        missing.push(f);
      }
    }
  }
  if (missing.length) {
    const labels = missing.map((f) => `"${f.fieldLabel}"`).join(', ');
    setNotice(`Cannot submit: required field${missing.length === 1 ? '' : 's'} with no answer: ${labels}.`);
    const firstInput = document.getElementById(`input-${missing[0].questionInstanceId}`);
    if (firstInput) firstInput.focus();
    return;
  }

  for (const inst of editable) {
    state.statuses[inst.id] = 'submitted';
  }
  persistState();
  renderApp();
  setNotice(`Submitted ${editable.length} unit${editable.length === 1 ? '' : 's'} for review.`);
}

function handleAcceptUnit(inst) {
  if (state.role !== 'researcher') return;
  state.statuses[inst.id] = 'accepted';
  persistState();
  renderApp();
  setNotice(`"${inst.heading}" was accepted. Response fields are now read-only.`);
}

function handleReopenUnit(inst) {
  if (state.role !== 'researcher') return;
  state.statuses[inst.id] = 'submitted';
  persistState();
  renderApp();
  setNotice(`"${inst.heading}" was reopened; the respondent can edit the answers again.`);
}

/* ---------------------------------------------------------------------
   12. Version history dialog and comparison
   --------------------------------------------------------------------- */

function unitContextText(inst) {
  let text = `Review unit: ${inst.heading}`;
  if (inst.kind === 'group') text += ` (group ${inst.baseId})`;
  if (inst.NACE_code) text += ` — sector: ${inst.sector_name} (NACE ${inst.NACE_code})`;
  text += ` — status: ${STATUS_LABELS[statusOf(inst.id)]}`;
  return text;
}

function openHistoryDialog(inst, field) {
  document.getElementById('history-title').textContent = `Answer history — ${field.fieldLabel}`;
  document.getElementById('history-context').textContent = unitContextText(inst);

  const content = document.getElementById('history-content');
  clearNode(content);

  const history = state.versions[field.questionInstanceId] || [];
  if (!history.length) {
    content.append(el('p', { class: 'muted', text: 'No versions recorded for this field.' }));
  } else {
    // Reverse chronological order; the first entry is the current version.
    const ordered = [...history].reverse();
    ordered.forEach((version, index) => {
      content.append(renderVersionEntry(field, version, index === 0));
    });
  }
  document.getElementById('history-dialog').showModal();
}

function renderVersionEntry(field, version, isCurrent) {
  const entry = el('div', { class: isCurrent ? 'version-entry is-current' : 'version-entry' });

  const meta = el('p', { class: 'version-meta' });
  meta.append(el('span', { class: 'version-number', text: `Version ${version.version_number}` }));
  if (isCurrent) meta.append(el('span', { class: 'status-badge status-accepted', text: 'Current' }));
  meta.append(document.createTextNode(` · ${formatTimestamp(version.timestamp)} · ${version.author_role}`));
  entry.append(meta);

  if (version.change_note) {
    entry.append(el('p', { class: 'version-note', text: `Note: ${version.change_note}` }));
  }
  entry.append(el('p', { class: 'version-value', text: displayAnswerValue(version.value) }));

  if (isCurrent) {
    const history = state.versions[field.questionInstanceId];
    if (history.length > 1) {
      const previous = history[history.length - 2];
      const panel = el('div', { class: 'diff-panel', props: { hidden: true } });
      const toggle = el('button', { class: 'secondary small', text: 'Compare with previous version' });
      toggle.addEventListener('click', () => {
        panel.hidden = !panel.hidden;
        toggle.textContent = panel.hidden ? 'Compare with previous version' : 'Hide comparison';
        if (!panel.hidden && !panel.firstChild) {
          renderVersionComparison(panel, field, previous, version);
        }
      });
      entry.append(toggle, panel);
    } else {
      entry.append(el('p', { class: 'muted small', text: 'No previous version to compare with.' }));
    }
  }
  return entry;
}

/**
 * Comparison view: word-level diff for text/longtext, previous vs
 * current side by side for number, boolean and date fields.
 */
function renderVersionComparison(container, field, previous, current) {
  container.append(el('h4', {
    class: 'diff-heading',
    text: `Comparison: version ${previous.version_number} → version ${current.version_number}`,
  }));

  if (field.questionType === 'text' || field.questionType === 'longtext') {
    container.append(renderWordDiff(String(previous.value ?? ''), String(current.value ?? '')));
    container.append(el('p', {
      class: 'diff-legend',
      text: 'Removed words are struck through; added words are highlighted.',
    }));
  } else {
    container.append(
      el('p', { class: 'diff-text', text: `Previous (version ${previous.version_number}): ${displayAnswerValue(previous.value)}` }),
      el('p', { class: 'diff-text', text: `Current (version ${current.version_number}): ${displayAnswerValue(current.value)}` })
    );
  }
}

/** Word-level diff using a longest-common-subsequence table. */
function wordDiff(oldText, newText) {
  const a = oldText.split(/\s+/).filter(Boolean);
  const b = newText.split(/\s+/).filter(Boolean);
  const n = a.length;
  const m = b.length;

  // lcs[i][j] = LCS length of a[i..] and b[j..]. Prototype-sized inputs,
  // so the O(n*m) table is acceptable.
  const lcs = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i][j] = a[i] === b[j]
        ? lcs[i + 1][j + 1] + 1
        : Math.max(lcs[i + 1][j], lcs[i][j + 1]);
    }
  }

  const parts = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      parts.push({ type: 'same', text: a[i] });
      i++;
      j++;
    } else if (lcs[i + 1][j] >= lcs[i][j + 1]) {
      parts.push({ type: 'removed', text: a[i] });
      i++;
    } else {
      parts.push({ type: 'added', text: b[j] });
      j++;
    }
  }
  while (i < n) parts.push({ type: 'removed', text: a[i++] });
  while (j < m) parts.push({ type: 'added', text: b[j++] });
  return parts;
}

/** Merge consecutive parts of the same type into readable runs. */
function mergeDiffRuns(parts) {
  const runs = [];
  for (const part of parts) {
    const last = runs[runs.length - 1];
    if (last && last.type === part.type) {
      last.text += ' ' + part.text;
    } else {
      runs.push({ type: part.type, text: part.text });
    }
  }
  return runs;
}

/** Render the word diff safely (spans via textContent, no innerHTML). */
function renderWordDiff(oldText, newText) {
  const paragraph = el('p', { class: 'diff-text' });
  const runs = mergeDiffRuns(wordDiff(oldText, newText));
  runs.forEach((run, index) => {
    if (index) paragraph.append(' ');
    if (run.type === 'same') {
      paragraph.append(document.createTextNode(run.text));
    } else {
      paragraph.append(el('mark', {
        class: run.type === 'removed' ? 'diff-removed' : 'diff-added',
        text: run.text,
      }));
    }
  });
  return paragraph;
}

/* ---------------------------------------------------------------------
   13. Comments
   --------------------------------------------------------------------- */

let commentDialogContext = null; // { inst, mode: 'comment' | 'revision', anchor: null | {...} }

function openCommentDialog(inst, mode, anchor) {
  commentDialogContext = { inst, mode, anchor: anchor || null };
  const isRevision = mode === 'revision';

  document.getElementById('comment-title').textContent = isRevision ? 'Request revision' : 'Add comment';
  document.getElementById('comment-context').textContent =
    `${unitContextText(inst)}` +
    (isRevision ? ' — the status changes to "Revision requested"; a comment explaining the change is required.' : '');
  document.getElementById('comment-submit').textContent = isRevision ? 'Request revision' : 'Save comment';

  // Quoted-text box: shown only when the comment is anchored to a
  // selection of the answer text.
  const anchorBox = document.getElementById('comment-anchor');
  const anchorValue = document.getElementById('comment-anchor-text');
  if (anchor) {
    anchorValue.textContent = anchor.text;
    anchorBox.hidden = false;
  } else {
    anchorBox.hidden = true;
  }

  const textarea = document.getElementById('comment-text');
  textarea.value = '';
  document.getElementById('comment-error').hidden = true;
  document.getElementById('comment-dialog').showModal();
  textarea.focus();
}

/**
 * Scroll to a longtext field and select the span a comment is anchored to,
 * with a brief highlight flash. Offsets are character positions into the
 * current value; if the answer has since changed the span may not match
 * the original quote, but the selection still points to the same region.
 */
function locateAnchor(button) {
  const fieldId = button.dataset.fieldId;
  const input = document.getElementById(`input-${fieldId}`);
  if (!input) return;
  const start = parseInt(button.dataset.anchorStart, 10);
  const end = parseInt(button.dataset.anchorEnd, 10);
  input.focus();
  input.setSelectionRange(start, end);
  input.scrollIntoView({ behavior: 'smooth', block: 'center' });
  // Flash the field wrapper briefly to draw the eye.
  const wrapper = input.closest('.field');
  if (wrapper) {
    wrapper.classList.add('anchor-flash');
    setTimeout(() => wrapper.classList.remove('anchor-flash'), 1500);
  }
}

/**
 * Open the comment dialog anchored to the text currently selected in a
 * longtext answer. Reads the textarea's selection offsets at click time;
 * if nothing is selected, tells the user to select some text first.
 */
function openCommentOnSelection(inst, fieldId) {
  const input = document.getElementById(`input-${fieldId}`);
  if (!input) return;
  const start = input.selectionStart;
  const end = input.selectionEnd;
  if (start === null || end === null || start === end) {
    setNotice('Select some of the answer text first, then click "Comment on selection".');
    return;
  }
  const value = input.value;
  const anchor = {
    fieldId,
    start,
    end,
    text: value.slice(start, end),
  };
  openCommentDialog(inst, 'comment', anchor);
}

/**
 * Append an immutable comment. Comments are stored in a separate
 * structure from answer versions: adding a comment never creates an
 * answer version.
 */
function addComment(reviewUnitId, text, anchor) {
  const thread = state.comments[reviewUnitId] || [];
  const comment = {
    comment_id: `${reviewUnitId}__c${thread.length + 1}__${Date.now()}`,
    review_unit_id: reviewUnitId,
    author_role: ROLES[state.role],
    timestamp: new Date().toISOString(),
    comment_text: text,
  };
  // Optional anchor: ties the comment to a span of a longtext answer.
  // Absent (or null) fields mean a whole-question comment.
  if (anchor) {
    comment.field_id = anchor.fieldId;
    comment.anchor_start = anchor.start;
    comment.anchor_end = anchor.end;
    comment.anchor_text = anchor.text;
  }
  thread.push(comment);
  state.comments[reviewUnitId] = thread;
}

function handleCommentSubmit(event) {
  event.preventDefault();
  const textarea = document.getElementById('comment-text');
  const text = textarea.value.trim();
  const error = document.getElementById('comment-error');
  if (!text) {
    error.textContent = commentDialogContext.mode === 'revision'
      ? 'A comment explaining the requested revision is required.'
      : 'Please enter a comment.';
    error.hidden = false;
    textarea.focus();
    return;
  }

  const { inst, mode, anchor } = commentDialogContext;
  addComment(inst.id, text, anchor);
  if (mode === 'revision') {
    // Status change only — never creates an answer version.
    state.statuses[inst.id] = 'revision_requested';
  }
  persistState();
  document.getElementById('comment-dialog').close();
  renderApp();
  setNotice(mode === 'revision'
    ? `Revision requested for "${inst.heading}".`
    : 'Comment added.');
}

/* ---------------------------------------------------------------------
   14. Import, export and resets
   --------------------------------------------------------------------- */

function buildSpecFromParsed(parsed, fileName, source) {
  return {
    meta: {
      fileName,
      source,
      importedAt: new Date().toISOString(),
      fingerprint: computeFingerprint({
        questions: parsed.questions,
        groups: parsed.groups,
        sectors: parsed.sectors,
      }),
    },
    questions: parsed.questions,
    groups: parsed.groups,
    sectors: parsed.sectors,
  };
}

/**
 * Show the validation report and workbook summary. Nothing is applied
 * until the user clicks "Use this specification"; if critical errors
 * exist, the questionnaire is never updated.
 */
function showImportDialog(fileName, parsed, report, spec) {
  pendingImportSpec = spec;
  const content = document.getElementById('import-content');
  clearNode(content);

  content.append(el('p', { class: 'import-file', text: `Workbook: ${fileName}` }));

  const counts = el('ul', { class: 'import-counts' }, [
    el('li', { text: `Question definitions: ${parsed.questions.length}` }),
    el('li', { text: `Question groups: ${parsed.groups.length}` }),
    el('li', { text: `Sectors: ${parsed.sectors.length}` }),
  ]);
  if (spec) {
    counts.append(el('li', { text: `Generated review units: ${generateInstances(spec).length}` }));
  }
  content.append(counts);

  if (report.errors.length) {
    content.append(el('h3', { class: 'report-heading', text: `Errors (${report.errors.length}) — the questionnaire was NOT updated` }));
    const list = el('ul', { class: 'report-list errors' });
    for (const message of report.errors) list.append(el('li', { text: message }));
    content.append(list);
  }
  if (report.warnings.length) {
    content.append(el('h3', { class: 'report-heading', text: `Warnings (${report.warnings.length}) — rendering is possible` }));
    const list = el('ul', { class: 'report-list warnings' });
    for (const message of report.warnings) list.append(el('li', { text: message }));
    content.append(list);
  }

  const applyButton = document.getElementById('import-apply');
  if (spec && !report.errors.length) {
    applyButton.hidden = false;
    if (hasExistingResponseData()) {
      content.append(el('p', {
        class: 'import-warning',
        text: 'Response data already exists. Applying this specification will clear all answers, versions, comments and statuses.',
      }));
    }
  } else {
    applyButton.hidden = true;
  }

  document.getElementById('import-dialog').showModal();
}

function applyImportedSpec() {
  if (!pendingImportSpec) return;
  if (hasExistingResponseData() && !window.confirm(
    'Response data (answers, versions, comments, statuses) already exists. ' +
    'Applying a new specification will reset it. Continue?'
  )) {
    return;
  }
  document.getElementById('import-dialog').close();
  state.spec = pendingImportSpec;
  pendingImportSpec = null;
  clearResponseData();
  persistState();
  renderApp();
  setNotice(`Imported specification "${state.spec.meta.fileName}".`);
}

function handleFileImport(event) {
  const file = event.target.files && event.target.files[0];
  event.target.value = ''; // allow selecting the same file again later
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    setNotice('The SheetJS library is not loaded. An internet connection is required when using the CDN script — see README.md for an offline alternative.');
    return;
  }

  file.arrayBuffer()
    .then((buffer) => {
      const parsed = parseWorkbook(buffer);
      const report = validateParsedSpec(parsed);
      // A specification is only built when validation found no errors.
      const spec = report.errors.length
        ? null
        : buildSpecFromParsed(parsed, file.name, 'excel');
      showImportDialog(file.name, parsed, report, spec);
    })
    .catch((err) => {
      console.error(err);
      setNotice(`Could not read "${file.name}": ${err.message}`);
    });
}

function loadSample() {
  if (hasExistingResponseData() || state.spec) {
    if (!window.confirm('Replace the current specification and all responses with the built-in sample (demonstration data)?')) {
      return;
    }
  }
  state = EMPTY_STATE();
  state.spec = buildSampleSpec();
  seedSampleResponses();
  persistState();
  renderApp();
  setNotice('Loaded the built-in sample specification (demonstration data).');
}

function reloadSavedData() {
  state = loadState();
  statusFilter = 'all';
  document.getElementById('status-filter').value = 'all';
  renderApp();
  setNotice('Reloaded the saved specification and data from browser storage.');
}

function resetResponses() {
  if (!state.spec) {
    setNotice('No specification is loaded; there is nothing to reset.');
    return;
  }
  if (!window.confirm('Reset all answers, versions, comments and statuses? The specification is kept.')) {
    return;
  }
  clearResponseData();
  persistState();
  renderApp();
  setNotice('Responses and review history were reset; the specification was kept.');
}

function resetAllData() {
  if (!window.confirm('Reset ALL prototype data, including the imported specification? This cannot be undone.')) {
    return;
  }
  state = EMPTY_STATE();
  persistState();
  renderApp();
  setNotice('All prototype data was reset.');
}

/** Export everything (spec, instances, answers, versions, comments,
    statuses) as a JSON download. */
function exportAllData() {
  const payload = {
    export_timestamp: new Date().toISOString(),
    specification: state.spec,
    generated_instances: state.spec ? generateInstances(state.spec) : [],
    current_answers: state.answers,
    answer_versions: state.versions,
    comments: state.comments,
    review_statuses: state.statuses,
    active_role: state.role,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'questionnaire-prototype-data.json';
  link.click();
  URL.revokeObjectURL(url);
  setNotice('Exported all prototype data as JSON.');
}

/* ---------------------------------------------------------------------
   15. Misc UI helpers
   ------------------------------------------------------------------- */

function setNotice(text) {
  document.getElementById('app-notice').textContent = text;
}

function populateStatusFilter() {
  const select = document.getElementById('status-filter');
  clearNode(select);
  select.append(el('option', { props: { value: 'all' }, text: 'All statuses' }));
  for (const status of STATUS_ORDER) {
    select.append(el('option', { props: { value: status }, text: STATUS_LABELS[status] }));
  }
  select.value = statusFilter;
}

function updateToolbarState() {
  document.getElementById('btn-reload').disabled = !state.spec;
  document.getElementById('btn-reset-responses').disabled = !state.spec;
}

function setRole(role) {
  state.role = role;
  persistState();
  renderApp();
  setNotice(`Active role: ${ROLES[role]}.`);
}

/* ---------------------------------------------------------------------
   16. Event wiring and initialisation
   ------------------------------------------------------------------- */

function handleQuestionnaireClick(event) {
  const button = event.target.closest('button[data-action]');
  if (!button) return;
  const inst = instances.find((i) => i.id === button.dataset.unitId);
  if (!inst) return;

  switch (button.dataset.action) {
    case 'save': handleSaveUnit(inst); break;
    case 'request-revision': openCommentDialog(inst, 'revision'); break;
    case 'accept': handleAcceptUnit(inst); break;
    case 'reopen': handleReopenUnit(inst); break;
    case 'comment': openCommentDialog(inst, 'comment'); break;
    case 'comment-selection': openCommentOnSelection(inst, button.dataset.fieldId); break;
    case 'locate-anchor': locateAnchor(button); break;
    case 'toggle-expand': {
      // Expand/collapse the longtext box in place — no re-render, so the
      // typing focus and scroll position are preserved.
      const fieldId = button.dataset.fieldId;
      const input = document.getElementById(`input-${fieldId}`);
      const expand = !expandedLongtextFields.has(fieldId);
      if (expand) expandedLongtextFields.add(fieldId);
      else expandedLongtextFields.delete(fieldId);
      input.classList.toggle('expanded', expand);
      button.textContent = expand ? 'Collapse' : 'Expand';
      button.setAttribute('aria-expanded', String(expand));
      if (expand) input.focus();
      break;
    }
    case 'history': {
      const field = inst.fields.find((f) => f.questionInstanceId === button.dataset.fieldId);
      if (field) openHistoryDialog(inst, field);
      break;
    }
  }
}

function init() {
  state = loadState();
  instances = state.spec ? generateInstances(state.spec) : [];
  populateStatusFilter();

  document.getElementById('role-respondent').addEventListener('click', () => setRole('respondent'));
  document.getElementById('role-researcher').addEventListener('click', () => setRole('researcher'));
  document.getElementById('file-input').addEventListener('change', handleFileImport);
  document.getElementById('btn-load-sample').addEventListener('click', loadSample);
  document.getElementById('btn-reload').addEventListener('click', reloadSavedData);
  document.getElementById('btn-export').addEventListener('click', exportAllData);
  document.getElementById('btn-reset-responses').addEventListener('click', resetResponses);
  document.getElementById('btn-reset-all').addEventListener('click', resetAllData);
  document.getElementById('btn-form-submit').addEventListener('click', handleFormSubmit);

  document.getElementById('status-filter').addEventListener('change', (event) => {
    statusFilter = event.target.value;
    renderQuestionnaire();
  });

  document.getElementById('questionnaire').addEventListener('click', handleQuestionnaireClick);

  document.getElementById('history-close').addEventListener('click', () => {
    document.getElementById('history-dialog').close();
  });
  document.getElementById('comment-cancel').addEventListener('click', () => {
    document.getElementById('comment-dialog').close();
  });
  document.getElementById('comment-form').addEventListener('submit', handleCommentSubmit);
  document.getElementById('import-close').addEventListener('click', () => {
    pendingImportSpec = null;
    document.getElementById('import-dialog').close();
  });
  document.getElementById('import-apply').addEventListener('click', applyImportedSpec);

  renderApp();
}

init();

/* APEX Dashboard - PREPAY tab configuration */
window.APEX = window.APEX || {};

APEX.prepayConfig = Object.freeze({
  /* First file found is used ("Pepay.xlsx" supported as the current file name) */
  files: ['Prepay.xlsx', 'Pepay.xlsx'],
  idbKey: 'prepayHandle',
  pollInterval: 15000,

  /* October snapshot sheets: 1001, 1002 … 1031 (also tolerates 10-01, 10_01, 10.01) */
  sheetPattern: /^\s*10[\s\-_.]?(0[1-9]|[12]\d|3[01])\s*$/,
  snapshotMonth: 10,
  /* Explicitly excluded sheets (all other non-October sheets are skipped automatically) */
  excludeSheets: ['Sheet1', 'Sheet2', 'JIRA', 'FOR JIRA'],
  /* JIRA Analytics source sheet (first match wins, case-insensitive) */
  jiraSheets: ['FOR JIRA', 'JIRA'],

  blankLabel: '(Blank)',
  highUrgency: 90,              // Urgency >= this value is "High Urgency"
  topN: 15,                     // bars shown in WGS chart

  columns: {
    dcn: ['DCN'],
    subStatus: ['Pega sub-status', 'Pega Sub-status', 'Pega Substatus'],
    urgency: ['Urgency'],
    skill: ['Claim specialization skill'],
    assignedTo: ['Assigned to'],
    lastResearch: ['Last research note on'],
    lastNote: ['Last note'],
    updates: ['Updates', 'UPDATES'],
    jiraDate: ['JIRA request date submitted'],
    jiraUpdate: ['JIRA update'],
    verified: ['VERIFIED', 'Verified'],
    comment: ['Comment'],
    wgs: ['WGS Status Code'],
    dmAge: ['DM Age'],
    lastTouched: ['Days Since Last Touched In CIW'],
    algorithm: ['Algorithm'],
    sccf: ['SCCF'],
    ticket: ['Ticket number'],
    penalty: ['Penalty amount']
  },
  requiredColumns: ['dcn', 'assignedTo'],
  numericColumns: ['urgency', 'dmAge', 'lastTouched'],
  dateColumns: ['lastResearch'],
  /* Formula-driven sheets sometimes return 0 for empty text cells */
  zeroAsBlank: ['wgs', 'algorithm', 'sccf', 'ticket', 'penalty', 'subStatus', 'comment', 'lastNote'],

  dmBuckets: [[0, 2, '0–2 days'], [3, 5, '3–5 days'], [6, 10, '6–10 days'], [11, 15, '11–15 days'], [16, 30, '16–30 days'], [31, Infinity, '31+ days']],
  touchBuckets: [[0, 2, '0–2 days'], [3, 5, '3–5 days'], [6, 10, '6–10 days'], [11, Infinity, '11+ days']],

  filters: [
    { key: 'assignedTo', label: 'Assigned To' },
    { key: 'subStatus', label: 'Pega Sub-status' },
    { key: 'urgency', label: 'Urgency' },
    { key: 'skill', label: 'Claim Skill' },
    { key: 'wgs', label: 'WGS Status Code' },
    { key: 'algorithm', label: 'Algorithm' },
    { key: 'snapshot', label: 'Snapshot Date' }
  ],

  storageKeys: { report: 'apex.prepay.report' }
});

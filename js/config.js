/* APEX Dashboard - central configuration */
window.APEX = window.APEX || {};

APEX.config = Object.freeze({
  workbook: 'DI Ageing-Pends Dashboard.xlsx',
  sheet: 'Update Here Please',

  /* How often (ms) to check the workbook for changes */
  pollInterval: 15000,

  blankLabel: '(Blank)',

  categories: { pends: 'pends', ageing: 'ageing' },
  statuses: { completed: 'completed', pending: 'pending' },

  /* Logical field -> accepted worksheet header names (first match wins, case-insensitive) */
  columns: {
    status: ['Status'],
    category: ['Category'],
    manager: ['Manager'],
    subStatus: ['Sub-status', 'Sub Status', 'Substatus'],
    date: ['Date'],
    inquiry: ['Inquiry number2', 'Inquiry number', 'Inquiry #'],
    status2: ['Status2'],
    age: ['Age'],
    pendedDays: ['Pended Days'],
    pendedDate: ['Pended Date'],
    domain: ['Domain'],
    assignedTo: ['Assigned To'],
    workgroup: ['Workgroup'],
    provider: ['Provider Name'],
    remarks: ['Remarks']
  },
  requiredColumns: ['status', 'category', 'manager', 'subStatus'],

  /* Columns shown in the Detail Records table */
  detailColumns: [
    { key: 'inquiry', title: 'Inquiry #' },
    { key: 'category', title: 'Category', type: 'category' },
    { key: 'status', title: 'Status', type: 'status' },
    { key: 'status2', title: 'Status Detail' },
    { key: 'subStatus', title: 'Sub-status' },
    { key: 'manager', title: 'Manager' },
    { key: 'date', title: 'Date', type: 'date' },
    { key: 'age', title: 'Age Bucket' },
    { key: 'pendedDays', title: 'Pended Days', type: 'number' },
    { key: 'pendedDate', title: 'Pended Date' },
    { key: 'domain', title: 'Domain' },
    { key: 'assignedTo', title: 'Assigned To' },
    { key: 'workgroup', title: 'Workgroup' },
    { key: 'provider', title: 'Provider Name', type: 'long' },
    { key: 'remarks', title: 'Remarks', type: 'long' }
  ],

  filterKeys: ['category', 'manager', 'subStatus'],

  storageKeys: { mode: 'apex.mode', accent: 'apex.accent', tab: 'apex.tab' },

  themeIcons: { light: 'pi black.png', dark: 'pi white.png' }
});

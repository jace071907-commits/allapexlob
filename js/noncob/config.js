/* APEX Dashboard - NON COB tab configuration */
window.APEX = window.APEX || {};

APEX.noncobConfig = Object.freeze({
  files: ['NON COB.xlsx', 'Non COB.xlsx', 'NONCOB.xlsx'],
  idbKey: 'noncobHandle',
  pollInterval: 15000,
  blankLabel: '(Blank)',

  /* Only these sheets are read – "Summary" and "Pivot" are ignored */
  sheets: { ageing: 'Ageing', rstatus: 'R Status', opid: 'OPID' },

  /* Logical field -> worksheet header aliases, per sheet */
  columns: {
    ageing: {
      weekEnding: ['Week Ending'], associate: ['USR_NAME'], associateId: ['ASSIGNED_USR_ID'], team: ['TEAM'], manager: ['MGR_ID'],
      status: ['STATUS_CODE_DESC'], recoveryType: ['CLM_RECOV_TYPE'], state: ['MCS_STATE'], provState: ['PROV_STATE'], mbu: ['MBU_CODE'],
      currAge: ['CURR_AGE'], currAgeCat: ['CURR_AGE_CAT'], pendAge: ['PEND_AGE_CURRENT'], pendAgeCat: ['PEND_AGE_CURRENT_CAT'],
      pendCode: ['PEND_CODE'], pendDesc: ['PEND_CODE_DESC'], claimId: ['CLM_ID'], memberId: ['MBR_ID'], errNum: ['ERR_NUM'],
      provName: ['PROV_NAME'], grpName: ['GRP_NAME'], paid: ['CLM_AMT_PD_SMRY'], project: ['PROJECT_NAME']
    },
    rstatus: {
      weekEnding: ['Week Ending'], associate: ['Assigned Usr Name'], associateId: ['Assigned Usr Id'], team: ['TEAM'], manager: ['Mgr Name'],
      rStatus: ['Financial Event Description'], recoveryType: ['Clm Recov Type'], state: ['Svc Prov St'], mbu: ['Mbu Code'],
      claimId: ['Clm Id'], errNum: ['Err Num'], rStatusDate: ['R Status Date'], adjReason: ['Adjstmnt Rsn'], srcAdjReason: ['Src Adjstmnt Rsn'],
      paid: ['Clm Amt Pd Smry'], grpName: ['Grp Name'], provName: ['Prov Lname']
    },
    opid: {
      weekEnding: ['WeekEnding', 'Week Ending'], associate: ['NAME'], domain: ['DOMAIN'], team: ['TEAM'], processor: ['Processor'],
      manager: ['Manager'], mktManager: ['Mkt Manager'], carryOver: ['Carry Over'], state: ['MBU State'], mbu: ['MBU Code'],
      claimNumber: ['Claim Number'], ageDays: ['Age In Days Total INT'], ageBucket: ['Claim WFM Total'], pendException: ['Pend Exception Level 3'],
      queueStatus: ['Queue Status Code'], pendCode: ['Pend Code 01'], lastProcessed: ['Date Last Processed'], area: ['Area']
    }
  },
  dateFields: ['weekEnding', 'rStatusDate', 'lastProcessed'],
  numFields: ['currAge', 'pendAge', 'ageDays', 'paid'],
  /* OPID stores names as "Last,First" – flipped to "First Last" to match the other sheets */
  flipNameFields: { opid: ['associate', 'processor', 'manager', 'mktManager'] },

  ageBuckets: [[0, 7, '0-7 Days'], [8, 15, '8-15 Days'], [16, 30, '16-30 Days'], [31, 45, '31-45 Days'], [46, 60, '46-60 Days'], [61, 90, '61-90 Days'], [91, Infinity, '90+ Days']],
  opidAgeOrder: ['00-03', '04-05', '06-10', '11-14', '15-25', '26-30', '31-60', '61-90', '91+'],

  /* Global filters – each applies only to sheets that contain the field */
  filters: [
    { key: 'associate', label: 'Associate' },
    { key: 'processor', label: 'Processor' },
    { key: 'manager', label: 'Manager' },
    { key: 'team', label: 'Team' },
    { key: 'status', label: 'Status' },
    { key: 'rStatus', label: 'R Status' },
    { key: 'carryOver', label: 'Carry Over' },
    { key: 'recoveryType', label: 'Recovery Type' },
    { key: 'state', label: 'State' },
    { key: 'mbu', label: 'MBU' }
  ],

  /* Pend-code KPIs (these values live in the Ageing sheet's PEND_CODE_DESC column) */
  pendKpis: [
    { key: 'rac', match: /rac support intake/i },
    { key: 'restoration', match: /claim restoration/i },
    { key: 'restricted', match: /restricted records/i },
    { key: 'hostPlan', match: /waiting on host plan/i }
  ],

  topN: 20,
  storageKeys: { report: 'apex.noncob.report' }
});

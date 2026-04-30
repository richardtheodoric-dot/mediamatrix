// ================================================================
// MEDIA MATRIX — Agency Dashboard Backend v5
// Google Apps Script — paste di Extensions > Apps Script
// Deploy: Web App, Execute as: Me, Who has access: Anyone
// ================================================================

const SS_ID = '1gh_lVYMByHE5V0Ncv4Qxf4_Cj8OBAN_dPLG2xRNr2GU';
const SS_URL = 'https://docs.google.com/spreadsheets/d/' + SS_ID + '/edit';
const CONTENT_TIMELINE_SS_ID = '1NwU6llMjEQew0oxa1nGsVXaDGUEZxx5BTVXAmfKj54I';
const CONTENT_TIMELINE_SS_URL = 'https://docs.google.com/spreadsheets/d/' + CONTENT_TIMELINE_SS_ID + '/edit';
const CONTENT_TIMELINE_TARGET_GID = 1291951677;
const ADS_REPORT_SS_ID = '1OMCIZSEtISHRpLlrNdH8Rbz9KWVV4YNwDEtzKzkoejg';
const ADS_REPORT_SS_URL = 'https://docs.google.com/spreadsheets/d/' + ADS_REPORT_SS_ID + '/edit';
const OWNER_CODE = 'MMOWNER2024'; // Ganti sesuai keinginan
const ADS_DOC_ID = '1ELCKAH6xenWRyvoAqk9dBDmkIq6QweRJEu-6Hz_3FOo';
const ADS_DOC_URL = 'https://docs.google.com/document/d/' + ADS_DOC_ID + '/edit';
let LAST_DOC_ERROR = '';

// ── ROUTING ──────────────────────────────────────────────────────
function doGet(e) {
  const a = e.parameter.action;
  if (a === 'getClients')      return j(getClients());
  if (a === 'getPeriodData')   return j(getPeriodData(e.parameter.periodKey));
  if (a === 'getAllCalendars') return j(getAllCalendars());
  if (a === 'getAllCampaignReports') return j(getAllCampaignReports());
  if (a === 'getAllBatches')   return j(getAllBatches());
  if (a === 'getAllCreatives') return j(getAllCreatives());
  if (a === 'getUsers')        return j(getUsers());
  if (a === 'verifyOwner')     return j({ valid: e.parameter.code === OWNER_CODE });
  return j({ error: 'Unknown GET action: ' + a });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const a = b.action;
  if (a === 'saveClients')          return j(saveClients(b.clients));
  if (a === 'savePeriodData')       return j(savePeriodData(b.periodKey, b.clientId, b.data, b.user));
  if (a === 'saveCalendar')         return j(saveCalendar(b.clientId, b.posts, b.user));
  if (a === 'saveCampaignReport')   return j(saveCampaignReport(b.periodKey, b.clientId, b.data, b.user));
  if (a === 'saveBatch')            return j(saveBatch(b.clientId, b.batch, b.user));
  if (a === 'deleteBatch')          return j(deleteBatch(b.clientId, b.batchId));
  if (a === 'saveCreative')         return j(saveCreative(b.clientId, b.creative, b.user));
  if (a === 'deleteCreative')       return j(deleteCreative(b.clientId, b.creativeId));
  if (a === 'saveClientInfo')       return j(saveClientInfo(b.clientId, b.info));
  if (a === 'registerUser')         return j(registerUser(b.email, b.name, b.role));
  if (a === 'exportAllToSheets')    return j(exportAllToSheets(b.periodKey));
  if (a === 'exportCalendarToSheet') return j(exportCalendarToSheet(b.clientId));
  if (a === 'sendManualReminder')   return j(sendManualReminder(b.emails, b.type, b.message));
  return j({ error: 'Unknown POST action: ' + a });
}

function j(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SHEET HELPER ─────────────────────────────────────────────────
function sh(name, headers) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName(name);
  let created = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
    if (headers) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,headers.length)
        .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');
    }
  }
  if (created && headers) styleRows(sheet);
  return sheet;
}

function styleRows(sheet) {
  try {
    const last = sheet.getLastRow();
    const cols = Math.max(sheet.getLastColumn(), 1);
    sheet.setFrozenRows(1);
    sheet.setHiddenGridlines(true);
    sheet.getRange(1,1,Math.max(last,1),cols)
      .setFontFamily('Arial')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.getRange(1,1,1,cols)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
    if (last < 2) {
      applyColumnWidths_(sheet);
      return;
    }
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(1,1,last,cols).createFilter();
    sheet.getRange(2,1,last-1,cols)
      .setBackground('#ffffff')
      .setWrap(true);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    applyColumnWidths_(sheet);
  } catch(e) {}
}

function applyColumnWidths_(sheet) {
  const name = sheet.getName();
  const widthsBySheet = {
    Users: [240,180,150,180,180],
    Clients: [165,220,280,120,430,180,150,150,280,120],
    PeriodData: [130,165,220,360,340,340,280,170,180],
    ContentCalendar: [165,220,170,120,120,150,420,120,130,170,180],
    CampaignReports: [130,165,220,260,520,160,180],
    Batches: [165,220,170,220,130,130,430,360,180,170],
    Creatives: [165,220,170,220,120,130,330,150,150,120,130,360,170]
  };
  const widths = widthsBySheet[name];
  if (!widths) return;
  widths.forEach((w,i) => sheet.setColumnWidth(i+1,w));
}

// ── USERS ─────────────────────────────────────────────────────────
function registerUser(email, name, role) {
  const sheet = sh('Users', ['email','name','role','registeredAt','lastLogin']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === email) {
      sheet.getRange(i+1,2,1,3).setValues([[name, role, new Date().toISOString()]]);
      styleRows(sheet);
      return { success: true, existing: true };
    }
  }
  sheet.appendRow([email, name, role, new Date().toISOString(), new Date().toISOString()]);
  styleRows(sheet);
  return { success: true, existing: false };
}

function getUsers() {
  const sheet = sh('Users', ['email','name','role','registeredAt','lastLogin']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { users: [] };
  return { users: rows.slice(1).map(r => ({ email:r[0], name:r[1], role:r[2] })) };
}

function getUsersByRole(role) {
  return getUsers().users.filter(u => u.role === role);
}

// ── CLIENTS ───────────────────────────────────────────────────────
function getClients() {
  const sheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { clients: [] };
  return {
    clients: rows.slice(1).map(r => ({
      id:r[0], name:r[1],
      platforms: JSON.parse(r[2]||'[]'),
      budget:r[3], docsUrl:r[4], createdAt:r[5],
      pic:r[6]||'', niche:r[7]||'', notes:r[8]||'', contractMonths:r[9]||''
    }))
  };
}

function saveClients(clients) {
  const sheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const headers = ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths'];
  const existingRows = sheet.getDataRange().getValues();
  const createdAtById = {};
  for (let i = 1; i < existingRows.length; i++) {
    if (existingRows[i][0]) createdAtById[existingRows[i][0]] = existingRows[i][5] || new Date().toISOString();
  }
  const rows = (clients || []).map(c => {
    const docsUrl = c.docsUrl || ADS_DOC_URL;
    c.docsUrl = docsUrl;
    return [
      c.id, c.name, JSON.stringify(c.platforms||[]),
      c.budget||'', docsUrl, c.createdAt || createdAtById[c.id] || new Date().toISOString(),
      c.pic||'', c.niche||'', c.notes||'', c.contractMonths||''
    ];
  });
  sheet.clearContents();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
  styleRows(sheet);
  return { success: true, clients };
}

function saveClientInfo(clientId, info) {
  const sheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === clientId) {
      sheet.getRange(i+1,7,1,3).setValues([[info.pic||'', info.niche||'', info.notes||'']]);
      styleRows(sheet);
      return { success: true };
    }
  }
  return { success: false };
}

// ── GOOGLE DOCS 5W1H ─────────────────────────────────────────────
function createClientDoc(clientName, clientId) {
  try {
    LAST_DOC_ERROR = '';
    const doc = DocumentApp.create(clientName + '_Documentation');
    const body = doc.getBody();

    const h = body.appendParagraph('📋 ' + clientName + ' — Client Documentation');
    h.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    h.editAsText().setFontSize(20).setBold(true);
    body.appendParagraph('ID: ' + clientId + ' · Dibuat: ' + new Date().toLocaleDateString('id-ID'))
      .editAsText().setForegroundColor('#888888').setFontSize(10);
    body.appendHorizontalRule();

    // Ads documentation log
    body.appendParagraph('ADS DOCUMENTATION LOG')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Setiap kali Ads form disimpan, laporan periode akan otomatis ditambahkan ke dokumen ini.')
      .editAsText().setForegroundColor('#555555').setFontSize(11).setItalic(true);
    body.appendHorizontalRule();

    // Action Plan Checklist
    body.appendParagraph('✅ Efektivitas Action Plan')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('Setiap action plan yang dibuat dan hasil follow-upnya tercatat otomatis di sini setiap periode.')
      .editAsText().setForegroundColor('#555555').setFontSize(11).setItalic(true);
    body.appendHorizontalRule();

    // Pertanyaan Strategis
    body.appendParagraph('📝 Pertanyaan Strategis (Diisi Manual)')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const questions = [
      'WHO lainnya — Siapa PIC pengambil keputusan di sisi klien?',
      'WHAT lebih dalam — Apa unique selling point klien yang paling resonan dengan audiens?',
      'WHERE tambahan — Apakah ada channel lain selain yang kita kelola?',
      'WHEN seasonal — Apakah ada momen/event yang mempengaruhi performa?',
      'WHY lebih dalam — Ada faktor eksternal (kompetitor, algoritma) yang berpengaruh?',
      'HOW strategis — Bagaimana feedback langsung dari klien terhadap hasil kita?',
      'Apakah target audiens masih tepat, atau perlu riset ulang?',
      'Apakah ada perubahan dari sisi klien (produk baru, promo, stok) yang belum dikomunikasikan?',
    ];

    questions.forEach((q, i) => {
      body.appendParagraph((i+1) + '. ' + q).editAsText().setFontSize(11).setBold(true);
      body.appendParagraph('Jawaban: _______________________________________________')
        .editAsText().setForegroundColor('#aaaaaa').setFontSize(11).setBold(false);
      body.appendParagraph('');
    });

    body.appendHorizontalRule();
    body.appendParagraph('— Auto-generated by Media Matrix Agency Dashboard —')
      .editAsText().setForegroundColor('#aaaaaa').setFontSize(10).setItalic(true);
    doc.saveAndClose();
    return doc.getUrl();
  } catch(e) {
    Logger.log('createClientDoc error: ' + e);
    LAST_DOC_ERROR = e.toString();
    return '';
  }
}

function appendPeriodToDoc(docsUrl, clientName, periodLabel, data, platforms, userName) {
  try {
    LAST_DOC_ERROR = '';
    const doc = DocumentApp.openById(ADS_DOC_ID);
    const body = getAdsDocBodyForClient_(doc, clientName);
    if (!body) {
      LAST_DOC_ERROR = 'Tab "' + clientName + '" tidak ditemukan di Ads Documentation Doc. Buat tab dengan nama klien yang sama persis.';
      return;
    }

    body.appendParagraph('REPORT ENTRY - ' + new Date().toLocaleString('id-ID'))
      .editAsText().setFontSize(11).setBold(true).setForegroundColor('#1a73e8');
    body.appendParagraph('TANGGAL').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph(periodLabel).editAsText().setFontSize(11).setBold(true);
    body.appendParagraph('');

    body.appendParagraph('Perubahan apa yang dilakukan?').editAsText().setFontSize(11).setBold(true);
    const changes = [];
    platforms.forEach(p => {
      const plD = (data.platforms||{})[p.platform] || {};
      changes.push(
        p.platform + ' - ' + p.kpi + ': ' +
        (plD.kpi_prev||'—') + ' -> ' + (plD.kpi_curr||'—') +
        ' | Spend: ' + (plD.spend||'—') +
        ' | Reach: ' + (plD.reach||'—') +
        ' | CTR: ' + (plD.ctr||'—') +
        ' | Status: ' + (plD.health||'—')
      );
    });
    body.appendParagraph(changes.length ? changes.join('\n') : '(belum ada data platform)')
      .editAsText().setFontSize(11).setBold(false);
    body.appendParagraph('');

    body.appendParagraph('Why?').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph(collectPlatformField_(data, platforms, 'penyebab', '(belum diisi)'))
      .editAsText().setFontSize(11).setBold(false);

    body.appendParagraph('How?').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph(collectPlatformField_(data, platforms, 'insight', '(belum diisi)'))
      .editAsText().setFontSize(11).setBold(false);

    body.appendParagraph('');
    body.appendParagraph('Results?').editAsText().setFontSize(11).setBold(true);
    body.appendParagraph(buildResultsText_(data, platforms))
      .editAsText().setFontSize(11).setBold(false);

    const plans = (data.plans||[]).filter(p=>p.text);
    if (plans.length) {
      body.appendParagraph('');
      body.appendParagraph('Rencana Periode Depan').editAsText().setFontSize(11).setBold(true);
      plans.forEach((p,i) => body.appendParagraph((i+1)+'. ['+(p.category||'General')+'] '+p.text+(p.deadline?' - Target: '+p.deadline:'')).editAsText().setFontSize(11).setBold(false));
    }

    const fus = Object.values(data.followupAnswers||{});
    if (fus.length) {
      body.appendParagraph('');
      body.appendParagraph('Hasil Follow-Up Action Plan').editAsText().setFontSize(11).setBold(true);
      fus.forEach(fu => {
        body.appendParagraph('['+(fu.done===true?'SELESAI':'BELUM')+'] '+(fu.result||'—'))
          .editAsText().setFontSize(11).setBold(false);
      });
    }

    if (data.notes) {
      body.appendParagraph('');
      body.appendParagraph('Catatan').editAsText().setFontSize(11).setBold(true);
      body.appendParagraph(data.notes).editAsText().setFontSize(11).setBold(false);
    }
    body.appendParagraph('');
    body.appendParagraph('Diisi oleh: ' + (userName||'—') + ' | Disimpan: ' + new Date().toLocaleString('id-ID'))
      .editAsText().setForegroundColor('#777777').setFontSize(9).setItalic(true);
    body.appendHorizontalRule();
    doc.saveAndClose();
  } catch(e) {
    Logger.log('appendPeriodToDoc error: ' + e);
    LAST_DOC_ERROR = e.toString();
  }
}

function getAdsDocBodyForClient_(doc, clientName) {
  const target = String(clientName || '').trim().toLowerCase();
  if (!target) return null;

  if (typeof doc.getTabs === 'function') {
    const tabs = flattenDocTabs_(doc.getTabs());
    for (let i = 0; i < tabs.length; i++) {
      const title = String(tabs[i].getTitle ? tabs[i].getTitle() : '').trim().toLowerCase();
      if (title === target) return tabs[i].asDocumentTab().getBody();
    }
  }

  return null;
}

function flattenDocTabs_(tabs) {
  let result = [];
  (tabs || []).forEach(tab => {
    result.push(tab);
    if (tab.getChildTabs) result = result.concat(flattenDocTabs_(tab.getChildTabs()));
  });
  return result;
}

function collectPlatformField_(data, platforms, field, fallback) {
  const lines = [];
  platforms.forEach(p => {
    const plD = (data.platforms||{})[p.platform] || {};
    if (plD[field]) lines.push(p.platform + ': ' + plD[field]);
  });
  return lines.length ? lines.join('\n') : fallback;
}

function buildResultsText_(data, platforms) {
  const lines = [];
  platforms.forEach(p => {
    const plD = (data.platforms||{})[p.platform] || {};
    const result = [
      p.platform,
      (p.kpi||'KPI') + ' ' + (plD.kpi_prev||'—') + ' -> ' + (plD.kpi_curr||'—'),
      'status ' + (plD.health||'—'),
      'next action: ' + (plD.nextAction||'—')
    ].join(' | ');
    lines.push(result);
  });
  return lines.length ? lines.join('\n') : '(belum ada hasil)';
}

// ── PERIOD DATA (ADS) ─────────────────────────────────────────────
function getPeriodData(periodKey) {
  const sheet = sh('PeriodData', ['periodKey','clientId','clientName','platformsData_json','plans_json','followupAnswers_json','notes','updatedBy','updatedAt']);
  const rows = sheet.getDataRange().getValues();
  const result = {};
  rows.slice(1).forEach(r => {
    if (r[0] === periodKey) {
      result[r[1]] = {
        platforms: JSON.parse(r[3]||'{}'),
        plans: JSON.parse(r[4]||'[]'),
        followupAnswers: JSON.parse(r[5]||'{}'),
        notes: r[6],
        updatedBy: r[7]
      };
    }
  });
  return { data: result };
}

function savePeriodData(periodKey, clientId, data, user) {
  const sheet = sh('PeriodData', ['periodKey','clientId','clientName','platformsData_json','plans_json','followupAnswers_json','notes','updatedBy','updatedAt']);

  // Get client info
  const cSheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const cRows = cSheet.getDataRange().getValues();
  let clientName = clientId, platforms = [], docsUrl = ADS_DOC_URL, clientRow = 0;
  for (let i = 1; i < cRows.length; i++) {
    if (cRows[i][0] === clientId) {
      clientName = cRows[i][1];
      platforms = JSON.parse(cRows[i][2]||'[]');
      docsUrl = ADS_DOC_URL;
      clientRow = i + 1;
      break;
    }
  }
  if (clientRow) cSheet.getRange(clientRow,5).setValue(ADS_DOC_URL);

  // Delete existing row for this period+client
  const all = sheet.getDataRange().getValues();
  for (let i = all.length-1; i >= 1; i--) {
    if (all[i][0] === periodKey && all[i][1] === clientId) sheet.deleteRow(i+1);
  }

  const now = new Date().toISOString();
  const clean = {...data}; delete clean._open;
  sheet.appendRow([
    periodKey, clientId, clientName,
    JSON.stringify(clean.platforms||{}),
    JSON.stringify(clean.plans||[]),
    JSON.stringify(clean.followupAnswers||{}),
    clean.notes||'', user?.name||'', now
  ]);
  styleRows(sheet);

  // Append to Google Doc
  const periodLabel = formatPeriodLabel(periodKey);
  appendPeriodToDoc(docsUrl, clientName, periodLabel, clean, platforms, user?.name||'');
  if (LAST_DOC_ERROR) {
    return { success: false, error: 'Google Docs gagal diupdate: ' + LAST_DOC_ERROR, docsUrl };
  }
  return { success: true, docsUrl };
}

function formatPeriodLabel(key) {
  const m = key.match(/^(\d{4})-(\d{2})-P([123])$/);
  if (!m) return key;
  const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const year = parseInt(m[1]), month = parseInt(m[2]), period = parseInt(m[3]);
  const lastDay = new Date(year, month, 0).getDate();
  const ranges = [[1,10],[11,20],[21,lastDay]];
  const [s,e] = ranges[period-1];
  return s+'-'+e+' '+monthNames[month-1]+' '+year;
}

// ── CONTENT CALENDAR ─────────────────────────────────────────────
function getAllCalendars() {
  const sheet = sh('ContentCalendar', ['clientId','clientName','postId','date','platform','type','caption','status','batch','updatedBy','updatedAt']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { posts: [] };
  const hasBatchColumn = String(rows[0][8] || '').toLowerCase() === 'batch';
  return { posts: rows.slice(1).map(r => ({
    clientId:r[0], clientName:r[1], postId:r[2],
    date:normalizeCalendarDate_(r[3]), platform:r[4], type:r[5],
    caption:r[6], status:r[7], batch:hasBatchColumn ? (r[8] || 'Batch 1') : 'Batch 1',
    updatedBy:hasBatchColumn ? r[9] : r[8], updatedAt:hasBatchColumn ? r[10] : r[9]
  }))};
}

function saveCalendar(clientId, posts, user) {
  const sheet = sh('ContentCalendar', ['clientId','clientName','postId','date','platform','type','caption','status','batch','updatedBy','updatedAt']);
  const headers = ['clientId','clientName','postId','date','platform','type','caption','status','batch','updatedBy','updatedAt'];
  const cSheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const cRows = cSheet.getDataRange().getValues();
  let clientName = clientId;
  for (let i=1;i<cRows.length;i++){if(cRows[i][0]===clientId){clientName=cRows[i][1];break;}}

  const all = sheet.getDataRange().getValues();
  const hasBatchColumn = all.length && String(all[0][8] || '').toLowerCase() === 'batch';
  const kept = [];
  for (let i=1;i<all.length;i++){
    if(all[i][0]===clientId) continue;
    if (hasBatchColumn) {
      kept.push(headers.map((_, idx) => all[i][idx] || ''));
    } else {
      kept.push([all[i][0],all[i][1],all[i][2],all[i][3],all[i][4],all[i][5],all[i][6],all[i][7],'Batch 1',all[i][8],all[i][9]]);
    }
  }

  const now = new Date().toISOString();
  const newRows = (posts || []).map(p => [clientId, clientName, p.postId||'', normalizeCalendarDate_(p.date)||'', p.platform||'', p.type||'', p.caption||'', p.status||'Draft', p.batch||'Batch 1', user?.name||'', now]);
  const out = [headers].concat(kept, newRows);
  sheet.clearContents();
  sheet.getRange(1,1,out.length,headers.length).setValues(out);
  styleRows(sheet);
  return { success: true };
}

// ── BATCHES ───────────────────────────────────────────────────────
function getAllBatches() {
  const sheet = sh('Batches', ['clientId','clientName','batchId','name','startDate','endDate','checklist_json','evaluation','createdAt','updatedBy']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { batches: [] };
  return { batches: rows.slice(1).map(r => ({
    clientId:r[0], clientName:r[1], id:r[2], name:r[3],
    startDate:r[4], endDate:r[5],
    checklist: JSON.parse(r[6]||'[]'),
    evaluation:r[7], createdAt:r[8], updatedBy:r[9]
  }))};
}

function saveBatch(clientId, batch, user) {
  const sheet = sh('Batches', ['clientId','clientName','batchId','name','startDate','endDate','checklist_json','evaluation','createdAt','updatedBy']);
  const cSheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const cRows = cSheet.getDataRange().getValues();
  let clientName = clientId;
  for(let i=1;i<cRows.length;i++){if(cRows[i][0]===clientId){clientName=cRows[i][1];break;}}

  const all = sheet.getDataRange().getValues();
  for(let i=all.length-1;i>=1;i--){if(all[i][0]===clientId&&all[i][2]===batch.id)sheet.deleteRow(i+1);}

  sheet.appendRow([clientId, clientName, batch.id, batch.name||'', batch.startDate||'', batch.endDate||'',
    JSON.stringify(batch.checklist||[]), batch.evaluation||'', batch.createdAt||new Date().toISOString(), user?.name||'']);
  styleRows(sheet);
  return { success: true };
}

function deleteBatch(clientId, batchId) {
  const sheet = sh('Batches', ['clientId','clientName','batchId','name','startDate','endDate','checklist_json','evaluation','createdAt','updatedBy']);
  const all = sheet.getDataRange().getValues();
  for(let i=all.length-1;i>=1;i--){if(all[i][0]===clientId&&all[i][2]===batchId)sheet.deleteRow(i+1);}
  return { success: true };
}

// ── CREATIVES ─────────────────────────────────────────────────────
function getAllCreatives() {
  const sheet = sh('Creatives', ['clientId','clientName','id','name','platform','type','link','metric_before','metric_after','result','date','notes','updatedBy']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { creatives: [] };
  return { creatives: rows.slice(1).map(r => ({
    clientId:r[0], clientName:r[1], id:r[2], name:r[3],
    platform:r[4], type:r[5], link:r[6],
    metric_before:r[7], metric_after:r[8], result:r[9],
    date:r[10], notes:r[11]
  }))};
}

function saveCreative(clientId, creative, user) {
  const sheet = sh('Creatives', ['clientId','clientName','id','name','platform','type','link','metric_before','metric_after','result','date','notes','updatedBy']);
  const cSheet = sh('Clients', ['id','name','platforms_json','budget','docsUrl','createdAt','pic','niche','notes','contractMonths']);
  const cRows = cSheet.getDataRange().getValues();
  let clientName = clientId;
  for(let i=1;i<cRows.length;i++){if(cRows[i][0]===clientId){clientName=cRows[i][1];break;}}
  sheet.appendRow([clientId, clientName, creative.id, creative.name||'', creative.platform||'', creative.type||'',
    creative.link||'', creative.metric_before||'', creative.metric_after||'', creative.result||'',
    creative.date||'', creative.notes||'', user?.name||'']);
  styleRows(sheet);
  return { success: true };
}

function deleteCreative(clientId, creativeId) {
  const sheet = sh('Creatives', ['clientId','clientName','id','name','platform','type','link','metric_before','metric_after','result','date','notes','updatedBy']);
  const all = sheet.getDataRange().getValues();
  for(let i=all.length-1;i>=1;i--){if(all[i][0]===clientId&&all[i][2]===creativeId)sheet.deleteRow(i+1);}
  return { success: true };
}

// ── EXPORT — ONE FILE, TABS PER CLIENT ───────────────────────────
function exportAllToSheets(periodKey) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const clients = getClients().clients;
    const periodLabel = formatPeriodLabel(periodKey||'');
    const allCalendars = getAllCalendars().posts;
    const allBatches = getAllBatches().batches;
    const timelineSs = SpreadsheetApp.openById(CONTENT_TIMELINE_SS_ID);
    let exported = 0;

    clients.forEach(client => {
      // ── TAB: ClientName_Ads ──
      const adsTabName = client.name.slice(0,20) + '_Ads';
      let adsSheet = ss.getSheetByName(adsTabName);
      if (!adsSheet) adsSheet = ss.insertSheet(adsTabName);
      adsSheet.clear();

      // Header
      adsSheet.appendRow(['MEDIA MATRIX — Laporan Ads: ' + client.name]);
      adsSheet.getRange(1,1).setFontSize(14).setFontWeight('bold');
      adsSheet.appendRow(['Diekspor:', new Date().toLocaleString('id-ID'), 'Periode:', periodLabel]);
      adsSheet.appendRow(['PIC Klien:', client.pic||'—', 'Niche:', client.niche||'—', 'Kontrak:', (client.contractMonths||'—')+' bulan']);
      if (client.docsUrl) adsSheet.appendRow(['Dokumentasi 5W1H:', client.docsUrl]);
      adsSheet.appendRow([]);

      // Get period data for this client
      const periodData = getPeriodData(periodKey).data[client.id];
      const platforms = client.platforms || [];

      if (periodData) {
        platforms.forEach(p => {
          const plD = (periodData.platforms||{})[p.platform] || {};
          adsSheet.appendRow(['PLATFORM: ' + p.platform + ' — KPI: ' + p.kpi]);
          adsSheet.getRange(adsSheet.getLastRow(),1).setFontWeight('bold').setBackground('#e8f4fd');
          adsSheet.appendRow([p.kpi+' Periode Lalu', plD.kpi_prev||'—', p.kpi+' Periode Ini', plD.kpi_curr||'—']);
          adsSheet.appendRow(['Total Spend', plD.spend||'—', 'Reach', plD.reach||'—', 'CTR', (plD.ctr||'—')+'%']);
          adsSheet.appendRow(['Status', plD.health||'—']);
          adsSheet.appendRow(['Penyebab Growth/Stagnant', plD.penyebab||'—']);
          adsSheet.appendRow(['Insight/Temuan', plD.insight||'—']);
          adsSheet.appendRow(['Next Action', plD.nextAction||'—']);
          adsSheet.appendRow([]);
        });

        // Plans
        const plans = (periodData.plans||[]).filter(p=>p.text);
        if (plans.length) {
          adsSheet.appendRow(['RENCANA PERIODE DEPAN']);
          adsSheet.getRange(adsSheet.getLastRow(),1).setFontWeight('bold').setBackground('#f0fdf4');
          adsSheet.appendRow(['#','Rencana','Kategori','Target']);
          plans.forEach((p,i) => adsSheet.appendRow([i+1, p.text, p.category||'—', p.deadline||'—']));
          adsSheet.appendRow([]);
        }
      } else {
        adsSheet.appendRow(['Belum ada data untuk periode ini']);
      }

      styleAdsExportSheet_(adsSheet);

      // ── TAB: ClientName_Content ──
      const contSheet = getContentTimelineSheet_(timelineSs, client.name);
      prepareTimelineSheetForAppend_(contSheet);

      const posts = allCalendars.filter(p => p.clientId === client.id).sort((a,b) => String(a.date||'').localeCompare(String(b.date||'')));
      renderContentTimelineSheet_(contSheet, client.name, posts, nextTimelineStartRow_(contSheet, posts));
      exported++;
    });

    return { success: true, exported, spreadsheetUrl: timelineSs.getUrl(), dataSpreadsheetUrl: ss.getUrl(), message: exported + ' klien diekspor (' + exported + ' timeline tab + ' + exported + ' ads tab)' };
  } catch(e) {
    Logger.log('exportAllToSheets error: ' + e);
    return { success: false, error: e.toString() };
  }
}

function styleAdsExportSheet_(sheet) {
  const last = sheet.getLastRow();
  const cols = Math.max(sheet.getLastColumn(), 6);
  if (!last) return;
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(4);
  sheet.getRange(1,1,last,cols)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(1,1,1,cols).merge()
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  sheet.getRange(2,1,2,cols)
    .setBackground('#eef4ff')
    .setFontColor('#1f2937')
    .setFontWeight('bold');
  for (let r = 1; r <= last; r++) {
    const first = String(sheet.getRange(r,1).getValue() || '');
    if (first.indexOf('PLATFORM:') === 0) {
      sheet.getRange(r,1,1,cols)
        .setBackground('#dbeafe')
        .setFontWeight('bold')
        .setFontColor('#1e40af');
    }
    if (first === 'RENCANA PERIODE DEPAN') {
      sheet.getRange(r,1,1,cols)
        .setBackground('#dcfce7')
        .setFontWeight('bold')
        .setFontColor('#166534');
    }
  }
  sheet.getRange(1,1,last,cols).setBorder(true,true,true,true,true,true,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 180);
}

function styleContentExportSheet_(sheet) {
  const last = sheet.getLastRow();
  if (!last) return;
  const cols = Math.max(sheet.getLastColumn(), 5);
  sheet.setHiddenGridlines(true);
  sheet.getRange(1,1,last,cols)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.getRange(1,1,last,Math.min(cols,5)).setBorder(true,true,true,true,true,true,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(1, 22);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 95);
}

function exportCalendarToSheet(clientId) {
  const clients = getClients().clients;
  const client = clients.find(c => c.id === clientId);
  if (!client) return { success: false, error: 'Client not found' };

  try {
    const ss = SpreadsheetApp.openById(CONTENT_TIMELINE_SS_ID);
    const sheet = getContentTimelineSheet_(ss, client.name);
    const tabName = sheet.getName();
    prepareTimelineSheetForAppend_(sheet);

    const posts = getAllCalendars().posts
      .filter(p => p.clientId === clientId)
      .sort((a,b) => String(a.date||'').localeCompare(String(b.date||'')));

    const startRow = nextTimelineStartRow_(sheet, posts);
    renderContentTimelineSheet_(sheet, client.name, posts, startRow);
    return {
      success: true,
      count: posts.length,
      sheetName: tabName,
      spreadsheetUrl: ss.getUrl(),
      sheetUrl: ss.getUrl() + '#gid=' + sheet.getSheetId()
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getContentTimelineSheet_(ss, clientName) {
  const target = getSheetByGid_(ss, CONTENT_TIMELINE_TARGET_GID);
  if (target) return target;
  const tabName = sanitizeSheetName_(clientName || 'Client');
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);
  return sheet;
}

function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

function sanitizeSheetName_(name) {
  return String(name || 'Client')
    .replace(/[\[\]\*\/\\\?:]/g, '-')
    .slice(0, 100)
    .trim() || 'Client';
}

function prepareTimelineSheetForAppend_(sheet) {
  sheet.setHiddenGridlines(true);
  const widths = [260,120,95,90,260,120,95,90,260,120,95,90,260,120,95,90,260,120,95,90,260,120,95,90];
  widths.forEach((w,i) => sheet.setColumnWidth(i + 1, w));
}

function nextTimelineStartRow_(sheet, posts) {
  const last = sheet.getLastRow();
  if (last <= 1) return 1;
  if (isEmptyTimelineTemplate_(sheet)) return 1;
  const period = getTimelinePeriod_(posts || []);
  if (period) {
    for (let r = 1; r <= last; r++) {
      const title = String(sheet.getRange(r,1).getValue() || '').toUpperCase();
      if (title.indexOf('SAMPLE TIMELINE') === 0) {
        const existingPeriod = String(sheet.getRange(r + 1,1).getValue() || '').trim();
        if (existingPeriod === period) return r;
      }
    }
  }
  return last + 3;
}

function isEmptyTimelineTemplate_(sheet) {
  try {
    const title = String(sheet.getRange(1,1).getValue() || '').toUpperCase();
    if (title.indexOf('SAMPLE TIMELINE') !== 0) return false;
    const values = sheet.getRange(5,1,Math.max(sheet.getLastRow() - 4, 1),Math.min(sheet.getLastColumn(), 8)).getValues();
    return values.every(row => !row[0] && !row[1] && !row[4] && !row[5]);
  } catch(e) {
    return false;
  }
}

function renderContentTimelineSheet_(sheet, clientName, posts, startRow) {
  startRow = startRow || 1;
  const groups = groupPostsByBatch_(posts);
  const maxPosts = groups.reduce((m,g) => Math.max(m, g.posts.length), 0);
  if (!groups.length) groups.push({ name: 'Batch 1', posts: [] });
  const lastCol = Math.max(4, groups.length * 4);
  const rowCount = Math.max(maxPosts + 4, 8);
  const clearCols = Math.max(lastCol, Math.min(Math.max(sheet.getLastColumn(), 8), 24));
  const clearRows = startRow === 1 ? Math.max(rowCount, sheet.getLastRow()) : rowCount;
  sheet.getRange(startRow,1,clearRows,clearCols).breakApart();
  sheet.getRange(startRow,1,clearRows,clearCols).clearContent();
  sheet.setHiddenGridlines(true);
  sheet.getRange(startRow,1,rowCount,lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(startRow,1,1,4).merge()
    .setValue('SAMPLE TIMELINE ' + posts.length + ' CONTENTS. POSTED WITHIN 30 DAYS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#ffffff')
    .setBorder(false,false,false,false,false,false);

  groups.forEach((group, index) => renderTimelineBatchBlock_(sheet, group.name, group.posts, startRow, 1 + (index * 4)));

  const endRow = startRow + rowCount - 1;
  if (startRow === 1) sheet.setFrozenRows(4);
  return { startRow, endRow };
}

function groupPostsByBatch_(posts) {
  const map = {};
  (posts || []).forEach(p => {
    const name = String(p.batch || 'Batch 1').trim() || 'Batch 1';
    if (!map[name]) map[name] = [];
    map[name].push(p);
  });
  return Object.keys(map).sort(batchNameSort_).map(name => ({
    name,
    posts: map[name].sort((a,b) => String(a.date||'').localeCompare(String(b.date||'')))
  }));
}

function batchNameSort_(a, b) {
  const an = String(a).match(/\d+/);
  const bn = String(b).match(/\d+/);
  if (an && bn && Number(an[0]) !== Number(bn[0])) return Number(an[0]) - Number(bn[0]);
  return String(a).localeCompare(String(b));
}

function renderTimelineBatchBlock_(sheet, batchName, posts, startRow, startCol) {
  const periodRow = startRow + 1;
  const sectionRow = startRow + 2;
  const headerRow = startRow + 3;
  const firstDataRow = startRow + 4;

  const periodText = getTimelinePeriod_(posts);
  sheet.getRange(periodRow,startCol,1,4).merge().setValue(periodText || 'PERIOD')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#cccccc')
    .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(sectionRow,startCol,1,4).merge().setValue('POSTING CONTENT')
    .setFontSize(11)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#f4cccc')
    .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(headerRow,startCol,1,4).setValues([['Content Type','Tanggal','Batch','Progres']])
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#ffffff')
    .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);

  if (posts.length) {
    const rows = posts.map((p,i) => [
      formatTimelineType_(p),
      timelineDateValue_(p.date),
      p.batch || batchName,
      p.status === 'Posted'
    ]);
    const dataRange = sheet.getRange(firstDataRow,startCol,rows.length,4);
    dataRange.setValues(rows)
      .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.getRange(firstDataRow,startCol + 1,rows.length,1).setNumberFormat('d mmm yyyy');
    sheet.getRange(firstDataRow,startCol + 3,rows.length,1).insertCheckboxes();
    const fills = rows.map((_, i) => Array(4).fill(Math.floor(i / 6) % 2 === 0 ? '#b6d7a8' : '#fce5cd'));
    dataRange.setBackgrounds(fills);
  } else {
    sheet.getRange(firstDataRow,startCol,1,4).merge().setValue('Belum ada jadwal konten')
      .setHorizontalAlignment('center').setFontStyle('italic')
      .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID);
  }
}

function formatTimelineType_(p) {
  const type = p.type || 'Content';
  const platform = p.platform ? ' - ' + p.platform : '';
  return type + platform;
}

function normalizeCalendarDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const raw = String(value).trim();
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
  try {
    const d = new Date(raw);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {}
  return raw;
}

function timelineDateValue_(value) {
  const normalized = normalizeCalendarDate_(value);
  const m = String(normalized).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return normalized || '';
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function formatTimelineDate_(value) {
  if (!value) return '';
  try {
    const d = timelineDateValue_(value);
    if (isNaN(d.getTime())) return value;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
  } catch(e) {
    return value;
  }
}

function getTimelinePeriod_(posts) {
  const dates = posts.map(p => normalizeCalendarDate_(p.date)).filter(Boolean).sort();
  if (!dates.length) return '';
  return formatTimelineDate_(dates[0]) + ' - ' + formatTimelineDate_(dates[dates.length-1]);
}

// ── MANUAL EMAIL REMINDER ─────────────────────────────────────────
function sendManualReminder(emails, type, message) {
  if (!emails || !emails.length) return { success: false, error: 'No emails provided' };
  let sent = 0, failed = 0;
  const subject = '[Media Matrix] ' + (type === 'ads' ? '📊 Reminder — Ads Update' : type === 'content' ? '🗓️ Reminder — Content Update' : '📢 Pesan dari Owner');
  emails.forEach(email => {
    try {
      MailApp.sendEmail(email.trim(), subject, message || 'Reminder dari Media Matrix Dashboard. Mohon cek dan update data kamu.');
      sent++;
    } catch(e) { failed++; }
  });
  return { success: true, sent, failed };
}

// ── AUTOMATED REMINDERS ───────────────────────────────────────────
function runDailyReminders() {
  sendAdsReminders();
  sendContentReminders();
}

function sendAdsReminders() {
  const users = getUsersByRole('Ads Specialist');
  if (!users.length) return;
  const today = new Date();
  if (![1,3,5].includes(today.getDay())) return;
  const periodKey = getPeriodKeyFromDate(today);
  const periodData = getPeriodData(periodKey).data;
  const clients = getClients().clients;

  users.forEach(user => {
    const unfilled = clients.filter(c => !periodData[c.id] || !Object.keys(periodData[c.id]?.platforms||{}).length);
    const subject = '[Media Matrix] 📊 Ads Check-In — ' + today.toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long'});
    let body = 'Halo ' + user.name + ',\n\n';
    if (unfilled.length) { body += '📋 KLIEN BELUM DIUPDATE:\n'; unfilled.forEach(c => { body += '   • ' + c.name + '\n'; }); body += '\n'; }
    else body += '✅ Semua klien sudah diupdate!\n\n';
    body += 'Pertanyaan check-in:\n1. Ada performa klien yang perlu disesuaikan segera?\n2. Ada creative yang mau di-test?\n3. Ada budget yang perlu dioptimasi?\n\n— Media Matrix Dashboard';
    try { MailApp.sendEmail(user.email, subject, body); } catch(e) {}
  });
}

function sendContentReminders() {
  const users = getUsersByRole('Content Team');
  if (!users.length) return;
  const today = new Date();
  if (![1,4].includes(today.getDay())) return;
  const clients = getClients().clients;
  const allBatches = getAllBatches().batches;

  users.forEach(user => {
    const subject = '[Media Matrix] 🗓️ Content Check-In — ' + today.toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long'});
    let body = 'Halo ' + user.name + ',\n\nStatus batch konten:\n\n';
    clients.forEach(c => {
      const batches = allBatches.filter(b => b.clientId === c.id);
      if (!batches.length) return;
      const latest = batches[batches.length-1];
      const done = (latest.checklist||[]).filter(s=>s.done).length;
      const curStep = (latest.checklist||[]).find(s=>!s.done)?.step || 'Selesai';
      body += '▸ ' + c.name + ': ' + curStep + ' (' + done + '/5 tahap)\n';
    });
    body += '\n— Media Matrix Dashboard';
    try { MailApp.sendEmail(user.email, subject, body); } catch(e) {}
  });
}

function getPeriodKeyFromDate(date) {
  const d = new Date(date);
  const day = d.getDate();
  const p = day <= 10 ? 1 : day <= 20 ? 2 : 3;
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-P' + p;
}

// ── ADS CAMPAIGN REPORTS ─────────────────────────────────────────
function getAllCampaignReports() {
  const sheet = sh('CampaignReports', ['periodKey','clientId','clientName','periodLabel','data_json','updatedBy','updatedAt']);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { reports: [] };
  return {
    reports: rows.slice(1).filter(r => r[0] && r[1]).map(r => ({
      periodKey:r[0], clientId:r[1], clientName:r[2], periodLabel:r[3],
      data: safeJson_(r[4], { sections:{} }), updatedBy:r[5], updatedAt:r[6]
    }))
  };
}

function saveCampaignReport(periodKey, clientId, data, user) {
  const sheet = sh('CampaignReports', ['periodKey','clientId','clientName','periodLabel','data_json','updatedBy','updatedAt']);
  const clients = getClients().clients;
  const client = clients.find(c => c.id === clientId) || { id:clientId, name:clientId };
  const periodLabel = formatPeriodLabel(periodKey);
  const rows = sheet.getDataRange().getValues();
  const payload = JSON.stringify(data || { sections:{} });
  const now = new Date().toISOString();
  let row = 0;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === periodKey && rows[i][1] === clientId) { row = i + 1; break; }
  }
  const values = [[periodKey, clientId, client.name, periodLabel, payload, user?.name || '', now]];
  if (row) sheet.getRange(row,1,1,7).setValues(values);
  else sheet.appendRow(values[0]);
  styleRows(sheet);
  const exportResult = renderAllAdsCampaignReports_();
  const clientSheetUrl = exportResult.clientUrls?.[clientId] || ADS_REPORT_SS_URL;
  return { success:true, spreadsheetUrl:ADS_REPORT_SS_URL, sheetUrl:clientSheetUrl, updatedTabs:exportResult.updatedTabs, clientName:client.name };
}

function safeJson_(value, fallback) {
  try { return JSON.parse(value || ''); } catch(e) { return fallback; }
}

function renderAllAdsCampaignReports_() {
  const reports = getAllCampaignReports().reports;
  const ss = SpreadsheetApp.openById(ADS_REPORT_SS_ID);
  const grouped = {};
  reports.forEach(rep => {
    if (!grouped[rep.clientId]) grouped[rep.clientId] = [];
    grouped[rep.clientId].push(rep);
  });
  const clientUrls = {};
  const updatedTabs = [];
  Object.keys(grouped).forEach(clientId => {
    const sheet = renderClientAdsReportSheet_(ss, grouped[clientId]);
    clientUrls[clientId] = ADS_REPORT_SS_URL + '#gid=' + sheet.getSheetId();
    updatedTabs.push(sheet.getName());
  });
  markOldObjectiveTabs_(ss);
  renderMonthlySummary_(ss, reports);
  updatedTabs.push('Monthly Summary');
  return { updatedTabs, clientUrls };
}

function cleanSheetName_(name) {
  return String(name || 'Client').replace(/[\\\/\?\*\[\]\:]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'Client';
}

function renderClientAdsReportSheet_(ss, reports) {
  const clientName = reports[0]?.clientName || 'Client';
  const sheet = sheetForReport_(ss, cleanSheetName_(clientName));
  [230,140,150,180,110,140,140,140].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  let r = 1;
  reports.sort((a,b)=>String(a.periodKey||'').localeCompare(String(b.periodKey||''))).forEach(rep => {
    const sections = rep.data?.sections || {};
    Object.keys(sections).forEach(id => {
      const sec = sections[id] || {};
      const objective = sec.objective || inferObjective_(sec.platform, sec.kpi);
      const agency = reportRows_(sec, 'agency');
      const client = reportRows_(sec, 'client');
      if (!agency.length && !client.length) return;
      if (agency.length) r = renderAgencyBlock_(sheet, r, clientName, rep.periodLabel || formatPeriodLabel(rep.periodKey), sec, objective, agency);
      if (objective === 'meta_leads' && client.length) r = renderLeadClientBlock_(sheet, r, clientName, rep.periodLabel || formatPeriodLabel(rep.periodKey), client);
    });
  });
  if (r === 1) sheet.getRange(1,1).setValue('Belum ada Ads campaign report untuk ' + clientName + '.');
  return sheet;
}

function objectiveTitle_(objective) {
  return {
    marketplace: 'SHOPEE ADS/TOKOPEDIA ADS',
    meta_profile: 'META (PROFILE VISITS)',
    meta_lpv: 'META (LANDING PAGE VIEWS)',
    meta_leads: 'META (LEADS)'
  }[objective] || 'META ADS';
}

function agencyHeaders_(objective) {
  if (objective === 'marketplace') return ['Campaign Name','Spend','Sales','ROAS','Impressions','Cost Per Click'];
  if (objective === 'meta_leads') return ['Campaign Name','Spend','Total Leads','Cost Per Lead','CTR%','Impressions','Cost Per Click'];
  if (objective === 'meta_lpv') return ['Campaign Name','Spend','Landing Page Views','Cost Per Landing Page View','CTR%','Impressions','Cost Per Click'];
  return ['Campaign Name','Spend','Profile Visits','Cost Per Visit','CTR%','Impressions','Cost Per Click'];
}

function divFormula_(numCol, denCol, row) {
  return '=IFERROR(' + numCol + row + '/' + denCol + row + ',"")';
}

function renderAgencyBlock_(sheet, startRow, clientName, periodLabel, section, objective, campaigns) {
  const headers = agencyHeaders_(objective);
  const cols = headers.length;
  sheet.getRange(startRow,1,1,cols).setValues([[clientName, 'Period: ' + periodLabel].concat(Array(Math.max(cols-2,0)).fill(''))]);
  sheet.getRange(startRow+1,1,1,cols).setValues([[objectiveTitle_(objective), 'Period (Every 10 Days)'].concat(Array(Math.max(cols-2,0)).fill(''))]);
  sheet.getRange(startRow+2,1,1,cols).setValues([headers]);
  let spend=0, sales=0, result=0, impressions=0, cpcTotal=0, cpcCount=0, costTotal=0, costCount=0;
  campaigns.forEach((c,i) => {
    const s=moneyNum_(c.spend), sale=moneyNum_(c.sales), res=moneyNum_(c.result || c.leads), imp=moneyNum_(c.impressions), cpc=moneyNum_(c.cpc), cost=moneyNum_(c.cost || c.cpl);
    spend+=s; sales+=sale; result+=res; impressions+=imp; if(cpc){cpcTotal+=cpc;cpcCount++;} if(cost){costTotal+=cost;costCount++;}
    let row;
    const sheetRow = startRow + 3 + i;
    if (objective === 'marketplace') row = [c.name || '', s || '', sale || '', divFormula_('C','B',sheetRow), imp || '', cpc || ''];
    else if (objective === 'meta_leads') row = [c.name || '', s || '', res || '', divFormula_('B','C',sheetRow), c.ctr || '', imp || '', cpc || ''];
    else row = [c.name || '', s || '', res || '', divFormula_('B','C',sheetRow), c.ctr || '', imp || '', cpc || ''];
    sheet.getRange(startRow+3+i,1,1,cols).setValues([row]);
  });
  const totalRow = startRow + 3 + campaigns.length;
  let total;
  if (objective === 'marketplace') total = ['TOTAL', '=SUM(B' + (startRow+3) + ':B' + (totalRow-1) + ')', '=SUM(C' + (startRow+3) + ':C' + (totalRow-1) + ')', divFormula_('C','B',totalRow), '=SUM(E' + (startRow+3) + ':E' + (totalRow-1) + ')', '=IFERROR(AVERAGE(F' + (startRow+3) + ':F' + (totalRow-1) + '),"")'];
  else total = ['TOTAL', '=SUM(B' + (startRow+3) + ':B' + (totalRow-1) + ')', '=SUM(C' + (startRow+3) + ':C' + (totalRow-1) + ')', divFormula_('B','C',totalRow), '', '=SUM(F' + (startRow+3) + ':F' + (totalRow-1) + ')', '=IFERROR(AVERAGE(G' + (startRow+3) + ':G' + (totalRow-1) + '),"")'];
  sheet.getRange(totalRow,1,1,cols).setValues([total]).setFontWeight('bold').setBackground('#eef2ff');
  styleReportBlock_(sheet, startRow, cols, objective === 'marketplace' ? '#ffedd5' : objective === 'meta_leads' ? '#dcfce7' : '#e0e7ff');
  return totalRow + 3;
}

function renderLeadClientBlock_(sheet, startRow, clientName, periodLabel, rows) {
  const headers = ['Campaign Name','Total Leads','Hot Leads/Potential','Warm Leads/Response Only','Cold Leads/No Response','Closed'];
  const cols = headers.length;
  sheet.getRange(startRow,1,1,cols).setValues([['DIISI OLEH ADMIN CLIENT (KEBUTUHAN EVALUASI UNTUK MM)', clientName, 'Period: ' + periodLabel, '', '', '']]);
  sheet.getRange(startRow+1,1,1,cols).setValues([['META (LEADS)', 'Period (Every 10 Days)', '', '', '', '']]);
  sheet.getRange(startRow+2,1,1,cols).setValues([headers]);
  let leads=0, hot=0, warm=0, cold=0, closed=0;
  rows.forEach((c,i) => {
    const l=moneyNum_(c.totalLeads), h=moneyNum_(c.hot), w=moneyNum_(c.warm), co=moneyNum_(c.cold), cl=moneyNum_(c.closed);
    leads+=l; hot+=h; warm+=w; cold+=co; closed+=cl;
    sheet.getRange(startRow+3+i,1,1,cols).setValues([[c.name || '', l || '', h || '', w || '', co || '', cl || '']]);
  });
  const totalRow = startRow + 3 + rows.length;
  sheet.getRange(totalRow,1,1,cols).setValues([['TOTAL', leads || '', hot || '', warm || '', cold || '', closed || '']]).setFontWeight('bold').setBackground('#dcfce7');
  styleReportBlock_(sheet, startRow, cols, '#dcfce7');
  return totalRow + 3;
}

function markOldObjectiveTabs_(ss) {
  ['Marketplace Ads','Meta Ads (Profile Visits)','Meta Ads (LP Views)','Meta Ads (WhatsApp)'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.clear();
    sheet.getRange(1,1).setValue('Report sekarang dipisah per tab nama client.');
    sheet.setHiddenGridlines(true);
  });
}

function inferObjective_(platform, kpi) {
  const p = String(platform || '').toLowerCase();
  const k = String(kpi || '').toLowerCase();
  if (p.includes('shopee') || p.includes('tokopedia')) return 'marketplace';
  if (k.includes('landing') || k.includes('lp')) return 'meta_lpv';
  if (k.includes('profile') || k.includes('visit')) return 'meta_profile';
  if (k.includes('lead') || k.includes('wa') || k.includes('whatsapp')) return 'meta_leads';
  return 'meta_profile';
}

function sheetForReport_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  sheet.setHiddenGridlines(true);
  return sheet;
}

function moneyNum_(value) {
  const n = parseFloat(String(value || '').replace(/[^0-9.-]/g,''));
  return isNaN(n) ? 0 : n;
}

function reportRows_(section, type) {
  const rows = type === 'client' ? (section.clientRows || []) : (section.campaigns || []);
  return rows.filter(r => Object.values(r || {}).some(v => String(v || '').trim() !== ''));
}

function styleReportBlock_(sheet, startRow, cols, titleColor) {
  sheet.getRange(startRow,1,1,cols).setBackground('#111827').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(startRow+1,1,1,cols).setBackground(titleColor || '#dbeafe').setFontWeight('bold');
  sheet.getRange(startRow+2,1,1,cols).setBackground('#f3f4f6').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sheet.getRange(startRow,1,Math.max(3, sheet.getLastRow() - startRow + 1),cols).setBorder(true,true,true,true,true,true,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow,1,Math.max(3, sheet.getLastRow() - startRow + 1),cols).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
}

function renderMarketplaceSheet_(ss, entries) {
  const sheet = sheetForReport_(ss, 'Marketplace Ads');
  const widths = [190,230,25,140,140,100,140,140];
  widths.forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  let r = 1;
  entries.forEach(entry => {
    const campaigns = reportRows_(entry.section, 'agency');
    if (!campaigns.length) return;
    sheet.getRange(r,1,1,8).setValues([['CLIENT: ' + entry.clientName, '', '', 'PERIOD: ' + entry.periodLabel, '', '', '', '']]);
    sheet.getRange(r+1,1,1,8).setValues([['SHOPEE ADS/TOKOPEDIA ADS', '', '', 'Period (Every 10 Days)', '', '', '', '']]);
    sheet.getRange(r+2,1,1,8).setValues([['CAMPAIGN NAME','Campaign Name','','Spend','Sales','ROAS','Impressions','Cost Per Click']]);
    let spend=0, sales=0, impressions=0, cpcTotal=0, cpcCount=0;
    campaigns.forEach((c,i) => {
      const s=moneyNum_(c.spend), sale=moneyNum_(c.sales), imp=moneyNum_(c.impressions), cpc=moneyNum_(c.cpc);
      spend+=s; sales+=sale; impressions+=imp; if(cpc){cpcTotal+=cpc;cpcCount++;}
      sheet.getRange(r+3+i,1,1,8).setValues([['Campaign ' + (i+1), c.name || '', '', s || '', sale || '', c.roas || (s ? sale / s : ''), imp || '', cpc || '']]);
    });
    const tr = r + 3 + campaigns.length;
    sheet.getRange(tr,1,1,8).setValues([['TOTAL','', '', spend || '', sales || '', spend ? sales / spend : '', impressions || '', cpcCount ? cpcTotal / cpcCount : '']]).setFontWeight('bold').setBackground('#fff7ed');
    styleReportBlock_(sheet, r, 8, '#ffedd5');
    r = tr + 3;
  });
  if (r === 1) sheet.getRange(1,1).setValue('Belum ada data Marketplace Ads.');
}

function renderMetaSheet_(ss, sheetName, title, resultHeader, costHeader, entries) {
  const sheet = sheetForReport_(ss, sheetName);
  const widths = [190,230,140,150,170,100,140,140];
  widths.forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  let r = 1;
  entries.forEach(entry => {
    const campaigns = reportRows_(entry.section, 'agency');
    if (!campaigns.length) return;
    sheet.getRange(r,1,1,8).setValues([['CLIENT: ' + entry.clientName, '', 'PERIOD: ' + entry.periodLabel, '', '', '', '', '']]);
    sheet.getRange(r+1,1,1,8).setValues([[title, '', 'Period (Every 10 Days)', '', '', '', '', '']]);
    sheet.getRange(r+2,1,1,8).setValues([['CAMPAIGN NAME','Campaign Name','Spend',resultHeader,costHeader,'CTR%','Impressions','Cost Per Click']]);
    let spend=0, result=0, imp=0, costTotal=0, costCount=0, cpcTotal=0, cpcCount=0;
    campaigns.forEach((c,i) => {
      const s=moneyNum_(c.spend), res=moneyNum_(c.result), cost=moneyNum_(c.cost), im=moneyNum_(c.impressions), cpc=moneyNum_(c.cpc);
      spend+=s; result+=res; imp+=im; if(cost){costTotal+=cost;costCount++;} if(cpc){cpcTotal+=cpc;cpcCount++;}
      sheet.getRange(r+3+i,1,1,8).setValues([['Campaign ' + (i+1), c.name || '', s || '', res || '', cost || (res ? s / res : ''), c.ctr || '', im || '', cpc || '']]);
    });
    const tr = r + 3 + campaigns.length;
    sheet.getRange(tr,1,1,8).setValues([['TOTAL','', spend || '', result || '', result ? spend / result : (costCount ? costTotal / costCount : ''), '', imp || '', cpcCount ? cpcTotal / cpcCount : '']]).setFontWeight('bold').setBackground('#eef2ff');
    styleReportBlock_(sheet, r, 8, '#e0e7ff');
    r = tr + 3;
  });
  if (r === 1) sheet.getRange(1,1).setValue('Belum ada data ' + sheetName + '.');
}

function renderLeadsSheet_(ss, entries) {
  const sheet = sheetForReport_(ss, 'Meta Ads (WhatsApp)');
  const widths = [190,230,140,120,150,100,140,140];
  widths.forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  let r = 1;
  entries.forEach(entry => {
    const agency = reportRows_(entry.section, 'agency');
    const client = reportRows_(entry.section, 'client');
    if (!agency.length && !client.length) return;
    if (agency.length) {
      sheet.getRange(r,1,1,8).setValues([['DIISI OLEH TEAM MEDIA MATRIX', 'CLIENT: ' + entry.clientName, 'PERIOD: ' + entry.periodLabel, '', '', '', '', '']]);
      sheet.getRange(r+1,1,1,8).setValues([['META (LEADS)', '', 'Period (Every 10 Days)', '', '', '', '', '']]);
      sheet.getRange(r+2,1,1,8).setValues([['CAMPAIGN NAME','Campaign Name','Spend','Total Leads','Cost Per Lead','CTR%','Impressions','Cost Per Click']]);
      let spend=0, leads=0, imp=0, cpcTotal=0, cpcCount=0;
      agency.forEach((c,i) => {
        const s=moneyNum_(c.spend), l=moneyNum_(c.leads), im=moneyNum_(c.impressions), cpc=moneyNum_(c.cpc);
        spend+=s; leads+=l; imp+=im; if(cpc){cpcTotal+=cpc;cpcCount++;}
        sheet.getRange(r+3+i,1,1,8).setValues([['Campaign ' + (i+1), c.name || '', s || '', l || '', c.cpl || (l ? s / l : ''), c.ctr || '', im || '', cpc || '']]);
      });
      const tr = r + 3 + agency.length;
      sheet.getRange(tr,1,1,8).setValues([['TOTAL','', spend || '', leads || '', leads ? spend / leads : '', '', imp || '', cpcCount ? cpcTotal / cpcCount : '']]).setFontWeight('bold').setBackground('#dcfce7');
      styleReportBlock_(sheet, r, 8, '#dcfce7');
      r = tr + 3;
    }
    if (client.length) {
      sheet.getRange(r,1,1,6).setValues([['DIISI OLEH ADMIN CLIENT (KEBUTUHAN EVALUASI UNTUK MM)', 'CLIENT: ' + entry.clientName, 'PERIOD: ' + entry.periodLabel, '', '', '']]);
      sheet.getRange(r+1,1,1,6).setValues([['META (LEADS)', '', 'Period (Every 10 Days)', '', '', '']]);
      sheet.getRange(r+2,1,1,6).setValues([['CAMPAIGN NAME','Total Leads','Hot Leads/Potential','Warm Leads/Response Only','Cold Leads/No Response','Closed']]);
      let leads=0, hot=0, warm=0, cold=0, closed=0;
      client.forEach((c,i) => {
        const l=moneyNum_(c.totalLeads), h=moneyNum_(c.hot), w=moneyNum_(c.warm), co=moneyNum_(c.cold), cl=moneyNum_(c.closed);
        leads+=l; hot+=h; warm+=w; cold+=co; closed+=cl;
        sheet.getRange(r+3+i,1,1,6).setValues([[c.name || 'Campaign ' + (i+1), l || '', h || '', w || '', co || '', cl || '']]);
      });
      const tr = r + 3 + client.length;
      sheet.getRange(tr,1,1,6).setValues([['TOTAL', leads || '', hot || '', warm || '', cold || '', closed || '']]).setFontWeight('bold').setBackground('#dcfce7');
      styleReportBlock_(sheet, r, 6, '#dcfce7');
      r = tr + 3;
    }
  });
  if (r === 1) sheet.getRange(1,1).setValue('Belum ada data Meta Leads.');
}

function renderMonthlySummary_(ss, reports) {
  const sheet = sheetForReport_(ss, 'Monthly Summary');
  const headers = ['Month','Client','Objective','Spend','Result','Cost Per Result','ROAS','Closed','Closing Rate','Perspective'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#111827').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  const rows = [];
  reports.forEach(rep => {
    const month = String(rep.periodLabel || '').replace(/^.*?([A-Za-z]+ \d{4})$/,'$1') || rep.periodKey;
    const sections = rep.data?.sections || {};
    Object.keys(sections).forEach(id => {
      const sec = sections[id] || {};
      const obj = sec.objective || inferObjective_(sec.platform, sec.kpi);
      const campaigns = reportRows_(sec, 'agency');
      const clientRows = reportRows_(sec, 'client');
      let spend=0, result=0, sales=0, closed=0;
      campaigns.forEach(c => {
        spend += moneyNum_(c.spend);
        sales += moneyNum_(c.sales);
        result += moneyNum_(c.result || c.leads || c.sales);
      });
      clientRows.forEach(c => closed += moneyNum_(c.closed));
      const perspective = obj === 'marketplace' ? 'Higher ROAS is better' : obj === 'meta_leads' ? 'Best judged by closing + CPL' : 'Lower cost per result is better';
      if (spend || result || closed) rows.push([month, rep.clientName, obj, spend || '', result || '', '', '', closed || '', '', perspective]);
    });
  });
  if (rows.length) {
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);
    for (let i = 0; i < rows.length; i++) {
      const row = i + 2;
      sheet.getRange(row,6).setFormula('=IFERROR(D' + row + '/E' + row + ',"")');
      sheet.getRange(row,7).setFormula('=IFERROR(E' + row + '/D' + row + ',"")');
      sheet.getRange(row,9).setFormula('=IFERROR(H' + row + '/E' + row + ',"")');
    }
  }
  sheet.getRange(1,1,Math.max(2,rows.length+1),headers.length).setFontFamily('Arial').setFontSize(10).setBorder(true,true,true,true,true,true,'#d1d5db',SpreadsheetApp.BorderStyle.SOLID);
  [120,220,170,130,120,150,100,100,120,260].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
  if (!rows.length) sheet.getRange(2,1).setValue('Belum ada data untuk summary.');
}

// ── TRIGGER SETUP ─────────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runDailyReminders').timeBased().everyDays(1).atHour(8).create();
  Logger.log('Daily trigger set at 8 AM');
}

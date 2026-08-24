/**
 * Business Leads & Clients extension for an existing productivity workbook.
 * Run buildBusinessLeadsSheet() once from the Apps Script editor.
 */
const BUSINESS_LEADS = Object.freeze({
  sheetName: 'Business Leads & Clients',
  adminName: 'ADMIN',
  headerRow: 8,
  firstDataRow: 9,
  dataRows: 1000,
  columns: 23,
  headers: [
    'Lead ID', 'Business Name', 'Owner / Contact Name', 'Business Phone',
    'Business Email', 'Website', 'Business Location', 'Google Rating',
    'Google Reviews', 'Business Type', 'Contacted?', 'Contact Date',
    'Contact Method', 'Contact Person From Our Team', 'Response Status',
    'Sale Made?', 'Sale Date', 'Package', 'Monthly Value', 'One-Off Value',
    'Next Follow-Up', 'Notes', 'Last Updated'
  ],
  colours: {
    navy: '#17324D', blue: '#486B8A', pale: '#EAF0F6', pale2: '#F5F8FB',
    green: '#C6EFCE', greenText: '#006100', lightGreen: '#E2F0D9',
    amber: '#FFEB9C', amberText: '#9C6500', red: '#FFC7CE', redText: '#9C0006',
    grey: '#D9E1F2', lightGrey: '#E7E6E6', quoteBlue: '#BDD7EE',
    purple: '#D9E1F2', purpleText: '#203864', white: '#FFFFFF'
  }
});

/** Adds/rebuilds only Business Leads & Clients and appends isolated ADMIN lists. */
function buildBusinessLeadsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = BUSINESS_LEADS;
  const lists = ensureBusinessLeadsAdminLists_(ss);
  let sheet = ss.getSheetByName(c.sheetName);
  if (!sheet) sheet = ss.insertSheet(c.sheetName);
  sheet.clear();
  sheet.clearConditionalFormatRules();
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  if (sheet.getFilter()) sheet.getFilter().remove();
  ensureSheetSize_(sheet, c.firstDataRow + c.dataRows - 1, 40);

  sheet.setHiddenGridlines(true).setFrozenRows(c.headerRow);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily('Arial').setVerticalAlignment('middle');
  buildSummary_(sheet);
  buildMainTable_(sheet, lists);
  seedBusinessLeadSamples_(sheet);
  buildSideSections_(sheet);
  applyBusinessLeadConditionalFormatting_(sheet);
  sheet.getRange(c.headerRow, 1, c.dataRows + 1, c.columns).createFilter();
  ensureBusinessLeadsEditTrigger_();
  sheet.activate();
  SpreadsheetApp.flush();
}

function ensureBusinessLeadsAdminLists_(ss) {
  let admin = ss.getSheetByName(BUSINESS_LEADS.adminName);
  if (!admin) admin = ss.insertSheet(BUSINESS_LEADS.adminName);
  const definitions = [
    ['Business Types', 'BusinessTypesList', ['Office','Retail','Gym','Medical','Dental','Hospitality','Real Estate','Construction','Warehouse','Childcare','Education','Automotive','Other']],
    ['Contact Methods', 'ContactMethodsList', ['Phone','Email','Website Form','In Person','SMS','Social Media','Referral','Other']],
    ['Lead Statuses', 'LeadStatusesList', ['Not Contacted','No Response','Interested','Follow Up','Quote Requested','Quote Sent','Negotiating','Not Interested','Won','Lost']],
    ['Client Packages', 'ClientPackagesList', ['Basic','Standard','Premium','Weekly Clean','Twice Weekly','Fortnightly','Monthly','One-Off Clean','Deep Clean','Floor Maintenance','Custom']],
    ['Our Team Members', 'BusinessLeadTeamList', ['Team Member 1']]
  ];
  const result = {};
  definitions.forEach(def => {
    const oldNames = ss.getNamedRanges().filter(n => n.getName() === def[1]);
    const existing = oldNames.length ? oldNames[0].getRange() : findAdminSection_(admin, def[0]);
    const col = existing ? existing.getColumn() : Math.max(1, admin.getLastColumn() + 2);
    if (admin.getMaxColumns() < col) admin.insertColumnsAfter(admin.getMaxColumns(), col - admin.getMaxColumns());
    // Clear only the old list cells owned by this extension, never an unrelated ADMIN area.
    if (existing) {
      const firstListRow = existing.getRow() === 1 ? 2 : existing.getRow();
      const oldHeight = existing.getRow() === 1 ? Math.max(1, existing.getNumRows() - 1) : existing.getNumRows();
      admin.getRange(firstListRow, col, oldHeight, 1).clearContent();
    }
    admin.getRange(1, col).setValue(def[0]).setFontFamily('Arial').setFontWeight('bold')
      .setBackground(BUSINESS_LEADS.colours.blue).setFontColor('#FFFFFF');
    admin.getRange(2, col, def[2].length, 1).setValues(def[2].map(v => [v])).setFontFamily('Arial');
    const range = admin.getRange(2, col, def[2].length, 1);
    // Removing every same-name definition first makes reruns duplicate-safe.
    oldNames.forEach(n => n.remove());
    ss.setNamedRange(def[1], range);
    result[def[1]] = range;
  });
  return result;
}

function findAdminSection_(admin, title) {
  if (!admin.getLastRow() || !admin.getLastColumn()) return null;
  const values = admin.getDataRange().getDisplayValues();
  for (let r=0; r<values.length; r++) for (let c=0; c<values[r].length; c++) {
    if (values[r][c].trim().toLowerCase() === title.toLowerCase()) {
      let count=0; while (r+1+count<values.length && values[r+1+count][c] !== '') count++;
      return admin.getRange(r+2, c+1, Math.max(1,count), 1);
    }
  }
  return null;
}

function buildSummary_(sheet) {
  const labels = ['Total Businesses','Not Contacted','Businesses Contacted','Interested Leads','Quotes Sent','Follow-Ups Due','Overdue Follow-Ups','Sales Won','Sales Lost','Conversion Rate','Total Monthly Client Value','Total One-Off Sales Value'];
  const f = BUSINESS_LEADS.firstDataRow;
  const end = f + BUSINESS_LEADS.dataRows - 1;
  const formulas = [
    `=COUNTIF(B${f}:B${end},"<>")`, `=COUNTIF(O${f}:O${end},"Not Contacted")`,
    `=COUNTIF(K${f}:K${end},TRUE)`, `=COUNTIF(O${f}:O${end},"Interested")`,
    `=COUNTIF(O${f}:O${end},"Quote Sent")`,
    `=COUNTIFS(U${f}:U${end},">="&TODAY(),U${f}:U${end},"<="&TODAY()+3,P${f}:P${end},FALSE)`,
    `=COUNTIFS(U${f}:U${end},"<"&TODAY(),U${f}:U${end},"<>",P${f}:P${end},FALSE)`,
    `=COUNTIF(O${f}:O${end},"Won")`, `=COUNTIF(O${f}:O${end},"Lost")`,
    `=IFERROR(COUNTIF(O${f}:O${end},"Won")/COUNTIF(K${f}:K${end},TRUE),0)`,
    `=SUM(S${f}:S${end})`, `=SUM(T${f}:T${end})`
  ];
  sheet.getRange(1, 1, 1, 23).merge().setValue('BUSINESS LEADS & CLIENTS')
    .setBackground(BUSINESS_LEADS.colours.navy).setFontColor('#FFFFFF').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  labels.forEach((label, i) => {
    const col = 1 + (i % 6) * 4;
    const row = i < 6 ? 2 : 5;
    sheet.getRange(row, col, 1, 3).merge().setValue(label).setBackground(BUSINESS_LEADS.colours.blue)
      .setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row + 1, col, 1, 3).merge().setFormula(formulas[i]).setBackground(BUSINESS_LEADS.colours.pale)
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  });
  sheet.getRange(6, 13, 1, 3).setNumberFormat('0.0%');
  sheet.getRange(6, 17, 1, 7).setNumberFormat('$#,##0.00');
}

function buildMainTable_(sheet, lists) {
  const c = BUSINESS_LEADS;
  const end = c.firstDataRow + c.dataRows - 1;
  sheet.getRange(c.headerRow, 1, 1, c.columns).setValues([c.headers]).setBackground(c.colours.navy)
    .setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  const body = sheet.getRange(c.firstDataRow, 1, c.dataRows, c.columns);
  body.setBackground(c.colours.pale2).setFontColor('#243746');
  for (let row = c.firstDataRow + 1; row <= end; row += 2) sheet.getRange(row, 1, 1, c.columns).setBackground(c.colours.pale);
  sheet.getRange(c.firstDataRow, 4, c.dataRows, 1).setNumberFormat('@');
  sheet.getRange(c.firstDataRow, 8, c.dataRows, 1).setNumberFormat('0.0');
  sheet.getRange(c.firstDataRow, 9, c.dataRows, 1).setNumberFormat('0');
  [12,17,21].forEach(col => sheet.getRange(c.firstDataRow, col, c.dataRows, 1).setNumberFormat('dd/mm/yyyy'));
  [19,20].forEach(col => sheet.getRange(c.firstDataRow, col, c.dataRows, 1).setNumberFormat('$#,##0.00'));
  sheet.getRange(c.firstDataRow, 23, c.dataRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(c.firstDataRow, 22, c.dataRows, 1).setWrap(true);
  sheet.getRange(c.firstDataRow, 11, c.dataRows, 1).insertCheckboxes();
  sheet.getRange(c.firstDataRow, 16, c.dataRows, 1).insertCheckboxes();
  const rangeRule = range => SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(false).build();
  sheet.getRange(c.firstDataRow, 10, c.dataRows, 1).setDataValidation(rangeRule(lists.BusinessTypesList));
  sheet.getRange(c.firstDataRow, 13, c.dataRows, 1).setDataValidation(rangeRule(lists.ContactMethodsList));
  sheet.getRange(c.firstDataRow, 15, c.dataRows, 1).setDataValidation(rangeRule(lists.LeadStatusesList));
  sheet.getRange(c.firstDataRow, 18, c.dataRows, 1).setDataValidation(rangeRule(lists.ClientPackagesList));
  const teamRule = findExistingTeamRange_();
  if (teamRule) sheet.getRange(c.firstDataRow, 14, c.dataRows, 1).setDataValidation(rangeRule(teamRule));
  const ratingRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 5).setAllowInvalid(false).build();
  const wholeRule = SpreadsheetApp.newDataValidation().requireFormulaSatisfied(`=AND(ISNUMBER(I${c.firstDataRow}),I${c.firstDataRow}>=0,MOD(I${c.firstDataRow},1)=0)`).setAllowInvalid(false).build();
  sheet.getRange(c.firstDataRow, 8, c.dataRows, 1).setDataValidation(ratingRule);
  sheet.getRange(c.firstDataRow, 9, c.dataRows, 1).setDataValidation(wholeRule);
  sheet.setColumnWidths(1, c.columns, 115);
  sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 180); sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 160); sheet.setColumnWidth(14, 180); sheet.setColumnWidth(15, 130);
  sheet.setColumnWidth(22, 260); sheet.setColumnWidth(23, 145);
  sheet.setRowHeight(c.headerRow, 42);
}

function findExistingTeamRange_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ['TeamMembersList','StaffList','TeamList','ContactPersonsList','BusinessLeadTeamList'];
  for (const name of names) { const r = ss.getRangeByName(name); if (r) return r; }
  const admin = ss.getSheetByName(BUSINESS_LEADS.adminName);
  if (!admin) return null;
  const values = admin.getDataRange().getDisplayValues();
  const wanted = ['team members','staff','our team','contact persons'];
  for (let r = 0; r < values.length; r++) for (let c = 0; c < values[r].length; c++) {
    if (wanted.indexOf(values[r][c].trim().toLowerCase()) !== -1) {
      let count = 0; while (r + 1 + count < values.length && values[r + 1 + count][c] !== '') count++;
      if (count) return admin.getRange(r + 2, c + 1, count, 1);
    }
  }
  return null; // Existing ADMIN is not altered merely to invent staff names.
}

function buildSideSections_(sheet) {
  const f = BUSINESS_LEADS.firstDataRow, end = f + BUSINESS_LEADS.dataRows - 1;
  const stages = ['Not Contacted','No Response','Interested','Follow Up','Quote Requested','Quote Sent','Negotiating','Won','Lost'];
  sheet.getRange(1,25,1,2).merge().setValue('SALES PIPELINE').setBackground(BUSINESS_LEADS.colours.blue).setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2,25,stages.length,1).setValues(stages.map(x => [x]));
  sheet.getRange(2,26,stages.length,1).setFormulas(stages.map((x,i) => [`=COUNTIF($O$${f}:$O$${end},Y${i+2})`]));
  sheet.getRange(1,28,1,6).merge().setValue('CLIENTS').setBackground(BUSINESS_LEADS.colours.blue).setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2,28,1,6).setValues([['Business Name','Package','Monthly Value','Sale Date','Location','Contact Name']]).setBackground(BUSINESS_LEADS.colours.navy).setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(3,28).setFormula(`=IFERROR(FILTER({B${f}:B${end},R${f}:R${end},S${f}:S${end},Q${f}:Q${end},G${f}:G${end},C${f}:C${end}},P${f}:P${end}=TRUE),"")`);
  sheet.getRange(3,30, BUSINESS_LEADS.dataRows, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(3,31, BUSINESS_LEADS.dataRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(1,35,1,6).merge().setValue('NEXT 10 FOLLOW-UPS').setBackground(BUSINESS_LEADS.colours.blue).setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2,35,1,6).setValues([['Business Name','Contact Name','Phone','Email','Response Status','Next Follow-Up']]).setBackground(BUSINESS_LEADS.colours.navy).setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(3,35).setFormula(`=IFERROR(ARRAY_CONSTRAIN(SORT(FILTER({B${f}:B${end},C${f}:C${end},D${f}:D${end},E${f}:E${end},O${f}:O${end},U${f}:U${end}},P${f}:P${end}=FALSE,O${f}:O${end}<>"Lost",O${f}:O${end}<>"Not Interested",U${f}:U${end}<>""),6,TRUE),10,6),"")`);
  sheet.getRange(3,40,10,1).setNumberFormat('dd/mm/yyyy');
  const chart = sheet.newChart().asBarChart().addRange(sheet.getRange(1,25,10,2))
    .setPosition(13,25,0,0).setOption('title','Businesses by Pipeline Stage')
    .setOption('legend',{position:'none'}).setOption('colors',[BUSINESS_LEADS.colours.blue]).build();
  sheet.insertChart(chart);
}

function applyBusinessLeadConditionalFormatting_(sheet) {
  const c = BUSINESS_LEADS, f = c.firstDataRow, n = c.dataRows;
  const rules = [], rating = sheet.getRange(f,8,n,1), status = sheet.getRange(f,15,n,1), follow = sheet.getRange(f,21,n,1);
  const addNumber = (formula, bg, fg) => rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(bg).setFontColor(fg).setRanges([rating]).build());
  addNumber(`=AND($H${f}>=4.5,$H${f}<=5)`,c.colours.green,c.colours.greenText);
  addNumber(`=AND($H${f}>=4,$H${f}<4.5)`,c.colours.lightGreen,c.colours.greenText);
  addNumber(`=AND($H${f}>=3,$H${f}<4)`,c.colours.amber,c.colours.amberText);
  addNumber(`=AND(ISNUMBER($H${f}),$H${f}<3)`,c.colours.red,c.colours.redText);
  const statusColours = {'Not Contacted':[c.colours.grey,'#444444'],'No Response':[c.colours.lightGrey,'#666666'],'Interested':[c.colours.quoteBlue,'#1F4E78'],'Follow Up':[c.colours.amber,c.colours.amberText],'Quote Requested':[c.colours.amber,c.colours.amberText],'Quote Sent':[c.colours.quoteBlue,'#1F4E78'],'Negotiating':[c.colours.purple,c.colours.purpleText],'Not Interested':[c.colours.red,c.colours.redText],'Won':[c.colours.green,c.colours.greenText],'Lost':[c.colours.red,c.colours.redText]};
  Object.keys(statusColours).forEach(value => rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(value).setBackground(statusColours[value][0]).setFontColor(statusColours[value][1]).setRanges([status]).build()));
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($U${f}<>"",$U${f}<TODAY(),$P${f}=FALSE)`).setBackground(c.colours.red).setFontColor(c.colours.redText).setRanges([follow]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($U${f}>=TODAY(),$U${f}<=TODAY()+3,$P${f}=FALSE)`).setBackground(c.colours.amber).setFontColor(c.colours.amberText).setRanges([follow]).build());
  sheet.setConditionalFormatRules(rules);
}

function seedBusinessLeadSamples_(sheet) {
  const today = new Date();
  const day = d => new Date(today.getFullYear(), today.getMonth(), today.getDate()+d);
  const data = [
    ['LEAD-0001','Harbourview Office Co','Ava Collins','0412000001','ava@example.invalid','https://example.com/harbourview','Sydney NSW',4.8,126,'Office',false,'','','','Not Contacted',false,'','','','',day(2),'Initial research completed.',today],
    ['LEAD-0002','Blue Wattle Retail','Noah Bennett','0412000002','noah@example.invalid','https://example.com/wattle','Parramatta NSW',4.2,84,'Retail',true,day(-7),'Phone','','No Response',false,'','','','',day(-1),'Left a fictional voicemail.',today],
    ['LEAD-0003','Northstar Fitness Studio','Mia Hart','0412000003','mia@example.invalid','https://example.com/northstar','Richmond VIC',4.6,203,'Gym',true,day(-5),'Email','','Interested',false,'','','','',day(1),'Asked about regular service.',today],
    ['LEAD-0004','Paperbark Dental Rooms','Leo Turner','0412000004','leo@example.invalid','https://example.com/paperbark','Brisbane QLD',4.9,311,'Dental',true,day(-4),'Referral','','Quote Requested',false,'','Weekly Clean','','',day(3),'Prepare a fictional quote.',today],
    ['LEAD-0005','Laneway Lantern Cafe','Zoe Martin','0412000005','zoe@example.invalid','https://example.com/lantern','Adelaide SA',3.8,98,'Hospitality',true,day(-3),'In Person','','Quote Sent',false,'','Twice Weekly',850,'',day(2),'Quote sent for review.',today],
    ['LEAD-0006','Silver Fern Property Group','Eli Walker','0412000006','eli@example.invalid','https://example.com/silverfern','Perth WA',4.4,67,'Real Estate',true,day(-8),'Website Form','','Negotiating',false,'','Premium',1450,250,day(1),'Discussing service scope.',today],
    ['LEAD-0007','Bright Koala Learning Hub','Isla Reed','0412000007','isla@example.invalid','https://example.com/brightkoala','Hobart TAS',4.7,145,'Education',true,day(-14),'Email','','Won',true,day(-2),'Weekly Clean',1200,300,'','Fictional client won.',today],
    ['LEAD-0008','Red Gum Auto Workshop','Jack Morris','0412000008','jack@example.invalid','https://example.com/redgum','Canberra ACT',2.7,42,'Automotive',true,day(-12),'Phone','','Lost',false,'','','','',day(-4),'Chose not to proceed.',today]
  ];
  sheet.getRange(BUSINESS_LEADS.firstDataRow,1,data.length,BUSINESS_LEADS.columns).setValues(data);
}

/** Menu command: prepares and selects the first unused input row. */
function addNewBusinessLead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(BUSINESS_LEADS.sheetName);
  if (!sheet) throw new Error('Run buildBusinessLeadsSheet() first.');
  const names = sheet.getRange(BUSINESS_LEADS.firstDataRow,2,BUSINESS_LEADS.dataRows,1).getDisplayValues();
  let offset = names.findIndex(row => row[0] === '');
  if (offset < 0) throw new Error('The 1,000 prepared lead rows are full.');
  const row = BUSINESS_LEADS.firstDataRow + offset;
  const ids = sheet.getRange(BUSINESS_LEADS.firstDataRow,1,BUSINESS_LEADS.dataRows,1).getDisplayValues().flat();
  const next = ids.reduce((m,id) => Math.max(m, Number((id.match(/^LEAD-(\d+)$/)||[])[1])||0),0)+1;
  sheet.getRange(row,1).setValue('LEAD-' + String(next).padStart(4,'0'));
  sheet.getRange(row,15).setValue('Not Contacted');
  sheet.getRange(row,2).activate();
}

function openBusinessLeadsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BUSINESS_LEADS.sheetName);
  if (!sheet) throw new Error('Run buildBusinessLeadsSheet() first.');
  sheet.activate();
}

/** Call these two addItem lines on the existing Productivity System menu builder. */
function addBusinessLeadsMenu_(menu) {
  if (menu) return menu.addItem('Add New Business Lead','addNewBusinessLead').addItem('Open Business Leads & Clients','openBusinessLeadsSheet');
  SpreadsheetApp.getUi().createMenu('Productivity System')
    .addItem('Add New Business Lead','addNewBusinessLead')
    .addItem('Open Business Leads & Clients','openBusinessLeadsSheet').addToUi();
}

/** Installable edit handler: timestamps only lead rows and turns ordinary URLs into links. */
function businessLeadsEditHandler(e) {
  if (!e || !e.range) return;
  const range = e.range, sheet = range.getSheet(), c = BUSINESS_LEADS;
  if (sheet.getName() !== c.sheetName || range.getRow() < c.firstDataRow || range.getRow() >= c.firstDataRow+c.dataRows || range.getColumn() > c.columns) return;
  const start = range.getRow(), rows = range.getNumRows();
  for (let row = start; row < start+rows; row++) {
    sheet.getRange(row,23).setValue(new Date());
    if (!sheet.getRange(row,1).getValue() && sheet.getRange(row,2).getValue()) assignLeadId_(sheet,row);
  }
  if (range.getColumn() <= 6 && range.getLastColumn() >= 6) {
    for (let row=start; row<start+rows; row++) {
      const cell=sheet.getRange(row,6), url=String(cell.getValue()||'');
      if (/^https?:\/\/\S+$/i.test(url)) cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build());
    }
  }
}

function assignLeadId_(sheet,row) {
  const ids=sheet.getRange(BUSINESS_LEADS.firstDataRow,1,BUSINESS_LEADS.dataRows,1).getDisplayValues().flat();
  const next=ids.reduce((m,id)=>Math.max(m,Number((id.match(/^LEAD-(\d+)$/)||[])[1])||0),0)+1;
  sheet.getRange(row,1).setValue('LEAD-'+String(next).padStart(4,'0'));
}

function ensureBusinessLeadsEditTrigger_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const matches=ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction()==='businessLeadsEditHandler' && t.getTriggerSourceId()===ss.getId());
  matches.slice(1).forEach(t=>ScriptApp.deleteTrigger(t));
  if (!matches.length) ScriptApp.newTrigger('businessLeadsEditHandler').forSpreadsheet(ss).onEdit().create();
}

function ensureSheetSize_(sheet, rows, columns) {
  if (sheet.getMaxRows()<rows) sheet.insertRowsAfter(sheet.getMaxRows(),rows-sheet.getMaxRows());
  if (sheet.getMaxColumns()<columns) sheet.insertColumnsAfter(sheet.getMaxColumns(),columns-sheet.getMaxColumns());
}

// Optional DASHBOARD formulas can later reference the summary cells A3, U3, A6 and Q6.
// They are intentionally not written automatically, so no existing DASHBOARD layout is changed.

/**
 * Logins & Accounts extension for an existing productivity workbook.
 * Tracks login metadata only. It must never be used to store passwords, PINs,
 * recovery codes, API keys, or authentication secrets.
 */
const LOGINS_ACCOUNTS = Object.freeze({
  sheetName: 'Logins & Accounts',
  adminName: 'ADMIN',
  headerRow: 8,
  firstDataRow: 9,
  dataRows: 1000,
  columns: 19,
  headers: [
    'Account ID', 'Service / Website', 'Category', 'Login URL', 'Login Email',
    'Username', 'Recovery Email', 'Recovery Phone', 'Owner / Access',
    '2FA Enabled?', '2FA Method', 'Password Manager', 'Password Manager Entry',
    'Backup Codes Stored?', 'Status', 'Last Password Change', 'Last Reviewed',
    'Notes', 'Last Updated'
  ],
  colours: {
    navy: '#17324D', blue: '#486B8A', pale: '#EAF0F6', pale2: '#F5F8FB',
    green: '#C6EFCE', greenText: '#006100', amber: '#FFEB9C',
    amberText: '#9C6500', red: '#FFC7CE', redText: '#9C0006',
    grey: '#E7E6E6', greyText: '#666666', shared: '#BDD7EE',
    sharedText: '#1F4E78', white: '#FFFFFF'
  }
});

/** Creates or rebuilds only Logins & Accounts and its isolated ADMIN lists. */
function buildLoginsAccountsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = LOGINS_ACCOUNTS;
  const lists = ensureLoginsAccountsAdminLists_(ss);
  let sheet = ss.getSheetByName(c.sheetName);
  if (!sheet) sheet = ss.insertSheet(c.sheetName);

  // Deliberately clear only this feature's sheet; no other sheet is reset/deleted.
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  ensureLoginsAccountsSheetSize_(sheet, c.firstDataRow + c.dataRows - 1, 35);
  sheet.setHiddenGridlines(true).setFrozenRows(c.headerRow);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily('Arial').setVerticalAlignment('middle');

  buildLoginsAccountsSummary_(sheet);
  buildLoginsAccountsTable_(sheet, lists);
  seedLoginsAccountsSamples_(sheet);
  buildLoginsAccountsReviewSections_(sheet);
  applyLoginsAccountsConditionalFormatting_(sheet);
  sheet.getRange(c.headerRow, 1, c.dataRows + 1, c.columns).createFilter();
  ensureLoginsAccountsEditTrigger_();
  sheet.activate();
  SpreadsheetApp.flush();
}

/** Adds editable list sections without clearing or rearranging existing ADMIN data. */
function ensureLoginsAccountsAdminLists_(ss) {
  let admin = ss.getSheetByName(LOGINS_ACCOUNTS.adminName);
  if (!admin) admin = ss.insertSheet(LOGINS_ACCOUNTS.adminName);
  const definitions = [
    ['Account Categories', 'AccountCategoriesList', ['Business', 'Personal', 'Banking', 'Social Media', 'Software', 'University', 'Government', 'Utilities', 'Shopping', 'Other']],
    ['Account Owners', 'AccountOwnersList', ['Me', 'Business Partner', 'Shared', 'Admin', 'Staff', 'Other']],
    ['2FA Methods', 'TwoFAMethodsList', ['Authenticator App', 'SMS', 'Email', 'Security Key', 'Passkey', 'None']],
    ['Password Managers', 'PasswordManagersList', ['Bitwarden', '1Password', 'iCloud Keychain', 'Google Password Manager', 'Chrome Password Manager', 'Other', 'None']],
    ['Account Statuses', 'AccountStatusList', ['Active', 'Shared', 'Needs Review', 'Needs Password Change', 'Missing 2FA', 'Disabled']]
  ];
  const result = {};

  definitions.forEach(def => {
    const sameNames = ss.getNamedRanges().filter(named => named.getName() === def[1]);
    const namedOnAdmin = sameNames.map(named => named.getRange())
      .find(range => range.getSheet().getSheetId() === admin.getSheetId());
    const found = findLoginsAccountsAdminSection_(admin, def[0]);
    const column = found ? found.column : (namedOnAdmin ? namedOnAdmin.getColumn() : Math.max(1, admin.getLastColumn() + 2));
    const headerRow = found ? found.headerRow : (namedOnAdmin ? Math.max(1, namedOnAdmin.getRow() - 1) : 1);
    if (admin.getMaxColumns() < column) admin.insertColumnsAfter(admin.getMaxColumns(), column - admin.getMaxColumns());
    if (admin.getMaxRows() < headerRow + def[2].length) admin.insertRowsAfter(admin.getMaxRows(), headerRow + def[2].length - admin.getMaxRows());

    // Only clear the prior cells belonging to this same section/range.
    const oldLength = Math.max(found ? found.length : 0, namedOnAdmin ? namedOnAdmin.getNumRows() : 0, def[2].length);
    if (oldLength) admin.getRange(headerRow + 1, column, oldLength, 1).clearContent().clearDataValidations();
    admin.getRange(headerRow, column).setValue(def[0]).setFontFamily('Arial').setFontWeight('bold')
      .setBackground(LOGINS_ACCOUNTS.colours.blue).setFontColor(LOGINS_ACCOUNTS.colours.white);
    const range = admin.getRange(headerRow + 1, column, def[2].length, 1);
    range.setValues(def[2].map(value => [value])).setFontFamily('Arial');
    sameNames.forEach(named => named.remove());
    ss.setNamedRange(def[1], range);
    result[def[1]] = range;
  });
  return result;
}

function findLoginsAccountsAdminSection_(admin, title) {
  if (!admin.getLastRow() || !admin.getLastColumn()) return null;
  const values = admin.getDataRange().getDisplayValues();
  for (let row = 0; row < values.length; row++) {
    for (let column = 0; column < values[row].length; column++) {
      if (values[row][column].trim().toLowerCase() === title.toLowerCase()) {
        let length = 0;
        while (row + 1 + length < values.length && values[row + 1 + length][column] !== '') length++;
        return {headerRow: row + 1, column: column + 1, length: length};
      }
    }
  }
  return null;
}

function buildLoginsAccountsSummary_(sheet) {
  const c = LOGINS_ACCOUNTS, first = c.firstDataRow, last = first + c.dataRows - 1;
  const labels = [
    'Total Accounts', 'Active Accounts', 'Accounts With 2FA', 'Accounts Without 2FA',
    'Accounts Needing Review', 'Accounts Needing Password Change', 'Shared Accounts',
    'Disabled Accounts', '2FA Coverage'
  ];
  const formulas = [
    `=COUNTIF(B${first}:B${last},"<>")`,
    `=COUNTIF(O${first}:O${last},"Active")`,
    `=COUNTIFS(B${first}:B${last},"<>",J${first}:J${last},TRUE)`,
    `=COUNTIFS(B${first}:B${last},"<>",J${first}:J${last},FALSE,O${first}:O${last},"<>Disabled")`,
    `=COUNTIF(O${first}:O${last},"Needs Review")`,
    `=COUNTIF(O${first}:O${last},"Needs Password Change")`,
    `=COUNTIF(O${first}:O${last},"Shared")`,
    `=COUNTIF(O${first}:O${last},"Disabled")`,
    `=IFERROR(COUNTIFS(B${first}:B${last},"<>",J${first}:J${last},TRUE,O${first}:O${last},"Active")/COUNTIF(O${first}:O${last},"Active"),0)`
  ];
  sheet.getRange(1, 1, 1, 19).merge().setValue('LOGINS & ACCOUNTS')
    .setBackground(c.colours.navy).setFontColor(c.colours.white).setFontSize(15)
    .setFontWeight('bold').setHorizontalAlignment('center');
  labels.forEach((label, index) => {
    const column = 1 + (index % 5) * 4;
    const row = index < 5 ? 2 : 5;
    const width = index === 4 || index === 8 ? 3 : 3;
    sheet.getRange(row, column, 1, width).merge().setValue(label)
      .setBackground(c.colours.blue).setFontColor(c.colours.white)
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(row + 1, column, 1, width).merge().setFormula(formulas[index])
      .setBackground(c.colours.pale).setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center');
  });
  sheet.getRange(6, 13, 1, 3).setNumberFormat('0.0%');
}

function buildLoginsAccountsTable_(sheet, lists) {
  const c = LOGINS_ACCOUNTS;
  sheet.getRange(c.headerRow, 1, 1, c.columns).setValues([c.headers])
    .setBackground(c.colours.navy).setFontColor(c.colours.white)
    .setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  const body = sheet.getRange(c.firstDataRow, 1, c.dataRows, c.columns);
  body.setBackground(c.colours.pale2).setFontColor('#243746');
  for (let row = c.firstDataRow + 1; row < c.firstDataRow + c.dataRows; row += 2) {
    sheet.getRange(row, 1, 1, c.columns).setBackground(c.colours.pale);
  }
  [5, 6, 7, 8, 13].forEach(column => sheet.getRange(c.firstDataRow, column, c.dataRows, 1).setNumberFormat('@'));
  [16, 17].forEach(column => sheet.getRange(c.firstDataRow, column, c.dataRows, 1).setNumberFormat('dd/mm/yyyy'));
  sheet.getRange(c.firstDataRow, 19, c.dataRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(c.firstDataRow, 18, c.dataRows, 1).setWrap(true);
  sheet.getRange(c.firstDataRow, 10, c.dataRows, 1).insertCheckboxes();
  sheet.getRange(c.firstDataRow, 14, c.dataRows, 1).insertCheckboxes();

  const listRule = range => SpreadsheetApp.newDataValidation().requireValueInRange(range, true).setAllowInvalid(false).build();
  sheet.getRange(c.firstDataRow, 3, c.dataRows, 1).setDataValidation(listRule(lists.AccountCategoriesList));
  sheet.getRange(c.firstDataRow, 9, c.dataRows, 1).setDataValidation(listRule(lists.AccountOwnersList));
  sheet.getRange(c.firstDataRow, 11, c.dataRows, 1).setDataValidation(listRule(lists.TwoFAMethodsList));
  sheet.getRange(c.firstDataRow, 12, c.dataRows, 1).setDataValidation(listRule(lists.PasswordManagersList));
  sheet.getRange(c.firstDataRow, 15, c.dataRows, 1).setDataValidation(listRule(lists.AccountStatusList));
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sheet.getRange(c.firstDataRow, 16, c.dataRows, 2).setDataValidation(dateRule);

  sheet.setColumnWidths(1, c.columns, 125);
  sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 170); sheet.setColumnWidth(4, 210);
  sheet.setColumnWidth(5, 190); sheet.setColumnWidth(7, 190); sheet.setColumnWidth(8, 135);
  sheet.setColumnWidth(9, 145); sheet.setColumnWidth(11, 150); sheet.setColumnWidth(12, 185);
  sheet.setColumnWidth(13, 230); sheet.setColumnWidth(15, 175); sheet.setColumnWidth(18, 275);
  sheet.setColumnWidth(19, 150); sheet.setRowHeight(c.headerRow, 46);
}

function buildLoginsAccountsReviewSections_(sheet) {
  const c = LOGINS_ACCOUNTS, first = c.firstDataRow, last = first + c.dataRows - 1;
  sheet.getRange(1, 21, 1, 7).merge().setValue('SECURITY REVIEW')
    .setBackground(c.colours.blue).setFontColor(c.colours.white).setFontWeight('bold');
  sheet.getRange(2, 21, 1, 7).setValues([['Service / Website', 'Login Email', 'Owner / Access', '2FA Enabled?', 'Password Manager', 'Status', 'Last Reviewed']])
    .setBackground(c.colours.navy).setFontColor(c.colours.white).setFontWeight('bold').setWrap(true);
  const risk = `(N(O${first}:O${last}="Needs Password Change")*5+N(O${first}:O${last}="Missing 2FA")*5+N((J${first}:J${last}=FALSE)*(O${first}:O${last}<>"Disabled"))*4+N(L${first}:L${last}="None")*3+N(O${first}:O${last}="Needs Review")*2+N((Q${first}:Q${last}<>"")*(Q${first}:Q${last}<TODAY()-180))*1)`;
  sheet.getRange(3, 21).setFormula(`=IFERROR(QUERY(SORT(FILTER({B${first}:B${last},E${first}:E${last},I${first}:I${last},J${first}:J${last},L${first}:L${last},O${first}:O${last},Q${first}:Q${last},${risk}},B${first}:B${last}<>"",((J${first}:J${last}=FALSE)+(O${first}:O${last}="Needs Review")+(O${first}:O${last}="Needs Password Change")+((Q${first}:Q${last}<>"")*(Q${first}:Q${last}<TODAY()-180))+(L${first}:L${last}="None"))>0),8,FALSE),"select Col1,Col2,Col3,Col4,Col5,Col6,Col7",0),"")`);
  sheet.getRange(3, 27, c.dataRows, 1).setNumberFormat('dd/mm/yyyy');

  sheet.getRange(1, 29, 1, 6).merge().setValue('ACCOUNT REVIEW — OVER 180 DAYS')
    .setBackground(c.colours.blue).setFontColor(c.colours.white).setFontWeight('bold');
  sheet.getRange(2, 29, 1, 6).setValues([['Service / Website', 'Category', 'Login Email', 'Owner / Access', 'Last Reviewed', 'Status']])
    .setBackground(c.colours.navy).setFontColor(c.colours.white).setFontWeight('bold').setWrap(true);
  sheet.getRange(3, 29).setFormula(`=IFERROR(SORT(FILTER({B${first}:B${last},C${first}:C${last},E${first}:E${last},I${first}:I${last},Q${first}:Q${last},O${first}:O${last}},B${first}:B${last}<>"",Q${first}:Q${last}<>"",Q${first}:Q${last}<TODAY()-180),5,TRUE),"")`);
  sheet.getRange(3, 33, c.dataRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.setColumnWidths(21, 15, 145);
  sheet.setColumnWidth(21, 180); sheet.setColumnWidth(22, 190); sheet.setColumnWidth(25, 185);
  sheet.setColumnWidth(29, 180); sheet.setColumnWidth(31, 190);
}

function applyLoginsAccountsConditionalFormatting_(sheet) {
  const c = LOGINS_ACCOUNTS, first = c.firstDataRow, count = c.dataRows;
  const rules = [], twoFA = sheet.getRange(first, 10, count, 1);
  const manager = sheet.getRange(first, 12, count, 1), backup = sheet.getRange(first, 14, count, 1);
  const status = sheet.getRange(first, 15, count, 1), reviewed = sheet.getRange(first, 17, count, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($B${first}<>"",$J${first}=FALSE,$O${first}<>"Disabled")`).setBackground(c.colours.red).setFontColor(c.colours.redText).setRanges([twoFA]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($B${first}<>"",$J${first}=TRUE)`).setBackground(c.colours.green).setFontColor(c.colours.greenText).setRanges([twoFA]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($Q${first}<>"",$Q${first}<TODAY()-365)`).setBackground(c.colours.red).setFontColor(c.colours.redText).setRanges([reviewed]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($Q${first}<>"",$Q${first}<TODAY()-180)`).setBackground(c.colours.amber).setFontColor(c.colours.amberText).setRanges([reviewed]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('None').setBackground(c.colours.red).setFontColor(c.colours.redText).setRanges([manager]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($B${first}<>"",$J${first}=TRUE,$N${first}=FALSE)`).setBackground(c.colours.amber).setFontColor(c.colours.amberText).setRanges([backup]).build());
  const statusColours = {
    'Active': [c.colours.green, c.colours.greenText], 'Shared': [c.colours.shared, c.colours.sharedText],
    'Needs Review': [c.colours.amber, c.colours.amberText],
    'Needs Password Change': [c.colours.red, c.colours.redText],
    'Missing 2FA': [c.colours.red, c.colours.redText], 'Disabled': [c.colours.grey, c.colours.greyText]
  };
  Object.keys(statusColours).forEach(value => rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(value).setBackground(statusColours[value][0]).setFontColor(statusColours[value][1])
    .setRanges([status]).build()));
  sheet.setConditionalFormatRules(rules);
}

function seedLoginsAccountsSamples_(sheet) {
  const today = new Date(), ago = days => new Date(today.getFullYear(), today.getMonth(), today.getDate() - days);
  const data = [
    ['ACC-0001', 'Harbour Business Mail', 'Business', 'https://example.invalid/business-mail', 'admin@harbour.example.invalid', 'harbour.admin', 'recovery@harbour.example.invalid', '0400000001', 'Admin', true, 'Authenticator App', 'Bitwarden', 'Bitwarden – Business Gmail', true, 'Active', ago(80), ago(20), 'Fictional primary business email account.', today],
    ['ACC-0002', 'Koala Photo Sharing', 'Personal', 'https://example.invalid/koala-social', 'alex@personal.example.invalid', 'alex.koala', 'backup@personal.example.invalid', '0400000002', 'Me', true, 'Passkey', 'iCloud Keychain', 'iCloud Keychain – Koala Photo Sharing', false, 'Active', ago(120), ago(45), 'Fictional personal social profile.', today],
    ['ACC-0003', 'Southern Cross Bank Demo', 'Banking', 'https://example.invalid/demo-bank', 'accounts@harbour.example.invalid', 'demo.customer', 'recovery@harbour.example.invalid', '0400000003', 'Me', true, 'Security Key', '1Password', '1Password – Demo Bank Login', true, 'Active', ago(35), ago(15), 'Fictional demonstration bank account; contains no financial credentials.', today],
    ['ACC-0004', 'Wattle Social', 'Social Media', 'https://example.invalid/wattle-social', 'social@harbour.example.invalid', 'harbour.social', 'admin@harbour.example.invalid', '0400000004', 'Shared', false, 'None', 'Bitwarden', 'Bitwarden – Wattle Social', false, 'Missing 2FA', ago(390), ago(210), 'Fictional shared marketing account; enable 2FA.', today],
    ['ACC-0005', 'Cloud Ledger Demo', 'Software', 'https://example.invalid/cloud-ledger', 'finance@harbour.example.invalid', 'finance.admin', 'admin@harbour.example.invalid', '0400000005', 'Business Partner', true, 'Authenticator App', '1Password', '1Password – Cloud Ledger Admin', true, 'Shared', ago(95), ago(70), 'Fictional bookkeeping software subscription.', today],
    ['ACC-0006', 'Campus Learning Portal', 'University', 'https://example.invalid/campus', 'student@campus.example.invalid', 'student0006', 'alex@personal.example.invalid', '0400000006', 'Me', false, 'None', 'None', 'No manager entry recorded', false, 'Needs Review', ago(500), ago(400), 'Fictional study account; review security and password-manager storage.', today],
    ['ACC-0007', 'Design Studio Demo', 'Software', 'https://example.invalid/design-studio', 'design@harbour.example.invalid', 'harbour.design', 'admin@harbour.example.invalid', '0400000007', 'Staff', true, 'SMS', 'Google Password Manager', 'Google Password Manager – Design Studio', false, 'Needs Password Change', ago(420), ago(190), 'Fictional design subscription used by the business.', today],
    ['ACC-0008', 'Old Utilities Portal', 'Utilities', 'https://example.invalid/old-utilities', 'old-account@harbour.example.invalid', 'old.utility', 'admin@harbour.example.invalid', '0400000008', 'Admin', false, 'None', 'None', 'No manager entry required – disabled', false, 'Disabled', ago(700), ago(380), 'Fictional closed utility account retained for ownership records.', today]
  ];
  sheet.getRange(LOGINS_ACCOUNTS.firstDataRow, 1, data.length, LOGINS_ACCOUNTS.columns).setValues(data);
}

/** Menu command: selects the first prepared row with no Service / Website. */
function addNewAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), c = LOGINS_ACCOUNTS;
  const sheet = ss.getSheetByName(c.sheetName);
  if (!sheet) throw new Error('Run buildLoginsAccountsSheet() first.');
  const services = sheet.getRange(c.firstDataRow, 2, c.dataRows, 1).getDisplayValues();
  const offset = services.findIndex(row => row[0] === '');
  if (offset < 0) throw new Error('The 1,000 prepared account rows are full.');
  const row = c.firstDataRow + offset;
  // The row was preformatted and prevalidated by the builder; change values only.
  if (!sheet.getRange(row, 1).getValue()) assignLoginsAccountId_(sheet, row);
  sheet.getRange(row, 15).setValue('Active');
  sheet.activate();
  sheet.getRange(row, 2).activate();
}

/** Menu command: opens Logins & Accounts. */
function openLoginsAccountsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOGINS_ACCOUNTS.sheetName);
  if (!sheet) throw new Error('Run buildLoginsAccountsSheet() first.');
  sheet.activate();
}

/**
 * Add this helper to the existing Productivity System menu builder before its
 * addToUi() call. Apps Script cannot safely mutate an already-rendered menu.
 */
function addLoginsAccountsMenu_(menu) {
  return menu.addItem('Add New Account', 'addNewAccount')
    .addItem('Open Logins & Accounts', 'openLoginsAccountsSheet');
}

/** Installable edit handler scoped strictly to the Logins & Accounts data rows. */
function loginsAccountsEditHandler(e) {
  if (!e || !e.range) return;
  const range = e.range, sheet = range.getSheet(), c = LOGINS_ACCOUNTS;
  if (sheet.getName() !== c.sheetName || range.getRow() < c.firstDataRow ||
      range.getRow() >= c.firstDataRow + c.dataRows || range.getColumn() > c.columns) return;
  const startRow = range.getRow(), endRow = Math.min(range.getLastRow(), c.firstDataRow + c.dataRows - 1);
  for (let row = startRow; row <= endRow; row++) {
    sheet.getRange(row, 19).setValue(new Date());
    if (!sheet.getRange(row, 1).getValue() && sheet.getRange(row, 2).getValue()) assignLoginsAccountId_(sheet, row);
  }
  if (range.getColumn() <= 4 && range.getLastColumn() >= 4) {
    for (let row = startRow; row <= endRow; row++) {
      const cell = sheet.getRange(row, 4), url = String(cell.getValue() || '');
      if (/^https?:\/\/\S+$/i.test(url)) {
        cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build());
      }
    }
  }
}

function assignLoginsAccountId_(sheet, row) {
  const c = LOGINS_ACCOUNTS;
  const ids = sheet.getRange(c.firstDataRow, 1, c.dataRows, 1).getDisplayValues().flat();
  const next = ids.reduce((maximum, id) => Math.max(maximum, Number((id.match(/^ACC-(\d+)$/) || [])[1]) || 0), 0) + 1;
  sheet.getRange(row, 1).setValue('ACC-' + String(next).padStart(4, '0'));
}

function ensureLoginsAccountsEditTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matches = ScriptApp.getProjectTriggers().filter(trigger =>
    trigger.getHandlerFunction() === 'loginsAccountsEditHandler' && trigger.getTriggerSourceId() === ss.getId());
  matches.slice(1).forEach(trigger => ScriptApp.deleteTrigger(trigger));
  if (!matches.length) ScriptApp.newTrigger('loginsAccountsEditHandler').forSpreadsheet(ss).onEdit().create();
}

function ensureLoginsAccountsSheetSize_(sheet, rows, columns) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < columns) sheet.insertColumnsAfter(sheet.getMaxColumns(), columns - sheet.getMaxColumns());
}

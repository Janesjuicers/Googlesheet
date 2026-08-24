# Business Leads & Clients — installation

1. Open the existing spreadsheet and choose **Extensions → Apps Script**.
2. Add a new script file named `BusinessLeads.gs` and paste the complete contents of [`BusinessLeads.gs`](BusinessLeads.gs).
3. Save the project. Do **not** replace or delete any existing script files.
4. In the function selector choose `buildBusinessLeadsSheet`, click **Run**, and approve the requested spreadsheet and trigger permissions. This is the only build function that needs to be run.
5. To retain the two commands alongside an existing **Productivity System** menu after every reopen, add these calls to the existing menu chain immediately before its `.addToUi()` call:

   ```javascript
   .addItem('Add New Business Lead', 'addNewBusinessLead')
   .addItem('Open Business Leads & Clients', 'openBusinessLeadsSheet')
   ```

   Alternatively, if the existing code stores its menu in a variable, pass it through `addBusinessLeadsMenu_(menu)` before calling `menu.addToUi()`. Google Apps Script cannot retrieve and mutate a menu after it has been added, so merging these two items into the existing menu builder is the safe way to preserve all current commands.

The builder never deletes a sheet. It clears/rebuilds only `Business Leads & Clients`, appends isolated list columns to `ADMIN`, replaces only same-name named-range definitions, and deliberately leaves `DASHBOARD` and every other existing sheet untouched.

## Logins & Accounts — installation

1. In the existing spreadsheet, choose **Extensions → Apps Script**.
2. Create a **new script file** named `LoginsAccounts.gs` (recommended) and paste the complete contents of [`LoginsAccounts.gs`](LoginsAccounts.gs). A separate `.gs` file keeps the extension maintainable; placing the code underneath an existing script also works, but do not replace or delete existing code.
3. Save the Apps Script project. In the function selector choose **`buildLoginsAccountsSheet`**, click **Run**, and approve the requested spreadsheet and trigger permissions. This is the only function to run manually.
4. To preserve the existing **Productivity System** menu and append the new commands, add the following two items to its existing menu chain immediately before `.addToUi()`:

   ```javascript
   .addItem('Add New Account', 'addNewAccount')
   .addItem('Open Logins & Accounts', 'openLoginsAccountsSheet')
   ```

   If the existing menu is stored in a variable, use `addLoginsAccountsMenu_(menu);` before `menu.addToUi()`. Apps Script cannot retrieve and mutate an existing menu after it is rendered, so this small addition to the existing builder preserves every current command instead of replacing the menu.

The builder clears/rebuilds only `Logins & Accounts`. It never deletes a sheet or clears data outside that sheet, and it updates only its five isolated `ADMIN` sections and same-name named ranges. The sheet records password-manager entry names/locations only; it includes no field or functionality for storing passwords or other authentication secrets.

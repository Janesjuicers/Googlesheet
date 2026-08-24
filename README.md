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

# Project Pickle Technical Notes

Application version: 2.2.1

These notes are for maintainers, reviewers, and developers. The operator-facing workflow belongs in `Project Pickle How-To Guide.html`; the repository overview belongs in the root `README.md`; and the focused in-app AI reference belongs in `Project Pickle App Files/AI-HELP-CONTEXT.md`.

## 1. Source of Truth and Documentation Roles

Use the following order when behavior and documentation differ:

1. The current released `ProjectPickle.ps1` implementation and a reproducible test result.
2. The current operator how-to guide.
3. These technical notes.
4. The focused AI Help context.
5. Historical README files or release notes.

Update all version references and workflow names in the same release. Version 2.2.1 retains the 2.1.8 feature set and fixes standard AppColumns visibility filtering in the normal Get Field Order and CFIDEER rebuild paths. Editable fields from `PlotMeasurements`, `TreeMeasurements`, and `RegenMeasurements` now follow their current `Active + QueryVisible` values just like custom fields. After the AppColumns setup has been loaded or reviewed, saved-query candidates, physical-table fallback candidates, and built-in default candidates are filtered before display, so an inactive or QueryVisible-off field such as `HabitatType` or `DensityClass` cannot be reintroduced by a fallback source. Required system fields remain protected, and exact query-copy mode remains unchanged.

The 2.1.8 baseline includes the 2.1.7 source-integrity and preservation safeguards, read-only source previews for AppColumns, copied CFIDEER Get fields, completed Reports, the portable root/internal launcher pair, and the wiggling-pickle splash startup. It also includes source inventory preflight, required Upgrade Mapping review, source-aware AppColumns physical-field copy, copied custom-table Design View order preservation, Access Description preservation, Species Merchantability identifier normalization, strengthened CFIDEER verification, and Application Title persistence checks. Version 2.2.1 includes the 2.1.7 source-integrity and preservation safeguards, read-only source previews for AppColumns, copied CFIDEER Get fields, and completed Reports, plus the portable root/internal launcher pair and the wiggling-pickle splash startup. Version 2.2.1 includes source inventory preflight, required Upgrade Mapping review, source-aware AppColumns physical-field copy, copied custom-table Design View order preservation, Access Description preservation, Species Merchantability identifier normalization, strengthened CFIDEER verification, and Application Title persistence checks.

## 2. Runtime Layout

The expected release layout is:

```text
ProjectPickle/
|-- Run Project Pickle.vbs
|-- Run-ProjectPickle-32bit.bat
|-- Project Pickle How-To Guide.html
`-- Project Pickle App Files/
    |-- AI-HELP-CONTEXT.md
    |-- ProjectPickle.ps1
    |-- ProjectPickle.vbs
    |-- Run-ProjectPickle-32bit.bat
    `-- ProjectPickle.ico
```

The root launchers provide a simple user entry point. The internal launchers and source script are support files. Keep this structure intact because launchers and bundled assets are resolved by relative path.

`$script:AppRoot` is derived from `$PSScriptRoot`, then `$PSCommandPath`, then the current directory. Bundled paths such as the icon, AI context, source-code context, and testing artifacts are resolved from that application root.

The current root screenshot shows `Run Project Pickle.vbs`, while older documentation may say `Project Pickle.vbs`. Select one public launcher name and make the root file, README, HTML guide, and release notes identical before release.

The owner-drawn main tab strip uses a pink accent (`#E91E63`) for Processing and a red accent (`#F44336`) for Build, with white foreground text for both strong-colored tabs. These are presentation-only changes and do not affect mode or build logic.

## 3. Application Entry Point and Parameters

`ProjectPickle.ps1` is the primary application entry point. Important parameters include:

| Parameter | Purpose |
| --- | --- |
| `-SelfTest` | Runs the built-in builder self-test without opening the normal UI. |
| `-UiSmokeTest` | Opens the UI, exercises selected startup behavior, then closes automatically. |
| `-ReferenceListTest` | Runs reference-list checks. |
| `-NoUi` | Reserved for scripted builds; current behavior prints a message and exits. |
| `-MasterDatabase` | Supplies a database path for tests or scripted startup behavior. |
| `-UpgradeTemplateDatabase` | Supplies a current master database path. |
| `-OutputDatabase` | Supplies an output database path. |
| `-SplashReadyFile` | Supplies the temporary launcher-created marker file. The main application deletes the marker when the primary form is ready, signaling the splash process to close. |
| `-AiEndpoint`, `-AiModel`, `-AiApiKey` | Supply AI Help values at startup. |

Top-level failures are written to `ProjectPickle.LastError.txt` and `ProjectPickle.Diagnostics.txt` near the root application folder when possible.

### Launcher And Splash Startup

The normal startup path is:

```text
Run Project Pickle.vbs
    -> Project Pickle App Files\ProjectPickle.vbs
        -> ProjectPickleSplash.ps1
        -> ProjectPickle.ps1 -SplashReadyFile <temporary marker>
```

The internal VBS prefers:

```text
%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe
```

so the normal process matches a 32-bit Access provider. It can fall back to the System32 PowerShell path when the 32-bit executable is unavailable.

The internal launcher:

1. Resolves its own folder.
2. Locates `ProjectPickle.ps1`.
3. Locates `ProjectPickleSplash.ps1`.
4. Creates a uniquely named temporary ready marker.
5. Starts the splash in a hidden STA PowerShell process.
6. Starts the main application in a separate hidden STA PowerShell process.
7. Passes the ready-marker path through `-SplashReadyFile`.

`Set-ProjectPickleSplashReady` deletes the marker when the main application is ready. The splash timer watches for that deletion and closes its form.

The splash also has a safety timeout so a splash process cannot remain visible indefinitely after a startup failure.

`ProjectPickleSplash.ps1`:

- Reads the application version from `ProjectPickle.ps1`.
- Uses `ProjectPickleSplash.png` when a valid PNG exists.
- Draws a built-in pickle when the PNG is absent or invalid.
- Pre-creates rotated image frames.
- Switches frames with a Windows Forms timer.
- Animates the Loading dots.
- Does not use `ProjectPickle.ico` as enlarged splash artwork because some ICO encodings can be decoded incorrectly by Windows PowerShell 5.1 and GDI+.
- Must never prevent the main application from starting.

The main `ProjectPickle.ico` remains the application, taskbar, and window icon.

## 4. Runtime Dependencies

Project Pickle is Windows-only because it uses Windows Forms, Windows shell APIs, and OLE DB access to Microsoft Access databases.

The application probes Access providers in this order:

For `.mdb` files:

1. `Microsoft.Jet.OLEDB.4.0`
2. `Microsoft.ACE.OLEDB.16.0`
3. `Microsoft.ACE.OLEDB.15.0`
4. `Microsoft.ACE.OLEDB.12.0`

For other Access files, including `.accdb`:

1. `Microsoft.ACE.OLEDB.16.0`
2. `Microsoft.ACE.OLEDB.15.0`
3. `Microsoft.ACE.OLEDB.12.0`
4. `Microsoft.Jet.OLEDB.4.0`

Provider bitness must match the PowerShell process. The 32-bit launcher is the normal recovery path when only a 32-bit Access Database Engine is installed.

Before a public release, record the exact tested combinations of Windows, Microsoft Access or Access Database Engine, provider version, database format, and process bitness.

## 5. Database Modes and Copy Source

| Mode | Base file copied to output | Additional source |
| --- | --- | --- |
| New Project | Current master database version | None |
| Upgrade Project | Current master database version | Old project database supplies imported data and optional completed setup tables |
| Existing Project | Current-format existing project database | None |

`Assert-BuilderConfigReady` validates required mode inputs, output extension, output directory, period data, AppColumns rows, report layouts, plot mappings, and mode-specific copy options before the build starts.

The output path must differ from both input paths. When an output file already exists, `Invoke-InventoryBuild` moves it to:

```text
<output-file>.backup_yyyyMMdd_HHmmss
```

The selected base file is then copied to the output path, and the output copy is opened for modification.

This is not one database-wide transaction. Recovery depends on the untouched input database, the timestamped prior-output backup, and targeted transactions inside selected operations.

### Read-Only Copy Preview Architecture

The preview systems are UI-only. They do not write to the source or output database.

All three previews follow the same lifecycle:

1. Save the current manual UI state.
2. Read the Existing/old source.
3. Populate the corresponding controls.
4. Mark the controls read-only.
5. Leave scrolling and inspection available.
6. Disable controls that would alter the preview.
7. Restore the prior manual state when the option is unchecked.
8. Turn the option off when a dependent database path changes.

#### AppColumns Preview

Important state and helpers include:

```text
$appColumnsCopyPreviewState
Load-AppColumnsCopyPreview
Refresh-AppColumnsCopyPreviewDisplay
Restore-AppColumnsManualPreviewState
New-AppColumnsPreviewItemCopies
```

The display can contain both:

```text
Existing - WILL COPY
Current master - reference
```

Only the Existing/old rows populate the working copied AppColumns list.

Current-master comparison rows are display-only and must never be read by `Populate-GetQueryOrderGrids` or another build configuration path.

#### Copied CFIDEER Preview

Important state and helpers include:

```text
$getOrderCopyPreviewState
Load-CopiedCfideerGetOrderPreview
Get-GetOrderGridSnapshot
Set-GetOrderGridSnapshot
Restore-ManualGetOrderPreview
```

The preview loads only the three Get-query SELECT lists because Get Field Order represents those lists.

The complete query-copy Build path still copies and verifies Get, Update, and Process definitions.

Checking query copy must immediately:

```text
queryCheck.Checked = false
queryCheck.Enabled = false
```

Unchecking query copy restores the manual field order and returns Upgrade mode to the required query-rebuild path.

#### Completed Reports Preview

Important state and helpers include:

```text
$reportCopyPreviewState
Copy-ProjectPickleReportLayoutDictionary
Load-ProjectPickleCopiedReportsPreview
Restore-ProjectPickleManualReportPreviewState
```

The preview reads:

```text
ReportColumns
ReportHeaders
AppColumns
```

`AppTables` is optional because `Get-ReportLayoutFromConnection` contains a compatibility fallback.

The visible preview covers the three manually supported Plot, Tree, and Regen layouts. The Build path copies the complete report tables, including rows outside those layouts.

The preview preserves:

- Exact ReportColumns order
- FieldFormat
- TallyFormat
- Fixed-width ReportHeaders text
- Repeated spaces
- Blank header lines
- Blank Regen group values

The fake-data preview remains available because it does not alter the stored layout.

#### Event Registration

Local UI event handlers should be registered after the local helper functions and required state objects are defined.

Registering a closure before its local helper functions or state object exists can cause a checkbox to appear unclickable, immediately revert, or fail before its preview loads.

## 6. High-Level Build Pipeline

`Invoke-InventoryBuild` performs the main build in this approximate order:

1. Validate mode, paths, periods, selected copy options, report layouts, Processing compatibility, and the required CFIDEER query action.
2. Run the Existing/old inventory source preflight.
3. Block blank or duplicate IDs, row-count differences, missing/orphan custom rows, and invalid PlotID/TreeID relationships.
4. Collect explicit user approval for valid repeated Regen keys and matched-ID Regen key differences.
5. Require a complete Upgrade Mapping review for the exact selected source/master pair.
6. Preview copied-AppColumns physical-field actions.
7. Validate copied-query source definitions when exact query copy is selected.
8. Move a prior output to a timestamped backup.
9. Copy the mode-appropriate base database to the output path.
10. In New and Upgrade modes, write and verify the Access `AppTitle` from Project name.
11. Create or refresh `AppColumnsArchive`.
12. Optionally archive old `AppColumns` and `AppColumnCodes` to `zAppColumns` and `zAppColumnsCodes`.
13. Resolve Region, Agency, and Reservation and write project setup.
14. Write periods for New/Upgrade or preserve periods in Existing mode.
15. Update Analysis table period IDs and prune reference rows where configured.
16. Copy working AppColumns metadata or synchronize the manual AppColumns setup.
17. Remove unusable unattached copied `AppColumnCodes` from working metadata.
18. Create only source-backed physical custom fields.
19. Remove physical and query references to `FieldNameSampleRemove`.
20. Import old inventory and apply moved-field and Upgrade Mapping operations.
21. Skip expensive legacy repair passes during normal production builds.
22. Verify source/output row counts, IDs, keys, and parent references.
23. Copy Processing tables and optional calculation/summary tables.
24. Normalize available `SpeciesMerchantabilityStandards.ProjectID` and `BIAProjectID`.
25. Copy and verify Access table/field Description properties.
26. Preserve Existing/old custom-table Design View order in copied-AppColumns mode, or apply key/active/history ordering in manual mode.
27. Update validation periods.
28. Import new plots and create required blank plot-period measurement rows.
29. Rebuild selected assignments.
30. Propagate and verify required project identifiers.
31. Copy completed Reports, apply manual layouts, or run automatic report updates.
32. Copy or rebuild `_prjCFIDEER_` queries.
33. Reopen and verify saved query SQL, required fields, custom fields, and placeholder exclusions.
34. Verify the Access Application Title again.
35. Finalize logs and optionally open the saved database.

Maintain this order carefully. Later operations depend on schema, fields, IDs, periods, mapped history fields, Processing tables, and copied metadata prepared earlier.

The full build is not one database-wide transaction. A Failed output can contain partial work and must not be treated as complete.

## 7. Database Access and Write Patterns

Database access uses `System.Data.OleDb`. Core helper groups include:

- Provider and connection helpers: `Get-ProviderCandidates`, `New-AccessConnectionString`, `Open-AccessConnection`
- Schema helpers: `Get-UserTables`, `Get-TableColumns`, `Get-TableColumnInfo`, `Test-TableExists`, `Test-ColumnExists`
- Command helpers: `Invoke-DbNonQuery`, `Get-DbScalar`, `Get-DbTable`
- Parameter helpers: `New-DbParam`, `DbText`, `DbGuid`, `DbInt`, `DbDouble`, `DbDate`, `DbBool`, and parameterized command functions

Use parameterized commands for data values whenever possible. Keep identifier quoting and Access SQL differences isolated in helper functions.

Complete-table copy operations for AppColumns, reports, and Processing use targeted transactions so a failed replacement does not intentionally leave a partially replaced target table set. The full build remains a sequence of operations rather than one transaction.

`Copy-AccessTableFromExternalDatabase` must attach the active target transaction to every target-side `OleDbCommand`, including schema-reading commands. Access rejects `ExecuteReader`, `Fill`, or `ExecuteNonQuery` operations on a connection with a pending local transaction when the command's `Transaction` property has not been initialized. Do not call a general target-table schema helper inside a pending transaction unless that helper accepts and assigns the transaction.

## 8. AppColumns, Archives, and Custom Measurement Tables

Important tables include:

- `AppColumns`
- `AppColumnCodes`
- `AppColumnsArchive`
- `zAppColumns`
- `zAppColumnsCodes`
- `PlotCustomMeasurements`
- `TreeCustomMeasurements`
- `RegenCustomMeasurements`

### Base AppColumns Archive

`New-AppColumnsArchive` snapshots the output copy's original `AppColumns` table before working AppColumns changes begin.

In Upgrade mode, this normally represents the current master database version's original AppColumns setup.

### Old Project AppColumns Archives

`Copy-ProjectPickleOldAppColumnsArchives` creates:

```text
zAppColumns
zAppColumnsCodes
```

These are reference archives and do not become the working AppColumns setup.

The archive can retain historical source rows that are intentionally omitted from the new working metadata.

### Working AppColumns Copy

`Copy-AppColumnsTablesFromExistingDatabase` replaces the working:

```text
AppColumns
AppColumnCodes
```

The copy is transactional and verifies:

- Source and target table presence
- Row counts
- AppColumns lookup relationships
- AppColumnCodes parent relationships
- AppTable, data-type, and category-type IDs
- Duplicate table/column metadata combinations

Legacy `AppColumnCodes` rows whose `ColumnID` has no matching copied `AppColumns.ID` cannot be used by the working setup. They are omitted from the working output and logged. The Existing/old database is not changed.

### Physical Custom-Field Policy

The source-aware physical-field planner uses these outcomes:

| Target field exists | Source field exists | Copied Active | Action |
| --- | --- | --- | --- |
| Yes | Any | Any | Keep the existing target field |
| No | Yes | Yes or No | Add the field using the source physical type |
| No | No | No | Keep the AppColumns row metadata-only |
| No | No | Yes | Block the build |

Important functions include:

```text
Get-CopiedAppColumnsPhysicalFieldAction
Get-CopiedAppColumnsPhysicalFieldPlan
Ensure-CustomMeasurementColumnsForCopiedAppColumns
```

This prevents inactive catalog-only metadata from materializing unwanted fields such as a field that was never physically part of the project.

### AppColumns Copy Preview

Selecting complete AppColumns copy loads the Existing/old source into a read-only comparison before Build.

The preview dropdown supports:

```text
Existing AppColumns - WILL COPY
Current master AppColumns - reference
Both sources
```

The Preview, Show, Data, and Search controls affect only what the user sees.

The working copied list remains the Existing/old list. Current-master rows must not be added to:

```text
$script:BuilderLoadedAppColumnItems
```

when that variable is later used for query rebuilding.

The manual AppColumns state is copied before preview mode and restored when preview mode is cleared.

Changing the Existing/old or current-master path invalidates the preview and clears the copy checkbox.

### Custom-Table Design View Order

`Set-CustomMeasurementTableColumnOrder` has two modes.

Manual AppColumns mode uses:

1. Required key fields
2. Active AppColumns fields
3. Inactive/history fields

Copied-AppColumns mode preserves the Existing/old field order for:

```text
PlotCustomMeasurements
TreeCustomMeasurements
RegenCustomMeasurements
```

Fields that exist only in the current master/output are appended after the preserved source order. DAO `OrdinalPosition` values are written and reopened for verification.

### Access Description Properties

`Copy-UpgradeTableAndFieldDescriptionsFromExistingDatabase` copies nonblank DAO `Description` properties from the Existing/old database.

It supports:

- Same-name tables and fields
- Explicit Upgrade Mapping destinations
- Known renamed-field destinations
- Paired measurement/custom tables
- Table-level descriptions
- Field-level descriptions
- Verification after write

A blank source description does not erase an existing current-master description.

Nonblank descriptions with no physical output target are logged for review.

## 9. Upgrade Import, Source Preflight, and Field Mapping

### Source Inventory Preflight

Normal Upgrade builds validate the Existing/old inventory before output work begins.

Fatal invariants include:

```text
Plots.PlotID is nonblank and unique
Trees.TreeID is nonblank and unique
Trees.PlotID -> Plots.PlotID
PlotMeasurements.MeasurementID is nonblank and unique
TreeMeasurements.MeasurementID is nonblank and unique
RegenMeasurements.MeasurementID is nonblank and unique
Each measurement table and custom table has the same total row count
Each measurement row has exactly one custom row by MeasurementID
Each custom row has exactly one measurement row by MeasurementID
PlotMeasurements.PlotID -> Plots.PlotID
TreeMeasurements.TreeID -> Trees.TreeID
RegenMeasurements.PlotID -> Plots.PlotID
Plot and Tree matched-ID measurement keys agree
```

Matching total counts are necessary but not sufficient. Equal counts can still hide one missing record and one unrelated orphan record, so direct MeasurementID coverage is also required.

### Regen Legacy Exceptions

Repeated `RegenMeasKey` values can be valid.

Matched rows can also contain different mirrored `RegenMeasKey` values while still being correctly related by `MeasurementID`.

The UI can collect explicit approval for:

- Repeated Regen key groups
- Matched-ID Regen key text differences

Approval does not bypass:

- Duplicate or blank MeasurementID
- Row-count differences
- Missing custom rows
- Orphan custom rows
- Invalid PlotID parents

The approval signature records the current key values and counts. Backend validation recalculates the signature so a modified source cannot reuse stale approval.

Important functions include:

```text
Confirm-SourceInventoryPreflightBeforeBuild
Get-ProjectPickleDuplicateRegenKeyReview
Get-ProjectPickleDuplicateRegenKeySignature
Assert-ProjectPickleInventoryIntegrity
Assert-UpgradeSourceIdentityAndReferencesPreserved
```

### Required Upgrade Mapping Review

Important functions include:

```text
Get-UpgradeMappingReviewKey
Get-UpgradeMappingReviewCoverage
Get-UpgradeMappingCustomTargetsForConfig
Confirm-UpgradeMappingReviewBeforeBuild
```

The review is tied to the exact source/master database paths.

Validation detects:

- Missing rows
- Extra/stale rows
- Duplicate review rows
- Invalid rows
- A different source/master pair
- Unconfirmed Skip rows

When the review is missing or stale, the UI refreshes the grid and stops before output database work.

`UpgradeFieldMappings` contains actual copy destinations. `UpgradeMappingReviewRows` contains the complete reviewed list, including intentional Skip rows.

### Import Policy

Normal production upgrades require a clean source and do not silently repair missing parent or custom rows.

Legacy repair routines remain available to `-SelfTest`, where malformed rows are created intentionally to verify compatibility behavior.

Normal production builds still retain all post-import verification:

- Per-table row-count checks
- Period-count comparisons
- Source/output identifier preservation
- Parent-reference preservation
- Measurement/custom coverage
- Final inventory integrity

### Moved And Mapped Fields

The upgrade pipeline attempts bulk Access operations for compatible moved fields and falls back to row handling where required.

Upgrade Mapping can:

- Copy a field to the same custom table
- Move a field between a measurement table and its paired custom table
- Map to a renamed target field
- Create a missing target custom field
- Preserve an old historical field without activating it in AppColumns

Datatype review must stop unsafe conversions before import. Text should not be forced into current numeric, Date, Boolean, or GUID targets merely because a small sample appears convertible.

## 10. Plot Import and Period Rows

`Import-TabularRows` and related Open XML helpers read CSV, XLSX, and XLSM inputs without automating Excel.

Plot import is additive by PlotNumber. Existing plot numbers are skipped. New plots can receive blank `PlotMeasurements` and matching `PlotCustomMeasurements` rows for applicable periods.

`Ensure-BlankPlotMeasurementsForPeriods` is also used after upgrade and period setup to fill missing plot-period combinations.

Spreadsheet input should be treated as untrusted user data. Preserve validation for missing PlotNumber values, malformed field mappings, dates, numbers, and identifiers.

## 11. Reports

The manual report designer supports selected Plot, Tree, and Regen layouts stored in `ReportColumns` and `ReportHeaders`.

Fixed-width header text is significant data. Preserve repeated spaces and blank lines. Validation checks include:

- Missing or duplicate report fields
- Missing FieldFormat values
- Missing required remarks fields
- Custom fields after required remarks fields
- Tabs or unsupported characters
- Lines longer than 255 characters
- Unequal nonblank line lengths

Saved manual layouts are held in memory until Build. Closing the app loses unsaved layout work.

### Completed Reports Copy Preview

In Upgrade mode, selecting completed report copy immediately reads the source report tables and displays the three supported layouts.

The UI remains usable for inspection:

- Report selector enabled
- ReportColumns grid enabled and read-only
- Header editor enabled and read-only
- Scrolling available
- Text selection available
- Fake-data preview available

Editing operations remain disabled:

- Load
- Save
- Validate
- Add ReportVisible fields
- Move
- Remove
- Generate headers
- Edit labels
- Edit FieldFormat/TallyFormat

The preview status displays total source `ReportColumns` and `ReportHeaders` row counts.

The visible three-layout count can differ from the total source-table row count because Build copies all source rows.

`Copy-ProjectPickleReportLayoutDictionary` must create independent layout copies when saving and restoring manual state. Preview objects must not mutate the prior manual layouts.

Changing the report selector while preview mode is active only changes the displayed layout. It must not call the manual layout-save path.

Changing the source database path turns completed-report copy off and restores the saved manual state.

### RemarksTree FieldFormat Protection

`TreeMeasurements.RemarksTree` is required to remain:

```text
TEXT,40,5
```

## 12. CFIDEER Query Generation, Copy, and Verification

Upgrade Project requires exactly one query action:

```text
BuildCfideerQueries = true
CopyExistingCFIDEERQueriesFromExistingDatabase = false
```

or:

```text
BuildCfideerQueries = false
CopyExistingCFIDEERQueriesFromExistingDatabase = true
```

Both true and both false are invalid for Upgrade mode.

### Query Rebuild

`Build-CfideerQueries` creates or replaces the nine working `_prjCFIDEER_` query definitions.

Field inclusion depends on:

- Physical table fields
- Current AppColumns state for editable standard measurement fields and custom fields
- Both `Active = true` and `QueryVisible = true` for user-controlled query fields
- Get Field Order for the three Get-query display lists
- Required system fields
- Query-specific exclusions
- Required starting and trailing fields

`ReportVisible` is not part of CFIDEER field eligibility.

`Test-CfideerRequiredSystemField` prevents AppColumns metadata from hiding required relationship fields, keys, criteria fields, Remarks fields, or audit fields. Some protected fields are criteria-only or query-managed and therefore do not appear in Get Field Order.

### Version 2.2.1 Standard-Field Visibility Filter

Version 2.2.1 closes the UI fallback paths that previously allowed an editable standard field to reappear after the user turned off `Active` or `QueryVisible`.

Important UI helpers are:

```text
Get-GetOrderSetupMaps
Test-GetOrderKeyAllowedBySetup
Add-GetOrderCandidatesFromSavedQuery
Add-GetOrderFieldSpecs
Add-GetOrderCandidateFromAppColumnItem
Test-GetOrderDefaultFieldAllowed
Add-GetOrderDefaultFields
```

`Get-GetOrderSetupMaps` first calls `Save-LoadedAppColumnGridState`, then builds case-insensitive `TableName|FieldName` sets for the current and original AppColumns states. The current set contains only fields that are both Active and QueryVisible. The original set allows Project Pickle to preserve the prior saved-query order for fields that were already eligible while adding newly activated fields later at the bottom.

`Add-GetOrderCandidatesFromSavedQuery` and `Add-GetOrderFieldSpecs` call `Test-GetOrderKeyAllowedBySetup`. A field that is currently inactive or QueryVisible-off is therefore removed whether the candidate came from saved Get SQL or from physical table order.

`Test-GetOrderDefaultFieldAllowed` prevents the built-in Plot, Tree, and Regen default lists from adding the field back. A default that has an editable AppColumns row must obey that row's current Active + QueryVisible state. A structural default with no editable AppColumns row remains available.

`Add-GetOrderCandidateFromAppColumnItem` adds only current Active + QueryVisible rows. This is the final path that appends new or newly activated fields after preserved order.

The generated-query layer uses:

```text
New-CfideerQueryVisibleLookup
Filter-CfideerFieldsByQueryVisible
Test-CfideerRequiredSystemField
```

`New-CfideerQueryVisibleLookup` creates a case-insensitive table/field lookup from the complete working AppColumns configuration. `Filter-CfideerFieldsByQueryVisible` applies it to standard and custom Get/Update field lists. Explicit custom fields added to Process queries use the same eligible custom-field lists; fixed or query-specific Process fields retain their defined structure.

Behavioral rules:

- `Active = false` excludes a user-controlled field even if QueryVisible is true.
- `QueryVisible = false` excludes a user-controlled field even if Active is true.
- `ReportVisible` does not affect Get Field Order or CFIDEER query inclusion.
- Required identifiers, relationships, criteria, keys, Remarks fields, and audit fields remain protected where the query requires them.
- Exact old-query copy mode bypasses regeneration and continues to preserve source SQL exactly.

### Query Copy

`Copy-CFIDEERQueriesFromExistingDatabase` copies the Existing/old saved query names and SQL exactly.

Copy mode validates all nine required query definitions before the long inventory import.

After DAO saves the definitions, Project Pickle reopens them and compares normalized SQL. Formatting whitespace and a trailing semicolon can be normalized, but parameters, fields, joins, SET assignments, and criteria must remain equivalent.

### Copied-Query Get Field Preview

Selecting exact query copy immediately parses and displays the three Existing/old Get-query SELECT lists in Get Field Order.

The preview uses:

```text
_prjCFIDEER_GetPlotMeasurementsForPeriod
_prjCFIDEER_GetTreeMeasurementsForPeriodByKey
_prjCFIDEER_GetRegenMeasurements
```

The grids remain enabled so the user can scroll and inspect them.

These controls are disabled:

```text
Load fields from AppColumns setup
Clear custom order
Move up
Move down
Build/refresh _prjCFIDEER_ get and update queries
```

The preview does not represent the Update or Process query field lists. Those definitions are still copied and verified by the Build path.

The manual Get Field Order is snapshotted before preview mode and restored when query copy is unchecked.

A source-path change invalidates and clears copied-query mode.

Copy mode and rebuild mode remain mutually exclusive. Upgrade Project still requires one of them.

### Placeholder Cleanup

`FieldNameSampleRemove` is a template-only placeholder.

Physical custom-table placeholder fields and corresponding AppColumns metadata are removed.

SELECT-query placeholder references can be removed from the SELECT list.

Inherited placeholder UPDATE definitions can be removed and later recreated by rebuild mode or copied from the clean old database.

### Saved-Query Verification

Important functions include:

```text
Assert-CfideerSavedQueries
Assert-CopiedCfideerQueriesCompatible
Test-CfideerSqlContainsQualifiedFieldReference
Remove-FieldNameSampleRemoveFromCfideerQueries
Remove-AccessQueryDefinition
```

Verification checks:

- All nine required saved definitions
- Nonblank SQL
- Literal `FieldNameSampleRemove`
- Access-generated `Exp#` and `Expr#` names
- Required relationship and criteria fields
- Expected custom fields
- First fields immediately following `SET`
- Final fields immediately preceding `FROM`
- Table-qualified and explicitly aliased references

The matcher must preserve word-separating whitespace. Removing all whitespace can create false negatives such as:

```text
SETPlotMeasurements.PlotMeasKey
```

or:

```text
PlotCustomMeasurements.FieldUnderSpp2FROM
```

A verification failure is fatal. A database whose query verification fails must not be marked successful or moved directly into processing.

## 13. Assignments

Plot assignment setup is a rebuild. Existing rows in `InventoryAssignmentTrees`, `InventoryAssignmentPlots`, `InventoryAssignments`, and `InventoryCrews` are cleared in the output before a clean assignment set is created.

Assignments are built for the current period. When no period is marked current, the implementation can fall back to the highest period and logs that choice.

Plot assignments require plots. Tree assignments are Upgrade-only, require tree rows, and depend on plot assignments.

The build verifies final assignment row counts and required Awaiting Work status values. Keep those checks whenever assignment logic changes.

## 14. Processing, Calculated Tables, and Project Identifiers

The Processing tab is available only in Upgrade Project mode and contains two independent, default-off options.

### Processing Setup Copy Plan

The normal four-table set is:

- `SpeciesMerchantabilityStandards`
- `SpeciesNationalVolumeParms`
- `SpeciesRegressionCriteria`
- `SpeciesRegressions`

The merchantability compatibility fields are:

- `HasDRC`
- `DMsusceptible`
- `DMcollected`

Core functions include:

```text
Get-ProcessingTableNames
Get-ProjectPickleProcessingTableColumnNames
Get-ProjectPickleProcessingCopyCompatibility
Get-ProjectPickleProcessingRequiredMerchantabilityFieldNames
Copy-ProcessingTablesFromExistingDatabase
Set-SpeciesMerchantabilityProjectId
```

### Compatible Four-Table Path

When the source merchantability table exists and contains all required fields, the copy plan includes all four tables.

The copy is a complete transactional replacement. Project Pickle verifies source and target table presence, clears target rows in safe order, copies common columns, commits, and logs source/output row counts.

### Three-Table Fallback

The fallback is selected when:

- The old database does not contain `SpeciesMerchantabilityStandards`, or
- The table is missing one or more required merchantability fields

The fallback copies:

- `SpeciesNationalVolumeParms`
- `SpeciesRegressionCriteria`
- `SpeciesRegressions`

The current-master `SpeciesMerchantabilityStandards` table remains in the output.

The three fallback tables remain mandatory. If one is missing, the Processing setup import stops.

### Species Merchantability Identifier Normalization

`Set-SpeciesMerchantabilityProjectId` supports:

- `ProjectID`
- `BIAProjectID`
- Both fields
- GUID storage
- Legacy text GUID storage

The normalization runs:

1. After project setup
2. After Processing-table handling
3. During final verification

Every available field is updated to the current project GUID and read back. Any blank, invalid, or mismatched value is fatal.

This covers:

- A compatible four-table Processing copy
- A three-table fallback retaining the current-master merchantability table
- A build that leaves Processing copy off

### Calculated And Summary Tables

The second Processing option replaces:

- `PlotSummaries`
- `TreeCalculations`
- `RegenCalculations`

This option is independent from the Processing setup copy.

An empty source table produces an empty target table. Source and target row counts are logged and verified.

## 15. Logging, Diagnostics, Application Title, and Cancellation

The production UI writes detailed text logs rather than relying only on a run-log table inside the project database.

Final text logs are normally written beside the output database. Live logs use a temporary `Project Pickle Logs` directory.

Crash and diagnostic files are written near the application folder when possible:

- `ProjectPickle.LastError.txt`
- `ProjectPickle.Diagnostics.txt`

Detailed run-log entries include:

- Step number
- Log time
- Category
- Elapsed time
- Step duration
- Mode
- Project
- Input paths
- Output path
- Message

Warnings and failures must remain distinguishable.

Approved repeated Regen keys and matched-ID Regen key differences can be warnings.

These remain fatal:

- Blank or duplicate IDs
- Count differences
- Missing or orphan custom rows
- Invalid parent links
- Stale Upgrade Mapping
- Description write failure
- Species Merchantability ID failure
- CFIDEER query verification failure
- Application Title persistence failure

### Failed Outputs

The build is not one database-wide transaction.

A Failed output can contain:

- A copied master database
- Project and period rows
- Imported inventory rows
- Copied metadata
- Partial reports or queries

A Failed output is not a usable upgraded database. Operator documentation must instruct users to delete or quarantine it and rebuild from clean inputs.

### Access Application Title

`New-AccessApplicationTitle` calculates the title from:

- Existing/template `AppTitle`
- Project name
- Build date placeholders

`Set-AccessApplicationTitle` writes the DAO `AppTitle` property.

For New and Upgrade modes:

1. Write the title after the output copy is created
2. Close and reopen the database
3. Verify the persisted title
4. Verify it again before successful completion

Existing Project mode preserves the existing title.

### Duration Formatting

`Format-WOOFDuration` must use `Math.Floor` for whole hours. Casting a floating-point `TotalHours` value directly to `[int]` can round the displayed hour upward.

### Cancellation

The UI processes cancellation at safe checkpoints. An Access operation already executing cannot always be interrupted safely. Do not replace safe-checkpoint cancellation with forced thread termination.

Do not commit generated logs. Logs can reveal internal paths, project names, identifiers, warnings, and row samples.

## 16. AI Help Architecture

AI Help is optional and database read/write behavior is not part of the AI request path.

Core functions are:

- `Get-WOOFAiContext`
- `Invoke-WOOFAiHelpRequest`
- `Invoke-WOOFAiChat`
- `Get-WOOFGuideTextFromFile`
- `Get-RelevantTextSnippets`
- `Assert-ApprovedAiEndpoint`

The recommended bundled context is:

```text
Project Pickle App Files/AI-HELP-CONTEXT.md
```

`Get-WOOFAiContext` should load that focused file automatically. The provided patch keeps a fallback to a legacy README when the new context file is missing.

A request can contain:

- User question
- Bundled AI Help context
- Relevant `ProjectPickle.ps1` snippets when Include code is checked
- Relevant snippets from an optional selected guide/reference file

Supported optional guide types are DOCX, XLSX, XLSM, TXT, MD, CSV, JSON, and PS1. PDF and legacy XLS are intentionally rejected by the in-app reader.

The current AI configuration in version 2.2.1 uses:

- Azure OpenAI API version `2024-12-01-preview`
- Default deployment/model `gpt-4.1-mini`
- An organization-specific default endpoint and approved host list
- HTTPS enforcement and exact approved-host checks
- Optional additional hosts from `CFI_PROJECT_PICKLE_APPROVED_AI_HOSTS`, with two legacy environment-variable fallbacks

The API key is sent in the Azure `api-key` header and is not intended to be persisted. Review endpoint names, model deployment names, and organization-specific infrastructure details before public source publication.

Keep `AI-HELP-CONTEXT.md` focused and below the configured text limit. Update it whenever mode rules, tab names, copy semantics, or major troubleshooting steps change.

## 17. Testing

Minimum pre-release testing begins with a parser check and the built-in self-test.

```powershell
$scriptPath =
    (Resolve-Path ".\ProjectPickle.ps1").Path

$tokens =
    $null

$parseErrors =
    $null

$ast =
    [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath,
        [ref]$tokens,
        [ref]$parseErrors
    )

if ($parseErrors.Count -gt 0) {
    $parseErrors |
        Format-List Message, Extent

    throw "PowerShell parser found errors."
}

$duplicateFunctions =
    @(
        $ast.FindAll(
            {
                param($Node)

                $Node -is
                [System.Management.Automation.Language.FunctionDefinitionAst]
            },
            $true
        ) |
        Group-Object Name |
        Where-Object {
            $_.Count -gt 1
        }
    )

if ($duplicateFunctions.Count -gt 0) {
    $duplicateFunctions |
        Format-Table Name, Count

    throw "Duplicate function names were found."
}
```

Run:

```powershell
.\Run-ProjectPickle-32bit.bat -SelfTest
```

Also test:

### Startup And Packaging

- Root VBS launcher
- Root 32-bit batch launcher
- UI smoke test
- Clean clone or clean ZIP extraction
- Missing icon and missing AI-context behavior
- Correct launcher-relative paths

### Source Inventory Preflight

- Tree measurement/custom row-count mismatch
- Plot measurement/custom row-count mismatch
- Regen measurement/custom row-count mismatch
- Equal counts with one missing and one unrelated orphan row
- Blank MeasurementID
- Duplicate MeasurementID
- Missing Trees.PlotID parent
- Missing PlotMeasurements.PlotID parent
- Missing TreeMeasurements.TreeID parent
- Missing RegenMeasurements.PlotID parent
- Plot matched-ID key mismatch
- Tree matched-ID key mismatch
- Valid repeated Regen keys
- Declined repeated-Regen approval
- Stale repeated-Regen approval signature
- Matched-ID Regen key difference warning
- Source change after warning approval

### Upgrade Mapping

- No source-only fields
- Grid never loaded
- Grid loaded for another source/master pair
- Missing required row
- Extra stale row
- Duplicate row
- Invalid row
- Intentional Skip confirmation
- Declined Skip confirmation
- Saved JSON for the same source/master pair
- Saved JSON for a different source/master pair
- Table-only history-field target
- Specific renamed field target
- Datatype incompatibility

### AppColumns Copy

- All AppColumns rows copied
- Linked AppColumnCodes copied
- Orphan AppColumnCodes omitted from working metadata
- Complete historical code archive retained
- Field already present in current master
- Active physical field present only in old source
- Inactive physical history field present only in old source
- Inactive catalog-only field absent from both schemas
- Active metadata field absent from both schemas
- Duplicate copied AppColumns table/field metadata
- `BearDamage`-style inactive metadata does not create a field
- Exact source custom-table Design View order
- Current-master-only fields appended after source order

### Access Description Preservation

- Table Description copied
- Field Description copied
- Blank source Description does not erase target Description
- Description follows same-name field
- Description follows Upgrade Mapping target
- Description follows renamed-field target
- Description follows paired measurement/custom-table move
- No-target Description is logged
- Conflicting source descriptions mapped to one target are blocked

### Inventory Import And Performance

- Per-table source/output count verification
- Source/output identity preservation
- Source/output parent-reference preservation
- Normal build skips legacy repair passes
- `-SelfTest` still exercises repair routines
- Large Tree measurement project
- Local output timing
- OneDrive/synchronized output timing comparison
- Correct duration display above one hour

### Processing

- Compatible four-table copy
- Missing merchantability table fallback
- Missing `HasDRC` fallback
- Missing `DMsusceptible` fallback
- Missing `DMcollected` fallback
- Missing each mandatory three-table fallback source
- Calculation/summary copy by itself
- Both Processing options together
- Empty calculation table
- `ProjectID` only
- `BIAProjectID` only
- Both project-ID fields
- GUID fields
- Text GUID fields
- Full Processing copy normalization
- Fallback retained-master normalization
- Final verification

### CFIDEER Queries

- Normal query rebuild
- Exact old-query copy
- Both query actions selected
- Both query actions off in Upgrade mode
- Missing one of the nine required source queries
- Blank source query SQL
- `FieldNameSampleRemove` in a SELECT query
- `FieldNameSampleRemove` in an UPDATE query
- `Exp1`
- `Expr1`
- Required first field after `SET`
- Required final field before `FROM`
- Bracketed qualified field
- Explicit table alias
- Copied SQL equivalence after DAO save
- Missing expected active custom field
- Inactive standard Plot field is absent from Get Field Order and rebuilt Get/Update SQL
- QueryVisible-off standard Plot field is absent from Get Field Order and rebuilt Get/Update SQL
- Equivalent standard-field filtering for Tree and Regen
- Saved Get-query order cannot re-add an ineligible field
- Physical-table fallback order cannot re-add an ineligible field
- Built-in default order cannot re-add an ineligible field
- Newly activated eligible fields append after preserved order
- Required system fields remain present even when they are not editable AppColumns choices
- Exact query-copy mode retains source SQL without applying the rebuild filter
- Plot Update query field order
- Tree Update query field order
- Regen Update query includes PlotKey and required relationship fields
- GetProcess Plot, Tree, and Regen custom fields

### Reports

- Manual Plot, Tree, and Regen layouts
- Completed report copy
- AppColumns ID compatibility
- `RemarksTree = TEXT,40,5`
- Final-custom-field separator
- Specialized text and date formats
- Header line-length validation
- Header whitespace preservation

### Application Title

- New Project title
- Upgrade Project title
- Existing Project title preservation
- Missing `AppTitle` property creation
- Existing `AppTitle` update
- Placeholder replacement
- Close/reopen persistence verification
- Final verification
- Access reopen/caption refresh

### Final Build

- One representative New Project Success
- One representative Upgrade Project Success
- One representative Existing Project Success
- Failed output is not opened automatically
- Success log includes all required final verification messages

Use synthetic or approved test databases. Never place production databases in the repository or CI artifacts.

### Read-Only Preview Tests

#### AppColumns

- Select Upgrade Project and valid old/master paths.
- Load a manual or standard AppColumns list.
- Check Copy AppColumns.
- Confirm Existing rows load and are labeled WILL COPY.
- Confirm current-master rows load as reference-only.
- Confirm the Preview dropdown shows Existing, master, and both.
- Confirm filters and search remain usable.
- Confirm rows cannot be edited.
- Confirm current-master display rows do not enter query rebuilding.
- Uncheck Copy AppColumns and confirm the prior manual list returns.
- Change the source path and confirm preview mode turns off.

#### CFIDEER

- Load a manual Get Field Order.
- Check copied-query mode.
- Confirm Plot, Tree, and Regen old Get-query fields load.
- Confirm grids remain scrollable and read-only.
- Confirm Load, Clear, Move up, and Move down are disabled.
- Confirm query rebuilding becomes unchecked and disabled.
- Uncheck query copy and confirm the prior manual order returns.
- Confirm query rebuilding becomes checked and enabled.
- Change the source path and confirm copied-query mode turns off.

#### Reports

- Load or edit a manual report layout.
- Check completed-report copy.
- Confirm old Plot, Tree, and Regen layouts load.
- Confirm ReportColumns and exact header text are visible.
- Confirm editing controls are disabled.
- Confirm Preview / print fake data remains enabled.
- Confirm Update reports is checked and disabled.
- Switch among all three layouts.
- Uncheck completed-report copy and confirm the prior manual layouts return.
- Change the source path and confirm completed-report copy turns off.

### Launcher And Splash Tests

Run the PowerShell parser against:

```text
ProjectPickle.ps1
ProjectPickleSplash.ps1
```

Then test:

```powershell
cscript.exe //nologo ".\Project Pickle App Files\ProjectPickle.vbs"
cscript.exe //nologo ".\Run Project Pickle.vbs"
```

Verify:

- Internal launcher starts the app.
- Root launcher finds the internal launcher.
- Root launcher does not search for ProjectPickle.ps1 in the root folder.
- Splash shows the built-in pickle without a PNG.
- A valid PNG overrides the built-in artwork.
- An invalid optional PNG falls back safely.
- Pickle rotates through prepared frames.
- Loading dots animate.
- Splash closes when the main app deletes the marker.
- Splash safety timeout closes a stuck splash.
- Main app still starts when the splash script is absent and the internal launcher takes its no-splash fallback path.
- Complete folder still works after being renamed and moved.
- VBS files compile under Windows Script Host using Windows-1252 encoding.

For every release:

1. Complete a parser check with zero errors.
2. Confirm no duplicate function names.
3. Run `-SelfTest`.
4. Complete New, Upgrade, and Existing Project test builds.
5. Complete at least one full-sized Upgrade test.
6. Verify source preflight warning and blocking paths.
7. Verify Upgrade Mapping review behavior.
8. Verify copied AppColumns physical-field classification.
9. Verify custom-table Design View order.
10. Verify Access table and field Descriptions.
11. Verify Species Merchantability `ProjectID`/`BIAProjectID`.
12. Verify copied and rebuilt CFIDEER query paths.
13. Verify standard and custom Active + QueryVisible filtering in Get Field Order and rebuilt Get/Update queries.
14. Verify required CFIDEER system fields remain present.
15. Verify Application Title persistence.
16. Update `$script:AppVersion`.
17. Update README version, features, and screenshot caption.
18. Update the HTML guide reviewed date and version.
19. Update these Technical Notes.
20. Update `AI-HELP-CONTEXT.md`.
21. Update `CHANGELOG.md`.
22. Search for stale version numbers and launcher names.
23. Run secret and sensitive-data checks.
24. Test from a clean clone or extracted release package.
25. Tag the tested commit, for example `v2.2.1`.
26. Build the release ZIP from committed files, not a working folder containing logs.

Useful checks include:

```powershell
git grep -n -E "2\.1\.7|2\.1\.8|v2\.1\.8|Project Pickle\.vbs"
git status --short
git diff --check
```

Review results rather than applying blind replacements. Older version numbers can be valid in historical changelog content.

Do not advance the documented release version until a full Upgrade test reaches:

```text
Inventory integrity (final build) passed
CFIDEER saved-query verification passed
Access application title verified
Database build complete
```

## 19. Known Constraints and Refactoring Priorities

Current constraints include:

- Windows-only runtime
- Access provider and process-bitness dependency
- A large single PowerShell source file
- Full build is not one database-wide transaction
- Large inventory import is still primarily row-oriented
- Large Tree measurement projects can take several hours
- Cloud-synchronized output folders can increase Access write time
- Report layouts are held in memory until Build
- Cancellation occurs at safe checkpoints rather than immediately
- Optional guide reader does not support PDF or legacy XLS
- AI Help accuracy depends on synchronized local documentation
- Access may require closing and reopening a database before the updated title is visible in the window caption
- Normal startup depends on Windows Script Host and Windows PowerShell.
- VBS source encoding can cause Windows Script Host compilation errors when hidden or unsupported characters are introduced.
- The splash and main application run in separate PowerShell processes.
- The ready-marker protocol depends on a writable temporary folder or application-folder fallback.
- Custom splash artwork should be PNG; the main ICO should not be enlarged as splash artwork.
- Read-only previews duplicate UI state in memory until the copy option is cleared.

Source preflight and production repair-pass removal reduce wasted work but do not eliminate the main row-by-row import cost.

The highest-value performance improvement is transaction or batch optimization for large inventory imports. Do not remove integrity verification merely to improve speed.

A low-risk refactoring path is to keep `ProjectPickle.ps1` as a small entry point and move stable groups into tested modules:

```text
Modules/
|-- ProjectPickle.Data.psm1
|-- ProjectPickle.Upgrade.psm1
|-- ProjectPickle.Reports.psm1
|-- ProjectPickle.Cfideer.psm1
|-- ProjectPickle.AiHelp.psm1
|-- ProjectPickle.Logging.psm1
`-- ProjectPickle.UI.psm1
```

Refactor only after a tagged known-good release and use small commits with behavioral comparison tests.

Future priorities:

1. Batch or transaction-optimize high-volume inventory inserts.
2. Add a structured preflight report that can be exported before Build.
3. Add a structured warning summary to successful run logs.
4. Add automated schema-property comparison for field order and Description values.
5. Add automated CFIDEER SQL snapshots for all nine required queries.
6. Reduce duplication between UI preflight and backend validation.
7. Move stable functions into modules without changing behavior.

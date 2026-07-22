# Project Pickle Technical Notes

Application version: 2.1.4

These notes are for maintainers, reviewers, and developers. The operator-facing workflow belongs in `Project Pickle How-To Guide.html`; the repository overview belongs in the root `README.md`; and the focused in-app AI reference belongs in `Project Pickle App Files/AI-HELP-CONTEXT.md`.

## 1. Source of Truth and Documentation Roles

Use the following order when behavior and documentation differ:

1. The current released `ProjectPickle.ps1` implementation and a reproducible test result.
2. The current operator how-to guide.
3. These technical notes.
4. The focused AI Help context.
5. Historical README files or release notes.

Update all version references and workflow names in the same release. Version 2.1.4 includes the Reports and Processing tabs, so documentation that stops at the older tab sequence is not current.

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
| `-AiEndpoint`, `-AiModel`, `-AiApiKey` | Supply AI Help values at startup. |

Top-level failures are written to `ProjectPickle.LastError.txt` and `ProjectPickle.Diagnostics.txt` near the root application folder when possible.

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

## 6. High-Level Build Pipeline

`Invoke-InventoryBuild` performs the main build in this approximate order:

1. Prepare the Processing copy plan when the first Processing option is selected.
2. Validate configuration, mode rules, required paths, selected copy options, Processing compatibility, report layouts, and required query definitions.
3. Move an existing output file to a timestamped backup.
4. Copy the mode-appropriate base database to the output path.
5. Open the output with the first working OLE DB provider.
6. Create or refresh `AppColumnsArchive` from the output copy's original AppColumns table.
7. In Upgrade Project mode, optionally archive the old database's AppColumns setup to `zAppColumns` and `zAppColumnsCodes`.
8. Resolve Region, Agency, and Reservation and write project setup.
9. Write periods for New or Upgrade mode, or preserve existing periods in Existing mode.
10. Update Analysis table period IDs and optionally prune reference rows.
11. Completely copy working AppColumns tables or synchronize selected AppColumns, code lists, and custom fields.
12. Import old inventory data and apply Upgrade Mapping in Upgrade mode when enabled.
13. Copy either four Processing setup tables or the three-table compatibility fallback when the first Processing option is selected.
14. Copy `PlotSummaries`, `TreeCalculations`, and `RegenCalculations` when the second Processing option is selected.
15. Reorder physical custom measurement table columns.
16. Update validation periods when selected.
17. Import new plots and create missing blank plot/custom measurement rows.
18. Rebuild plot and optional tree assignments when selected and plots exist.
19. Propagate ProjectID values where required.
20. Copy completed reports, apply saved manual layouts, or run the automatic report updater.
21. During the manual report path, preserve `RemarksTree` as `TEXT,40,5` and apply the final-custom-field separator before Remarks.
22. Either copy saved `_prjCFIDEER_` query definitions from the old database or rebuild CFIDEER queries. These paths are mutually exclusive.
23. Preserve the existing Access application title in Existing mode or update it in New and Upgrade mode.
24. Finalize logs and optionally open the saved database.

Maintain this order carefully. Several later operations depend on IDs, periods, custom fields, inventory rows, archived setup, or copied setup objects created earlier in the build.

The old AppColumns archive executes after `AppColumnsArchive` is created and before working AppColumns synchronization. This preserves both the base database's original setup and the old project database's setup even when the working AppColumns are later replaced or manually changed.

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

In Upgrade Project mode, this normally represents the current master database version's AppColumns setup because the current master is copied to the output first.

### Old Project AppColumns Archives

`Copy-ProjectPickleOldAppColumnsArchives` is a separate, Upgrade-only archive path controlled by:

```text
ArchiveOldAppColumnsFromExistingDatabase

## 9. Upgrade Import and Field Mapping

Upgrade behavior is concentrated in functions whose names begin with `Upgrade`, including mapping review, datatype checks, row import, moved-field bulk operations, repair, and diagnostics.

The upgrade pipeline attempts fast bulk Access operations for compatible moved fields and falls back to row-by-row handling when the old database shape or field type prevents a bulk operation.

Datatype review should stop unsafe conversions before import. Old text should not be silently written to a current integer, date, Boolean, or GUID target simply because a few sample values appear convertible.

Repair routines can recreate missing parent plot or tree rows when enough identifiers are available and can create missing custom measurement rows. These repairs are logged with identifiers and counts for later review.

Post-import diagnostics compare source and output counts, counts by period, parent-child relationships, and matching custom measurement rows. Treat those diagnostics as release-critical evidence, not informational noise.

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

### RemarksTree FieldFormat Protection

`TreeMeasurements.RemarksTree` is required to remain:

```text
TEXT,40,5

## 12. CFIDEER Query Generation and Query Copy

`Build-CfideerQueries` creates or replaces the working `_prjCFIDEER_` query definitions in the output database.

Query field inclusion depends on active and QueryVisible AppColumns. Get Field Order preserves saved order where possible and allows new active/query-visible fields to be positioned before query rebuilding.

### Normal Query-Rebuild Path

The normal path uses:
BuildCfideerQueries = true
CopyExistingCFIDEERQueriesFromExistingDatabase = false

## 13. Assignments

Plot assignment setup is a rebuild. Existing rows in `InventoryAssignmentTrees`, `InventoryAssignmentPlots`, `InventoryAssignments`, and `InventoryCrews` are cleared in the output before a clean assignment set is created.

Assignments are built for the current period. When no period is marked current, the implementation can fall back to the highest period and logs that choice.

Plot assignments require plots. Tree assignments are Upgrade-only, require tree rows, and depend on plot assignments.

The build verifies final assignment row counts and required Awaiting Work status values. Keep those checks whenever assignment logic changes.

## 14. Processing and Calculated Table Copy

The Processing tab is available only in Upgrade Project mode and contains two independent, default-off options.

### Processing Setup Copy Plan

The first option imports project-specific Processing setup from the old database.

The normal four-table set is:

- `SpeciesMerchantabilityStandards`
- `SpeciesNationalVolumeParms`
- `SpeciesRegressionCriteria`
- `SpeciesRegressions`

The three merchantability field names inspected by the compatibility planner are:

- `HasDRC`
- `DMsusceptible`
- `DMcollected`

Core implementation elements include:

- `Get-ProcessingTableNames`
- `Get-ProjectPickleProcessingTableColumnNames`
- `Get-ProjectPickleProcessingCopyCompatibility`
- `Get-ProjectPickleProcessingRequiredMerchantabilityFieldNames`
- `Copy-ProcessingTablesFromExistingDatabase`
- `$script:ProjectPickleProcessingTableNamesOverride`
- `$script:ProjectPickleSkipSpeciesMerchantabilityStandards`

### Compatible Four-Table Path

When `SpeciesMerchantabilityStandards` exists and contains all three required field names, the copy plan includes all four tables.

`Copy-ProcessingTablesFromExistingDatabase` validates source and target table presence, validates required merchantability fields, begins a target transaction, deletes target rows in safe order, copies common columns, commits, and logs source and output row counts.

This is a complete replacement rather than a merge.

### Three-Table Compatibility Fallback

The fallback is selected when either:

- The old database does not contain `SpeciesMerchantabilityStandards`, or
- The table is missing one or more of `HasDRC`, `DMsusceptible`, and `DMcollected`.

The fallback table plan is:

- `SpeciesNationalVolumeParms`
- `SpeciesRegressionCriteria`
- `SpeciesRegressions`

The compatibility planner displays a warning and explains that:

- The three compatible tables will be completely replaced from the old database.
- `SpeciesMerchantabilityStandards` will not be copied.
- The current-master version of `SpeciesMerchantabilityStandards` will remain in the output.
- Old project-specific merchantability standards may need to be recreated manually.

The existing four-table validator is temporarily suppressed only for this planned fallback. The user's Processing checkbox remains selected, and the three-table name override remains active for the actual copy.

The fallback-aware required-field helper returns no merchantability fields while `SpeciesMerchantabilityStandards` is intentionally excluded from the copy plan.

### Non-Skippable Processing Tables

The following source tables remain mandatory:

- `SpeciesNationalVolumeParms`
- `SpeciesRegressionCriteria`
- `SpeciesRegressions`

If any of these is missing, `Get-ProjectPickleProcessingCopyCompatibility` throws before the copy plan can proceed. No Processing setup table replacement should begin.

Target-table absence or a later copy failure causes the target transaction to roll back.

### Calculated And Summary Tables

The second Processing option independently replaces:

- `PlotSummaries`
- `TreeCalculations`
- `RegenCalculations`

`Copy-ProcessingCalculatedTablesFromExistingDatabase` validates that all three tables exist in the source and target, deletes target rows in reverse table order, copies all compatible columns, commits the transaction, and logs source and output row counts.

The second option does not require the first option. Either Processing option can run by itself or both can run in the same upgrade.

Because this is a complete replacement, an empty source table produces an empty target table. Use this option only when the source calculation and summary rows belong to the inventory data being imported into the upgraded database.

## 15. Logging, Diagnostics, and Cancellation

The production UI writes detailed text logs rather than relying on a run-log table inside the project database.

Final text logs are normally written beside the output database. Live logs use a temporary `Project Pickle Logs` directory. Crash and diagnostic files are written near the root application folder when possible:

- `ProjectPickle.LastError.txt`
- `ProjectPickle.Diagnostics.txt`

Detailed run-log entries include step number, time, category, elapsed time, step duration, mode, project, input paths, output path, and message.

The UI processes cancellation at safe checkpoints. An Access operation already executing cannot always be interrupted safely. Do not replace safe-checkpoint cancellation with forced thread termination.

Do not commit any generated logs. Logs can reveal internal paths, project names, identifiers, and row samples.

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

The current AI configuration in version 2.1.4 uses:

- Azure OpenAI API version `2024-12-01-preview`
- Default deployment/model `gpt-4.1-mini`
- An organization-specific default endpoint and approved host list
- HTTPS enforcement and exact approved-host checks
- Optional additional hosts from `CFI_PROJECT_PICKLE_APPROVED_AI_HOSTS`, with two legacy environment-variable fallbacks

The API key is sent in the Azure `api-key` header and is not intended to be persisted. Review endpoint names, model deployment names, and organization-specific infrastructure details before public source publication.

Keep `AI-HELP-CONTEXT.md` focused and below the configured text limit. Update it whenever mode rules, tab names, copy semantics, or major troubleshooting steps change.

## 17. Testing

Minimum pre-release testing should include:

```powershell
.\Run-ProjectPickle-32bit.bat -SelfTest
```

Also run:

- UI startup through `Run Project Pickle.vbs`
- Startup and self-test through the 32-bit batch launcher
- `-UiSmokeTest` through the internal script or a launcher that forwards arguments
- Reference-list test where the required test database is available
- One representative New Project build
- One representative Upgrade Project build with datatype review and diagnostics
- One representative Existing Project build
- Report layout validation and build
- CFIDEER query comparison against expected SQL and consuming-app behavior
- Assignment count and status verification
- AI Help questions that exercise mode choice, AppColumns, Reports, and missing-context fallback
- Clean clone or clean ZIP extraction on a second test folder
- Upgrade-mode UI check confirming Processing is enabled only in Upgrade Project mode.
- First Processing option with a compatible source containing all four tables and all three required merchantability fields, confirming that all four tables are copied.
- First Processing option with `HasDRC` missing, confirming the warning appears, the three compatible tables are copied, and target `SpeciesMerchantabilityStandards` remains inherited from the current master.
- Repeat the fallback test separately with `DMsusceptible` missing.
- Repeat the fallback test separately with `DMcollected` missing.
- First Processing option with the entire source `SpeciesMerchantabilityStandards` table missing, confirming the same three-table fallback.
- First Processing option with `SpeciesNationalVolumeParms` missing, confirming the setup import stops and no Processing setup replacement begins.
- Repeat the mandatory-table failure test separately for `SpeciesRegressionCriteria`.
- Repeat the mandatory-table failure test separately for `SpeciesRegressions`.
- Second Processing option selected by itself.
- Both Processing options selected together with a compatible four-table source.
- Both Processing options selected together while the first option uses the three-table fallback.
- Calculated/summary source tables containing rows and containing zero rows.
- New Project and Existing Project UI checks confirming that the old AppColumns archive checkbox is hidden.
- Upgrade Project UI check confirming that the old AppColumns archive checkbox is visible, enabled, and checked by default.
- Upgrade build with old AppColumns archive checked and working AppColumns copy unchecked.
- Upgrade build with both old AppColumns archive and working AppColumns copy checked.
- Upgrade build with old AppColumns archive unchecked.
- Existing output containing prior `zAppColumns` and `zAppColumnsCodes`, confirming both are refreshed when the option is checked.
- Archive source with different AppColumns/AppColumnCodes row counts, confirming exact source/output count verification.
- Archive source missing AppColumns, confirming Build stops before leaving an incomplete archive.
- Archive source missing AppColumnCodes, confirming Build stops before leaving an incomplete archive.
- Forced failure after creation of one z table, confirming both z tables are removed.
- Reports test confirming Add ReportVisible fields leaves `RemarksTree` as `TEXT,40,5`.
- Plot report test confirming `{0,1}` on the final custom field becomes `{0,2:0\ }` for storage and printing.
- Tree report test confirming `{0,4:g0}` on the final custom field becomes `{0,5:0\ }`.
- Regen report test confirming the same automatic separator rule.
- Report reload/save test confirming the stored width is not increased a second time.
- Report test confirming specialized text and date formats are not converted.
- ReportHeaders validation after the extra separator character changes the printed width.
- Upgrade CFIDEER query-copy path with valid `_prjCFIDEER_` source queries.
- Upgrade CFIDEER query-copy path with no matching source queries.
- New Project and Existing Project checks confirming query copy is disabled.
- Query-copy toggle check confirming Get Field Order and normal rebuild controls disable and re-enable correctly.
- Normal CFIDEER query rebuild after loading Get Field Order.
- Verification that copied query definitions replace target `_prjCFIDEER_` definitions and that normal query rebuilding is skipped.

Use synthetic or approved test databases. Never place production databases in the repository or CI artifacts.


## 18. Release Maintenance

For every release:

1. Update `$script:AppVersion`.
2. Update the root README version and feature list.
3. Update the HTML guide reviewed version and tab order.
4. Update `AI-HELP-CONTEXT.md`.
5. Update these technical notes when architecture or dependencies change.
6. Update `CHANGELOG.md`.
7. Capture a sanitized screenshot when the UI changes materially.
8. Search for stale version numbers and launcher names.
9. Run secret and sensitive-data checks.
10. Test from a clean clone or extracted release package.
11. Tag the tested commit, for example `v2.1.4`.
12. Build the release ZIP from committed files, not from a working folder containing logs.

Useful checks include:

```powershell
git grep -n -E "2\.0\.08|Project Pickle\.vbs"
git status --short
git diff --check
```

Review the results rather than applying blind replacements. A string can be valid historical changelog content.

## 19. Known Constraints and Refactoring Priorities

Current constraints include:

- Windows-only runtime
- Access provider and process-bitness dependency
- A large single PowerShell source file
- Report layouts held in memory until Build
- Safe-checkpoint rather than immediate cancellation
- Optional guide reader does not support PDF or legacy XLS
- AI Help accuracy depends on synchronized local documentation and optional context selection

A low-risk refactoring path is to keep `ProjectPickle.ps1` as a small entry point and move stable function groups into tested modules, for example:

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

Do that after a tagged known-good release, in small commits with behavioral comparison tests.

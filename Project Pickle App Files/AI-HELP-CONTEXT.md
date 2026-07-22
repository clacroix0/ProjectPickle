# Project Pickle AI Help Context

Application version: 2.1.2

This file is the focused, bundled reference for the Project Pickle AI Help tab. It is not a substitute for the current how-to guide, the Build log, or review of the output database.

## AI Help Response Rules

When answering a Project Pickle question:

- Give practical instructions for the current Project Pickle workflow.
- Use the exact tab, field, button, table, and mode names shown here.
- Do not claim that AI Help changed the app, changed a database, ran a build, or verified an output file.
- Do not invent database contents, field mappings, provider installations, query results, or organizational policy.
- When the answer depends on the user's database or a run result, tell the user what to inspect in the app, output database, or detailed run log.
- Treat AI Help as advisory. Important outputs must be reviewed in Microsoft Access before production use.
- Never ask the user to paste an API key, complete project database, sensitive project data, or an unsanitized run log into a question.
- When information is not present in the bundled context, relevant code snippets, or selected guide, say what must be verified rather than guessing.

## What Project Pickle Does

Project Pickle is a Windows desktop utility for CFI Microsoft Access project databases. It can create a new project database from the current master database version, upgrade an older project database into the current structure, or update a current-format project database.

Project Pickle performs build work on an output copy. The selected current master database version or Existing/old project database should not be used as the Save Project Location. Existing output files are moved to a timestamped backup before replacement.

The version 2.1.2 tab order is:

Project, Periods, AppColumns, Upgrade Mapping, Plots, Get Field Order, Reports, CFIDEER Queries, Plot Assignments, Tree Assignments, Processing, Build, and AI Help.

## Starting Project Pickle

From the root Project Pickle folder, start the app with `Run Project Pickle.vbs`.

Use `Run-ProjectPickle-32bit.bat` for startup or Access-provider troubleshooting. Keep the root launchers and the complete `Project Pickle App Files` folder together because the launchers use relative paths. Regular users should not start `ProjectPickle.ps1` directly.

## Choosing a Mode

### New Project

Use New Project when there is no previous project database.

- Choose the Current master database version.
- Do not choose an Existing/old project database.
- Enter the project name and cycle interval.
- Load and select Region, Agency, and Reservation.
- Create the first measurement period and mark it current.
- Add plots, AppColumns, reports, queries, or assignments now only when those inputs are ready.

### Upgrade Project

Use Upgrade Project when an older database must be moved into the current standard structure or when a remeasurement is being prepared as part of an upgrade.

- Choose the Current master database version.
- Choose the old project database as Existing/old project database.
- Choose a different Save Project Location.
- Load existing periods, add the new period when needed, and mark the correct period current.
- Review AppColumns, Upgrade Mapping, Reports, assignments, and Processing options.
- Old plot, tree, regen, and assignment data import is normally enabled in this mode.

Upgrade Project is the only mode that uses both database inputs. Upgrade Mapping, Tree Assignments, and Processing are enabled only in Upgrade Project mode.

### Existing Project

Use Existing Project when the selected database is already in the current structure and only needs updates.

Typical Existing Project work includes adding custom fields, activating or deactivating AppColumns, adding new plots, changing Get Field Order, rebuilding CFIDEER queries, rebuilding plot assignments, or updating reports and validation.

- Choose the current-format project database as Existing/old project database.
- Choose a different Save Project Location.
- The current master database input is not used.
- Project name, cycle interval, Region, Agency, Reservation, and measurement periods are preserved from the selected database.
- Periods is read-only. Do not use Existing Project mode to add a new measurement period.

## Project Tab

The Project tab selects the mode, input database or databases, output path, and project reference information.

Required database inputs by mode:

- New Project: Current master database version.
- Upgrade Project: Current master database version and Existing/old project database.
- Existing Project: Existing/old project database.

Save Project Location must be a full `.mdb` or `.accdb` path and must not equal either selected input database path.

For New Project, click Load lists and select Region, Agency, and Reservation. Upgrade Project normally keeps those values from the old database. Existing Project preserves them from the selected current-format database.

The cycle interval is the planned number of years between full measurements. It is not the number of period rows.

## Periods Tab

Periods is used in New Project and Upgrade Project mode.

For a new project, create at least one positive whole-number PeriodNumber, usually period 1, and mark one row IsCurrent.

For an upgrade or remeasurement, load the old period rows, review them, add the new period when needed, and mark the new active period current. Only one period can be current.

Dates may be blank. A typed date must be valid. ApplicablePeriod is used when ValidationRules.ApplicablePeriods is updated.

Existing Project mode preserves current period rows and keeps the Periods tab read-only.

## AppColumns Tab

Use AppColumns for fields that are collected in the current measurement.

The upper grid loads existing or standard AppColumns from the database selected for the current mode. The lower grid is for new custom fields.

For a new custom field:

- Choose the matching plot, tree, or regen custom measurement table.
- Use a PascalCase field name with no spaces or punctuation.
- Choose the data type and category.
- Leave Active, ReportVisible, and QueryVisible checked for a normal collected field.
- For Category=Code, enter all allowed values in one Codes cell, such as `1=Yes; 2=No; 9=Unknown`.

Active controls whether the field is collected. QueryVisible controls whether an active field can be included in CFIDEER queries. ReportVisible controls whether an active field is available for reports.

Inactive history fields can remain physically present in custom tables without being included in current data-entry queries.

Before AppColumns changes, Project Pickle creates or refreshes AppColumnsArchive in the output database.

### Copy AppColumns from existing database

This option is available in Upgrade Project and Existing Project mode. It is a complete replacement, not a merge.

During Build, Project Pickle replaces the output AppColumns and AppColumnCodes rows with compatible rows from the Existing/old project database, then makes sure required physical custom-table fields exist. Manual AppColumns editing is locked while this option is selected.

Copying AppColumns does not copy plot, tree, or regen measurement records.

## Upgrade Mapping Tab

Upgrade Mapping is available only in Upgrade Project mode. Use it for old fields that must be preserved but do not map automatically into the current structure.

Use AppColumns for fields collected now. Use Upgrade Mapping for older history fields that should remain in the upgraded database but normally should not be active in current data-entry queries.

Basic workflow:

1. Choose the old database and current master database version on Project.
2. Create a new custom AppColumn first when the desired renamed target does not yet exist.
3. Click Load unmapped fields.
4. Click Check data types.
5. For each old field, choose a target table, an exact `Table.Field`, or Skip.
6. Save the mapping JSON only when it should be reused later.

A table-only target keeps the old field name and adds it to the selected compatible output table when needed. An exact `Table.Field` target maps or renames the old field into that field. Skip means the field is not copied.

Datatype mismatches must be resolved before Build. Do not force old text values into a current integer, date, Boolean, or GUID field without a valid crosswalk or conversion.

Upgrade Mapping copies data and can create a physical custom-table field. It does not automatically turn that field on in AppColumns.

## Plots Tab

The Add New Plots input accepts `.csv`, `.xlsx`, and `.xlsm` files.

After choosing a file, click Load columns and map PlotNumber to the source column that contains plot numbers. PlotNumber is required. Optional mappings include PlotLabel, PlotTypeID, UTM coordinates, UTMZone, and FLCCommercial.

The import is additive by PlotNumber. Existing plot numbers are skipped and new plot numbers are inserted. Fully blank spreadsheet rows are skipped, but a row with other mapped values still requires PlotNumber.

When required by the project period structure, Project Pickle creates missing blank PlotMeasurements and matching PlotCustomMeasurements rows for the new plots or period combinations.

In Upgrade Project mode, old inventory data import is separate from Add New Plots. After a database has already been upgraded, use Existing Project mode and Add New Plots for later additions rather than importing the old database again.

## Get Field Order Tab

Get Field Order controls display order for the three CFIDEER Get queries for plot, tree, and regen measurements.

Load fields after reviewing AppColumns. Project Pickle tries to preserve the selected database's saved query order and places newly added or newly activated Active plus QueryVisible fields at the bottom. Use Move up and Move down to arrange them.

Required remarks fields remain at the end: RemarksPlot, RemarksTree, and RemarksRegen. Criteria-only fields are not shown in this tab.

## Reports Tab

The Reports tab controls ReportColumns and fixed-width ReportHeaders layouts for the supported Plot, Tree, and Regen reports.

Click Load all from selected database before editing. New and Upgrade mode load from the current master database version. Existing mode loads from the Existing/old project database.

Use Add ReportVisible fields to add active report-visible custom fields to the matching layout. Place custom fields before the required remarks field. Use Move up and Move down to set order.

Generate starter header creates a starting fixed-width header. It must still be reviewed. Spaces are data. Use spaces rather than tabs, keep every nonblank line the same length, use printable ASCII, and keep each line at or below 255 characters.

Save current layout stores the selected layout in app memory. It does not write to a database until Build. Closing Project Pickle before Build loses unsaved in-memory report layouts.

Build chooses reports in this order:

1. Completed report-table copy, when selected in Upgrade Project mode.
2. Saved manual Reports-tab layouts, when Update reports is enabled.
3. The legacy automatic report updater, when Update reports is enabled but no manual layouts were saved.
4. The inherited base-database report setup, when Update reports is disabled.

Copy completed ReportHeaders and ReportColumns is an Upgrade Project option and is a complete replacement, not a merge. When copied reports reference custom AppColumns, copy or create the matching AppColumns too.

## CFIDEER Queries Tab

Build or refresh CFIDEER Get and Update queries is enabled by default. Leave it enabled for normal builds so queries match the final active and QueryVisible AppColumns.

The query rebuild changes only the output database, but it replaces the working saved query definitions without separate query backups. Review generated queries before production data entry.

Fields must be both Active and QueryVisible to be included as current custom data-entry fields. Upgrade-only history fields should normally remain inactive or not query-visible.

For exact query SQL, parameter order, or required trailing fields, use the current code context or technical guide. Do not infer an exact SQL definition from a general description.

## Plot Assignments and Tree Assignments

Plot assignment setup rebuilds InventoryCrews, InventoryAssignments, and InventoryAssignmentPlots for the current period. It clears existing assignment rows in the output copy before rebuilding to avoid duplicates.

Use plot assignments only when plots already exist or are being imported in the same build. If no plots are available, leave assignment setup off and run Project Pickle again later.

Tree Assignments is available only in Upgrade Project mode and requires tree rows. Tree assignments depend on plot assignments.

Rebuilt plot and tree assignment rows are expected to use the Awaiting Work status. Review final row counts in the detailed run log.

## Processing Tab

Processing is available only in Upgrade Project mode.

Import Processing tables from existing database replaces all output rows in these tables with rows from the old database:

- SpeciesMerchanabilityStandards
- SpeciesNationalVolumeParms
- SpeciesRegressionCriteria
- SpeciesRegressions

This is a complete replacement, not a merge. Leave it off when the upgraded output should use the current master database version's Processing setup.

## Build Tab and Output Review

Before Build, correct every validation message. Build does not begin database changes when required entries are missing.

The build copies the selected base database to the output path, then applies the selected project, periods, AppColumns, upgrade import, Processing, plot, assignment, report, validation, and query operations to that output copy.

An existing output file is moved to a timestamped `.backup_yyyyMMdd_HHmmss` file before replacement.

After a successful build, review the output database and run log. Check at least:

- Project, Region, Agency, and Reservation values
- Measurement periods and current period
- Plot, tree, and regen row counts when imported
- Matching custom measurement rows
- AppColumns and AppColumnCodes
- AppColumnsArchive
- ReportHeaders and ReportColumns
- CFIDEER query definitions
- Assignment row counts and statuses
- Processing table counts when copied

## Run Logs and Troubleshooting

The Build tab is the live progress view. Project Pickle writes a detailed text log beside the output database for both successful and handled failed builds. If no output folder is available, a fallback log may be written near the app files.

A log name resembles:

`ProjectName.ProjectPickleRunLog.Success.20260608_123456.txt`

or

`ProjectName.ProjectPickleRunLog.Failed.20260608_123456.txt`

The log includes step numbers, timestamps, categories, elapsed time, step duration, mode, input and output paths, and detailed messages.

For a build problem:

1. Read the popup and the last Build-tab messages.
2. Open the detailed run log named by the popup.
3. Find the first error and the preceding step, not only the final cleanup line.
4. Review row-count and relationship diagnostics during upgrades.
5. Remove sensitive information before sharing a log.

Large Access writes can temporarily make Windows show Not Responding. If new log lines continue to appear, the build is still making progress. Cancellation waits for a safe checkpoint and cannot safely interrupt every active Access write.

## AI Help Data and Limitations

AI Help does not modify the database. A request can include the user's question, this bundled context file, relevant source-code snippets when enabled, and relevant text from an optional selected guide or reference file.

Supported optional guide types are DOCX, XLSX, XLSM, TXT, MD, CSV, JSON, and PS1. PDF and old XLS files are not supported by the in-app reader.

The API key is entered at run time and is not intended to be saved. Use only an organization-approved HTTPS Azure OpenAI endpoint. Do not include sensitive project data unless it is approved for that service.

AI Help answers are advisory. When an answer conflicts with the current app screen, source code, how-to guide, or run log, use the current application evidence and ask a maintainer to update this context file.

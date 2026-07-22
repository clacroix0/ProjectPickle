# Project Pickle Codex Thread Backup

Saved: 2026-06-04

Project folder:

`C:\Users\christopher.lacroix\OneDrive - DOI\Documents\Database apps\Project Pickle`

Codex thread:

- Thread title: Build inventory app from DADA
- Thread ID: `019e83ab-0995-7b03-bdd6-39b5f5fe7619`
- Workspace: `C:\Users\christopher.lacroix\OneDrive - DOI\Documents\Database apps`

Important note:

This is a readable local backup of the prompt thread context visible to Codex when this file was created. It is intended to preserve the project story, user requests, decisions, and implementation state if Codex access is lost. It is not a byte-for-byte platform export of every hidden instruction, tool call, or large command output.

## Current App State

Project Pickle, short name WOOF, is a Windows PowerShell WinForms app for building and upgrading CFI inventory Access databases.

Main files in this folder:

- `WOOF.ps1`
- `Run-WOOF-32bit.bat`
- `WOOF.vbs`
- `WOOF.lnk`
- `ProjectPickle.ico`
- `README.md`
- `Project Pickle WOOF How-To Guide.html`

Current visible app build in `WOOF.ps1`:

- `Project Pickle (WOOF) - Build 2026-06-04.22`

Recent verification performed:

- PowerShell parser check passed.
- Full WOOF self-test passed.
- UI smoke test passed.
- Upgrade moved-field bulk mapping was exercised by self-test.
- Upgrade datatype pre-check was added and self-test still passed.

## Latest User Request Before This Backup

User asked:

> save a copy of this whole prompt thread in the project folder on my computer in case I lose access to codex over the weekend

This file is that backup.

## Key Implementation Outcomes

### App Identity

- App named Project Pickle.
- Short name WOOF.
- Branding uses pickle icon for shortcut/window where PowerShell allows it.
- Author/credit added:
  - Christopher LaCroix, USDI BIA Division of Forestry, Branch of Inventory and Planning.
  - In-app credit line: Developed by the BIA Division of Forestry, Branch of Forest Inventory and Planning.

### Project Setup

- Master/existing database selection.
- Upgrade mode uses old database as source plus a clean master template as output base.
- Output database naming hint preserves the selected template structure, such as:
  - `[TribeName] CFI [Date 20260520] v3.06.00.mdb`
- Version number is not meant to be baked into code.
- Access application title is updated from placeholders:
  - `[TRIBENAME]` becomes project name.
  - `[DATE]` becomes file creation date in `yyyyMMdd`.

### Region, Agency, Reservation

- Region, Agency, Reservation name are selection fields.
- Lists currently load as full lists rather than key-filtered cascading lists, because filtering caused repeated load issues.
- Reservation name is the preferred selection display rather than reservation code.
- Selecting agency, region, reservation can prune non-applicable reference rows.

### Periods

- No limit to one or three periods.
- New projects commonly start with one period.
- Existing/remeasurement projects can load multiple periods.
- Periods tab has horizontal scroll support.
- Only one `IsCurrent` period can be checked at a time.
- New manually added period rows become current automatically.
- Help text explains checking the new period when adding a new period to a multiperiod project.

### Plots And Assignments

- Plot data is optional.
- Plot and assignment setup can be skipped and run later.
- Plot assignment and tree assignment were split into separate tabs.
- Tree assignments are for remeasurement projects.
- Plot imports can map spreadsheet columns to WOOF fields instead of requiring a standard spreadsheet format.
- Removed plot import options that are not currently needed:
  - ManagementUnit
  - Compartment
  - Elevation
  - Aspect
  - SlopePercent
  - SlopePosition
- FLCCommercial was added to plot import support.
- For existing/upgraded projects with previous measurements, WOOF creates blank PlotMeasurements rows for new/missing plot-period combinations and matching PlotCustomMeasurements rows with the same MeasurementID.
- For new plots added to a project with previous periods, WOOF can back-create blank plot records for prior periods.

### AppColumns

- Tab renamed to AppColumns.
- Existing standard AppColumns load into a separate review area from new custom AppColumns.
- Existing AppColumns can be searched by field name.
- Master filters were added for active/non-active and plot/regen/tree data.
- Existing AppColumns grid is taller to show more rows.
- Standard fields that should never be toggled are hidden from review lists.
- Standard code moved next to Active checkbox in loaded AppColumns.
- Standard code removed from new custom AppColumns.
- Active now implies collected; ArchiveOnly was hidden/removed from normal workflow.
- ReportVisible and QueryVisible default to checked for new active fields.
- If a new field is clicked Active, it also activates ReportVisible and QueryVisible.
- Codes entry supports long code lists and an editor. Multiple codes can be entered like:
  - `0=None; 1=Trace; 2=Light; 3=Moderate; 4=Heavy; 5=Severe`
  - or one code per line.

### Upgrade Mode

- Upgrade mode copies the clean master template first, then imports compatible data from the old database.
- AppColumns in Upgrade Project mode load from the clean master template, not from the old project database.
- Project periods can be loaded from the old project database.
- Old plot, tree, regen, and assignment data can be imported into the newer output structure.
- Missing parent inventory rows are repaired when enough information exists:
  - missing Plots from PlotMeasurements/RegenMeasurements/TreeMeasurements
  - missing Trees from TreeMeasurements
  - missing custom measurement rows
- Repaired rows use `UpdUser=Pickle` and IDs are written to the run log.
- A field mapping tab supports mapping old fields to a target table or exact table.field.
- Mapping copies values and creates physical fields as needed; it does not activate AppColumns.
- NeedsReview fields are skipped automatically.
- Automapping rules include:
  - TreeRemarks -> TreeMeasurements.RemarksTree
  - PlotRemarks -> PlotMeasurements.RemarksPlot
  - Remarks -> RegenMeasurements.RemarksRegen
  - Plots.HabitatType -> PlotMeasurements.HabitatType
  - FieldOverCover -> PlotMeasurements.CoverType
  - FieldOverDensity -> PlotMeasurements.DensityClass
  - FieldOverSize -> PlotMeasurements.SizeClass
  - ForestClassification -> Plots.FLCCommercial
  - unmapped old TreeMeasurements fields -> TreeCustomMeasurements
  - unmapped old Plots/PlotMeasurements fields -> PlotCustomMeasurements
  - old custom measurement fields -> matching new custom measurement table

### Upgrade Datatype Check

Latest implemented feature:

- Upgrade Mapping tab now has `Check data types`.
- Pre-build validation also checks datatype changes.
- It checks:
  - same-table imports
  - automatic moved fields
  - manual Upgrade Mapping rows
  - active custom fields from AppColumns
- It flags risky changes such as old Text values going into new Integer, Date/Time, Boolean, or Guid fields.
- If actual old values will not fit, Build stops before copying/changing the output database.
- The fix guidance is:
  - clean the old values,
  - choose a Text/Memo target,
  - or map to a new custom field that uses the old field's data type.

### CFIDEER Queries

- CFIDEER query builder is optional and off by default.
- It has its own tab because it is powerful.
- It can be run later in Existing Project mode after AppColumns/custom fields are settled.
- It rebuilds saved Access queries whose names start with `_prjCFIDEER_`.
- It backs up existing matching queries before replacement.
- Query builder avoids table aliases like `pm`; it keeps original table names.
- PARAMETERS format was corrected to match examples without unnecessary brackets around parameter names in the parameter list.
- Custom fields are inserted into get queries after RemarksPlot, RemarksTree, or RemarksRegen where appropriate.
- FieldNameSampleRemove is removed from get/process/custom tables and query output.
- Query builder now carries AppColumns datatype into query parameters.

### Reports

- WOOF updates ReportHeaders and ReportColumns from report-visible custom fields.
- Report setup is tricky and was adjusted several times:
  - Plot custom fields should create/add a custom plot report line/area.
  - Tree custom fields should be formatted into tree measurement report between RAD INC and Remarks.
  - Regen custom fields should be added between STEM Count and Remarks.
- Known caution: report layout logic is one of the more fragile areas and should be reviewed with real project outputs.

### Run Log

- WOOF writes a `WOOFRunLog` table into the output database.
- It contains:
  - RunID
  - StepNumber
  - LogTime
  - RunStatus
  - StepCategory
  - ElapsedSeconds
  - StepDurationSeconds
  - ElapsedTime
  - StepDuration
  - Mode
  - ProjectName
  - SourceDatabase
  - OutputDatabase
  - Message
- It records total run time and process timing.
- For moved-field bulk mapping, it logs summary counts, source/target table, field names, and duration.
- It does not log every single moved MeasurementID/value because that could make the log enormous.
- Repair steps log individual IDs when creating missing Plots/Trees/custom measurement rows.

### Performance

- Upgrade moved-field mapping was sped up.
- WOOF now tries bulk Access table-to-table operations first.
- If Access rejects the bulk path, it falls back to the slower row-by-row method.
- Bulk moved-field path was verified in self-test.

### AI Help

- Optional AI Help tab was added.
- It does not change the database.
- It can use WOOF README/code snippets and optional guide/reference snippets.
- It is intended for "how do I do this" user support.

### Documentation

- `README.md` updated repeatedly.
- `Project Pickle WOOF How-To Guide.html` created as a Word-friendly readable guide.
- Requirement: every new update should review and update the guide.

## Chronological User Request History

The following list preserves the visible user prompts and requests in order as available in this thread context.

1. I want to build an app that set up inventory database from scratch, reference my DADA dad app

2. make the plot information optional as that sometime isnt know when setting up the database. Add an option to build out the get and update queries _prjCFIDEER_ as these take a while to set up and are prone to human errors espcially the update queries. Refererence the CFI project guide on how these are built. How will this app be use for exisitng projects? as existing projects are common. Do not limit period to 1 or 3 period. New projects are only set up for one period, but existing projects need to be set up for multiple periods depending how how many previous measurements there are

3. can the CFIDEER Query Builder be run whenever? In theory, the custom tables need to be set up before running the CFIDEER Query Builder so all the current project fields are added to the queries, maybe this is an option where it can be built in the app? Say theres a add button and uses have to put the name of the field, and the data type, this will need to be done seperately for plot custom measurements, tree custom measurements, and regen custom measurements. Sometimes fields need to be removed from the queries (as noted in the user guide), so users need to be able to verify what field they want to remove, maybe that can be done in the add/remove list?

4. What about setting up the inventory assignment tables? The inventory assignment plots table can only be set up once plot information is added. inventory assignment Trees can only be set up on remeasurement projects not new projects so this needs to be option. What about set up of the Report columns and headers tables? this needs to be done once custom fields are added/removed to the measurement tables. Lastly, the validation rule table needs the applicable period date change to the current project measurement period date. By entering the agency, region, and reservation - this should clear out all non applicable data in those tables to that agency, region, and reservation. Make sure that Project ID in all applicable tables uses the correct Project ID from the project table - this needs to be set up for new projects. Is there an easy way to aid the user in filling out the app columns table?

5. implement the plan

6. PLEASE IMPLEMENT THIS PLAN:
   - Add `InventoryDatabaseBuilder.ps1`
   - Add `Run-InventoryDatabaseBuilder-32bit.bat`
   - Update README
   - Reuse helper patterns from CFICodeCrosswalk
   - Build wizard sections for source/output, mode, project setup, periods, plot import, custom fields, AppColumns assistant, reports, validation, assignments, CFIDEER query builder
   - Always copy source first and modify only output copy
   - Add database routines for copy/backup, ProjectID propagation, pruning, periods, optional plots, assignments, custom columns, AppColumns/AppColumnCodes, ReportHeaders/ReportColumns, ValidationRules, CFIDEER backup/rebuild
   - Add validation reports before destructive changes
   - Run parser check, UI smoke test, and output database tests

7. put this app and all its files it its own project folder

8. Reservation name should be a drop down select with a search field. same with reservation code - but make it the ReservationName instead of Reservation Code, add Region as a selection too. period, plots and assignments, and custom fields need a scroll bar at the bottom to scroll left and right. what is table name in custom fields? this tab should be call app columns

9. remove any old testing files from the folder

10. lets call it Project Pickles, short name WOOF since pickles is the new name of my dog, and woof it sure does make database set up easy

11. sorry Project Pickle*

12. is there a way to have the power shell window hid inside of being shown when the app lauches? give the new launcher a simple name

13. can you make a shortcut to the .vbs but change the icon to a pickle icon so it looks more like an app launcher

14. the woof shorcut is outside of the project folder! same with the readme

15. window/taskbar icon needs to be a pickle too

16. add warning pop up about closing the app if it is still processing

17. add the ability to open the output after it successfully runs. I dont want to navigate to where it was saved to open it. include the run log as an table export.

18. is it normal to add the author to the readme file? IF so add author as Christopher LaCroix, USDI BIA Division of Forestry, Branch of Inventory and Planning

19. ok when I go to enter data for app columns tab, and I add a data entry, for the codes how do I enter multiple codes correctly? also it runs a little slow, any way to speed it up?

20. can you add this as tool tips so people know how to enter data into the appcolumns tab?

21. remove the detault text from master/existing database

22. is cycle length number of measurement periods? if so rename it to make it more obvious

23. so what do I put here for a new project or remeasurement project? I need a tool tip, why do I need this for creating a database?

24. why do I need this to build the database though?

25. walk me through the plots and assignments tab, I dont have any plot data yet, can it be made clear what this tab does and it can be completed once plot data is loaded into the database?

26. can you add an optional AI feature to the app like my other apps so users can ask how to do certain things? AI model can analyze the code and also the CFI user guide an any other useful info

27. Similar to the plots and assignments table, I need some user direction for the appcolumns table!

28. period and appcolumns needs a scroll bar at the bottom, when the app isnt maximized I cant scroll on the screen

29. On the main project tab can you add a some instruction on what this app is and what is does noting the optional AI Assistant Tab

30. what is the build/refresh CFIDEER get and update queries, how does it work and it is it off by default?

31. can explination be added to this window for that, maybe it should be its own tab since its so powerful and can be done after the fact

32. add a tool tip for periods, and for the check boxes for update validation rules projectid and applicable period, report headers and report columns

33. the text in measurement cycle text is cut off, this also needs a tool tip

34. can Master/existing database have a tool tip so users when to know which one to use since the same location uses both

35. the text on the queries tab is broken up with big spaces, can it be brought all together?

36. in app columns, what does collected, archive only, report visible, and query visble do?

37. can you add all of this to the info on that tab? also the app still needs to be in full screen to see everything can you fix that?

38. can the app update the application title found in the options of the access database? [tribename] and [date] should be replaced with the project name and the date the file is created

39. agency isnt a selection box on the project tab so how does it know the correct agency?

40. hmm now the app going outside of the app window!

41. the code entry on appcolumns needs a better sample on how to enter multiple codes for one field, some field have 10+ codes in them

42. I got this error, what does it mean? if its a missed data entry field can you provide clear instructions on how to address the error and all errors? Error: Syntax error in INSERT INTO statement.

43. why did access deny it? can you fix it so that doesnt happen?

44. did you test it yourself?

45. signing out for the day anything else I should have you test and fix? note I lose access to codex tomorrow so I need this project to be perfect

46. on the app columns tab, I need to be able to see the standard fields loaded so I can uncheck active ones that dont pertain to my project

47. can the standard rules be in a different box than the custom rules that are being added so they are easier to tell apart? are the standard rules loaded in when you load in the database? Can custom rules that already have standard codes be preloaded when a user decides to use that rule and codes? Users also need the option to create true standard rules on the fly too

48. ok actually I dont need the ability to Use selected loaded rule, but I do like the idea of a "copy" button so I can copy the format of a rule such as the table it goes, data type, and category. The "loaded standard rules window needs to be bigger with a vertical scroll bar

49. Do I really need the archive only button? Report visible and Query visble need to be on by default? Do I need and active and collect? if a field is active that means it is collected. Lastly I need a master filter that allows users to only show active fields and lets users only show non active fields. The helper text that starts with "active" goes off the app screen

50. move standard code on the exsiting app column for review next to the active check box. Remove it entirely from the "new custom appcolumn to add"

51. when I load an existing project, all the periods need to load with that project too in the periods table

52. how come its not loading planned end data actual start date and actual end date?

53. instead of clicking load lists, can it auto load when the database is selected? less clicking

54. I just got this new error, you broke it! The term `Start-SourceDatabaseAutoLoad` is not recognized.

55. did you fix it?

56. the icon for the app when its running is a powershell icon, is it possible to change it to the pickle icon like the rest of the app?

57. now I got this error, but if I hit continue the app still works. You cannot call a method on a null-valued expression. I guess you can roll back the auto load function because its not working

58. now I get this error when trying to load lists The term `Set-ComboItems` is not recognized.

59. ugh it semi works, i can load list and after I select a region I get this error: You cannot call a method on a null-valued expression.

60. ok now after I select a region, Agency and Reservation lists are empty! make it work like before!

61. its still not loading agency and reservation after I select a region

62. ugh its still NOT working for agency and tribe

63. I dont understand, its still not working at all, only region populates

64. now I get the error "couldnt load lists from the selected database"

65. seems like you are having issue with the smoke test, can you run it another way?

66. i get the error when clicking load "you cannot call a method on a null valued expression"

67. ok woo it finally loads again? no way to reconnect region > agency > tribe to filter by the keys?

68. ok woo it finally loads again? no way to reconnect region > agency > reservation to filter by the keys?

69. that broke it again! Just give me the full lists again without the filtering

70. how come its changing the table names in the queries from say PlotMeasurements to pm? it should keep the orginal table names

71. the custom data didnt get added to the reportheaders properly, this might be a harder one to code in as it is a bit tricky

72. so will it put new entries in the correct report area?

73. why is it putting all the names in brackets? Query PARAMETERS should stay in the original format without every parameter name bracketed. New custom fields need to carry over the data type from AppColumns. FieldNameSampleRemove Short needs to be removed from queries.

74. this is an example how a report header table with custom fields added should look like

75. its still incorrect putting all the names in brackets ex [SpeciesCode], it should just be SpeciesCode Short

76. when its adding codes to app columns its not creating a new line for each code, its making them all the same

77. lets drop forest cover type from the Site tab in the logger app for the time being, we can revisit that one once we have more access to these tools

78. its still doing the brackets, the one on the left is the wrong format, the one on the right is the correct format

79. can you not load the following tables when appcolumns is loaded since these will never get turned on of off: plots, plot summaries, regencalculations, treecalculations, and trees

80. also add a filter similar to show all fields that lets user filter plot data regen data for tree data fields

81. you can also hide TreeID, TreeKey, TreeMeasKey, ReadDBH, TreeMeasID, RegenMeasKey, ProjectPeriodID, Plot number, MeasurementID, GMP, LastGrowingSeason, CalcsiteIndex in all tables they appear in

82. queries are looking good! except for the getplot, getregen, and gettrees query, custom items should go after remarksplot, remarks tree, and remarks regen. Do this for these tables only. You dont need to create a backup of the old tables. When user clicks to output the data base can it keep this naming structure as a naming hint? `[TribeName] CFI [Date 20260427] v3.05.00`

83. why did it clear out the get process tables when it added the custom field? It shouldnt do that, it should only add the new custom field at the very end of its respective get process table

84. also its still getting the reports wrong; custom plot info needs a new line, tree custom item between RAD INC and Remarks, regen between STEM Count and Remarks

85. we are mostly there! it didnt remove the FieldNameSampleRemove from the getprocess tables, only remove that field. Also, FieldNameSampleRemove needs to be removed from treecustom, plotcustom, and regencustom tables

86. ok lets say I open up a database in our old database structure, does the app know to produce the new project database in the ver 3.05 structure?

87. are you able to set that up?

88. remove the following from showing in loaded appcolumns data from the plots measurement table: crew, measurement date, period number, plotid, plotkey, plotmeaskey, plot status, remarks plot, stand class, stand age, stockability factor, stockabilitypct. From regenmeasurements remove IDBH, stemcount. From tree measurements remove crown class, crown ratio, IDBH, period number, remarks tree, total height, and tree history. For long code lists like problem 1, I currently see all items and need to be able to see all of them to make edits

89. so is it built on only using 3.05 as a template, what if we get database ver 3.06 will it still work?

90. ok I am testing the upgrade mode, 1st, the period data didnt load when I hit the load button on the project tab, second when I tried to load app column data I got this error. The app column data should always be loading from the master template

91. is that only in upgrade mode? you can actually just load the Region / Agency / Reservation from the old project database

92. ugh it still didnt load the project periods from the old database, does that tab needs its own load button? Also, dont have any default text in the clean master template input

93. i got this error when trying to load periods: The term `Add-PeriodRowToGrid` is not recognized

94. now I get this error again when trying to load data into appcolumns: No value given for one or more required parameters.

95. there also needs to be a feature to load in all the plot, tree, and regen data from an old database so it can added to the correct tables in the new database but also loaded into the inventory assignment tables. How does the app handle conversion of a field being in tree measurements in old version but now in tree custom measurements in new version? Would it be best not to load in this data to break the app?

96. can you add a search field to appcolumns so users can search the existing app columns by field name

97. make a seprate tab for plot assignments and tree assignments

98. for plot data only blank plot measurement records and matching custom measurement records with same measurementID need to be added for the new period. If old database has 5 measurements and I am setting up a 6th, add blank 6th measurement plot record for all plots with MeasurementID, PlotMeasKey, PlotID, PlotKey, ProjectPeriodID, PeriodNumber, UpdUser=Pickle, rest blank

99. which tab imports the regen data?

100. so when its adding blank records, how does it create the PlotMeasKey, PlotID, PlotKey?

101. perfect! what if I need to load in a list of new plots to be added?

102. add FLCCommercial to that list

103. so if I have a 5 measurement period project and need to load in 10 new plots for the 6th measurement, will it back create blank plot records for periods 1 - 5?

104. for uploading plots, so it doesnt have to be a standard form, can it be when someone uploads a spreadsheet they have to connect the field name to the column in the spreadsheet?

105. you can remove these from options for the plot import: ManagementUnit, Compartment, Elevation, Aspect, SlopePercent, SlopePosition

106. what is the is current check box for in the period tab?

107. I am running an upgrade and it looks stuck/frozen on upgrade import note. Can you add in something that stops it from freezing like this

108. its still slowly moving along, what does NeedsReview mean?

109. so does it looks stuck to you? it doesnt look like its making much progress

110. the app failed with run log export error: field too small to accept amount of data while adding a row to WOOFRunLog. Would it be easier if I mapped fields from old database version to correct tables in new database version for fields that dont auto map?

111. can we add the field mapping option in

112. reviewing the database I just upgraded, plots and plot measurement mostly look good, but existing data did not make it into custom measurement tables besides new test custom fields. Tree data for 5th measurement did not make it in. Maybe mapping feature will fix this?

113. added this clean credit line under the header subtitle in the app: Developed by the BIA Division of Forestry, Branch of Forest Inventory & Planning.

114. can you create a reader friendly how to guide on how to work and navigate the app written so a freshman in collage can understand it without reaching out to developer? Every time new updates are added, review guide and keep it up to date. Produce an html file formatted with clean headers, section breaks, bullets, etc. when opened in Microsoft Word.

115. the & sign in the credit line didnt come out. Change to Developed by the BIA Division of Forestry, Branch of Forest Inventory and Planning.

116. mode dropdown text "upgrade old project to selected" gets cut off, change to "upgrade project". In most tabs text is cut off horizontally; fix by containing text in smaller area. In AppColumns tab make existing AppColumns take up more vertical screen space.

117. update suggested name in output database saver to `[TribeName] CFI [Date 20260520] v3.06.00`

118. latest database is 3.06 now so make documentation say 3.06, but do not bake version into app.

119. in mapped fields I need to map what table the field goes to, not exact field only, because custom fields likely wont have an existing field and need to be added to table.

120. include plots, trees, plot measurements, and tree measurements; put plot custom measurements first then tree custom measurements second. Automap old plot custom to new plot custom, same for trees. What does save button do? If I dont save will it still bring in all my data?

121. where do NeedsReview fields come from? They are in my table? They should be auto skipped and not loaded. Automap TreeRemarks, PlotRemarks, Remarks. TreeMeasurements unmapped to TreeCustomMeasurements. Plots and PlotMeasurements unmapped to PlotCustomMeasurements. Exclude many standard fields from dropdown.

122. habitat type from plots should always map to PlotMeasurements.HabitatType

123. fieldOverCover -> PlotMeasurements.CoverType, FieldOverDensity -> PlotMeasurements.DensityClass, FieldOverSize -> PlotMeasurements.SizeClass, ForestClassification -> Plots.FLCCommercial

124. if a new field is clicked active in AppColumns tab it also activates report visible and query visible

125. when a new IsCurrent period is checked, any other selections uncheck, and only one period selected at a time. Add text that says to check new period when adding a new period to multiperiod project. When period is added by user, check itself.

126. in the run log will it contain how fast each process took and total run time? what else?

127. lets do it, its good to know the run details

128. looks like it takes a while to copy over records, why is this, any way to speed it up?

129. got error: Upgrade Mapping tab: Plots.HabitatType cannot map to PlotMeasurements by table name because tables use different row keys. Fix and make sure other fields dont do this.

130. what does this error mean? Just because a field is mapped doesnt mean it needs to be turned on in AppColumns.

131. The mapping just make sure that all fields and data make it into the new database, not set up AppColumns tables

132. is it loading AppColumns from the database template or from old database?

133. can you verify? I am not sure that is the case

134. its for sure loading AppColumns from master/existing database; it needs to load from clean master template. Reverify code.

135. almost there! still not bringing all tree data over from 5th measurement for project I am testing.

136. So this database has trees in the tree measurement table not in the trees table because someone screwed up data entry, could this be the issue?

137. create missing Trees rows when enough info exists, such as TreeID, TreeKey, PlotID, PlotKey, and tree number. Do same for plot and regen records because it could happen to all. Note IDs in log. Maybe add Pickle as updated user.

138. instead of suggesting output name GuideChecks_20260603_1150, change to `[ProjectName]_DADACleaning_` plus date/time.

139. Great now lets call it version 2.0 since we are done with testing!

140. this should be pickles not validude?

141. that was an accident, roll pickle back to 2026.06.04.22

142. the "upgrade moved field mapping takes a very long time, anyway to make this faster?

143. is it writing all of the upgraded moved to the log file in the database??

144. I have a new issue, some data isnt writing to the mapped field because the old field was say text, and the new field is say integer. How can I have the app catch when fields change datatypes before I run it?

145. save a copy of this whole prompt thread in the project folder on my computer in case I lose access to codex over the weekend

## Recent Codex Answers And Outcomes

### Upgrade moved-field speed

Codex implemented a bulk moved-field path:

- WOOF now creates a temporary source snapshot from the old database.
- It bulk-creates missing target/custom measurement rows.
- It bulk-updates moved fields.
- If Access rejects bulk SQL, WOOF falls back to the old row-by-row method.
- Documentation updated.
- Parser check passed.
- WOOF self-test passed.

### Run log behavior for moved fields

Codex explained:

- WOOFRunLog logs moved-field work as summaries, not one row per MeasurementID/value.
- Summary includes source table, target table, count of rows copied, fields moved, duration, and missing target rows created.
- Repair steps log individual IDs because those are important to find later.

### Upgrade datatype check

Codex implemented:

- `Check data types` button on Upgrade Mapping.
- Pre-build data type validation.
- Stops build before output is copied/changed when old values cannot fit the new target type.
- README and How-To Guide updated.
- Parser check, full self-test, and UI smoke test passed.

## Files Most Recently Updated

- `Project Pickle\WOOF.ps1`
- `Project Pickle\README.md`
- `Project Pickle\Project Pickle WOOF How-To Guide.html`
- This backup file.

## Known Local Temporary/Test Files

Some old `.mdb` and `.ldb` self-test/probe files remained in the Project Pickle folder because Access/OneDrive/Windows locks denied deletion earlier:

- `_inventory_builder_selftest_6538f9a9aa764b22affca76d5a8a4d3e.mdb`
- `_inventory_builder_selftest_6538f9a9aa764b22affca76d5a8a4d3e.ldb`
- `_inventory_builder_selftest_6cfea80545b74332aad40ca6ec0f9c76.mdb`
- `_inventory_builder_selftest_6cfea80545b74332aad40ca6ec0f9c76.ldb`
- `_reference_list_test_copy.mdb`
- `_reference_list_test_copy.ldb`
- `_woof_schema_probe.mdb`
- `_woof_schema_probe.ldb`

Do not rely on these as production databases. They are test/reference artifacts.

## If Work Continues Later

Recommended next checks:

1. Run a real upgrade using a copy of a known old project.
2. Use Upgrade Mapping -> Load unmapped fields.
3. Use Upgrade Mapping -> Check data types.
4. Review any datatype warnings/errors before build.
5. Build output copy.
6. Open output database.
7. Review WOOFRunLog, ProjectID consistency, period rows, imported plot/tree/regen counts, custom measurement data, ReportHeaders/ReportColumns, and CFIDEER queries.


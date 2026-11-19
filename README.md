# CAG CHRONLOG v2.0
## Administrator User Manual
**Version:** 2.0  
**Release Date:** November 19, 2026  
**Platform:** Microsoft Excel for Windows  
**File Format:** Excel Macro-Enabled Workbook (.xlsm)  
**Intended Users:** Site Security Managers (SSMs), and Authorized Administrators

---

# 1. PURPOSE AND OVERVIEW

CAG ChronLog v2.0 is an Excel-based chronological security logging system designed for use by Cleared American Guards (CAGs).

The system provides a structured method for documenting security activities, observations, incidents, and other events requiring chronological records.

ChronLog incorporates automatic timestamps, entry locking, worksheet protection, automated navigation, and controlled worksheet visibility to reduce accidental modifications and improve consistency.

The system is intended to operate locally using Microsoft Excel without requiring an internet connection.

## 1.1 Primary Objectives

- Maintain accurate chronological security records.
- Automatically timestamp log entries.
- Prevent accidental modification of completed entries.
- Standardize log-entry procedures.
- Reduce repetitive manual tasks.
- Preserve existing entries during daily operations.
- Allow authorized administrators to correct entries when necessary.
- Maintain separate daily log files for review and retention.

## 1.2 Terminology

**CAG:** Cleared American Guard or Cleared American Guards.

**SSM:** Site Security Manager.

**ChronLog:** The Excel-based chronological logging system.

**Template:** The original blank workbook used to create daily logs.

**Daily Log:** A separately saved ChronLog file containing entries for a particular date.

**Macros:** Microsoft Excel VBA instructions responsible for automatic timestamps, navigation, cell locking, and row management.

---

# 2. SYSTEM REQUIREMENTS

ChronLog v2.0 requires:

- A Windows computer running Microsoft Excel with VBA support.
- An Excel version capable of opening `.xlsm` files.
- Permission to enable approved Excel macros.
- Access to the original ChronLog template.
- A designated storage location for daily log files.

The workbook does not require an active internet connection.

ChronLog is designed for desktop Microsoft Excel. Functionality is not guaranteed in Excel Online, alternative spreadsheet programs, or mobile spreadsheet applications.

---

# 3. WORKBOOK STRUCTURE

ChronLog contains two Excel sheets.

## 3.1 Sheet1 — Chronological Log

Sheet1 is the primary logging interface.

The sheet contains three operational columns:

| Column | Purpose |
|---|---|
| A | Automatic date and time |
| B | Log description |
| C | Last name and initials |

Rows 1 and 2 contain instructions and column headings.

Log entries begin on row 3.

The configured logging range extends through row 10,000.

Columns beyond C are hidden to reduce unnecessary navigation and maintain a consistent interface.

## 3.2 Enable Macros — Startup Instructions

The Enable Macros sheet explains that macros are required for normal ChronLog operation.

When the workbook is saved in its startup configuration, Sheet1 is hidden and the Enable Macros sheet is displayed.

After macros are enabled, Excel automatically displays Sheet1 and hides the instruction sheet.

This feature helps prevent CAGs from accidentally entering information without the automation required for proper logging.

---

# 4. DAILY LOG CREATION

## 4.1 Beginning a New Shift

At the beginning of each new logging period:

1. Open the original blank ChronLog template.
2. Enable macros if Excel requests permission.
3. Verify that Sheet1 opens.
4. Select File → Save As.
5. Choose the Windows Desktop as the destination.
6. Enter the date as the filename.
7. Save the new file in `.xlsm` format.
8. Confirm that the new daily file is open before beginning log entries.

**Recommended filename format:**

`YYYY-MM-DD.xlsm`

Example:

`2026-11-19.xlsm`

This naming convention allows daily logs to be sorted chronologically by filename.

## 4.2 Preserving the Original Template

The original blank template must not be used for daily entries.

Each shift should begin by creating a separate daily file using Save As.

The original template should remain blank and available for subsequent shifts.

The administrator should maintain a backup of the original template.

## 4.3 Important Save As Requirements

Daily files must retain the `.xlsm` extension.

Saving as `.xlsx` removes the VBA functionality required for automatic timestamps, entry locking, navigation, and row management.

---

# 5. ENTERING LOG INFORMATION

## 5.1 Standard Entry Procedure

1. Click the next available cell in column B.
2. Enter the chronological log description.
3. Press Enter or Tab.
4. Excel automatically moves to column C.
5. Enter the CAG's last name and initials.
6. Press Enter or Tab.
7. Excel automatically moves to the next available cell in column B.

The next row becomes available after the current row has been completed.

## 5.2 Automatic Timestamps

When information is first entered into an eligible log-entry cell, ChronLog automatically records the current system date and time in column A.

The timestamp uses the computer's local system clock.

The configured format is:

`MM/DD/YYYY HH:MM:SS`

Example:

`11/19/2026 23:45:12`

Timestamps are generated automatically and are not intended to be entered manually.

## 5.3 Timestamp Restrictions

Cells A3 through A10000 are locked.

CAGs cannot manually modify timestamps through normal protected-sheet operation.

An existing timestamp is preserved when the corresponding row receives additional information.

Administrators should ensure that the computer's date, time, and time zone are correct before beginning a shift.

ChronLog does not independently verify the accuracy of the Windows system clock.

---

# 6. AUTOMATIC CELL LOCKING

ChronLog automatically locks populated log-entry cells.

## 6.1 Column B

After information is submitted in column B:

- The entered information is retained.
- The cell becomes locked.
- The cursor advances to column C.

## 6.2 Column C

After information is submitted in column C:

- The entered information is retained.
- The cell becomes locked.
- The next available log row is revealed.
- The cursor advances to column B of the next available row.

## 6.3 Empty Cells

Empty entry cells remain unlocked until information is submitted.

This allows CAGs to continue recording events without administrator intervention.

## 6.4 Protection of Completed Entries

Completed entries cannot be edited through normal Excel operations while sheet protection is active.

This helps prevent accidental changes, overwriting, or deletion.

Authorized administrators can remove sheet protection when corrections are required.

**Important:** Excel worksheet protection is intended to discourage unauthorized or accidental modification. It is not encryption or a tamper-proof security mechanism.

---

# 7. AUTOMATIC ROW MANAGEMENT

ChronLog automatically manages the visibility of log-entry rows.

## 7.1 Unused Rows

Rows not currently needed for data entry remain hidden.

This reduces unnecessary scrolling and maintains a clean logging interface.

## 7.2 Revealing New Rows

When both columns B and C contain information, the next log-entry row becomes visible.

## 7.3 Preserving Existing Entries

Previously completed rows remain visible.

Existing entries are not removed when new rows are revealed.

## 7.4 Maximum Logging Range

The configured log-entry range extends through row 10000.

If this limit is reached, an administrator must evaluate whether the workbook requires expansion.

---

# 8. AUTOMATIC NAVIGATION

ChronLog includes VBA navigation designed to reduce repetitive mouse and keyboard operations.

The standard sequence is:

**Column B → Column C → Next Available Column B**

This process repeats as entries are completed.

The system searches for the next available unlocked entry cell.

When the workbook opens with macros enabled, ChronLog attempts to resume at the first available unlocked, empty entry cell.

This reduces the need to manually locate the next logging position.

---

# 9. WORKSHEET DISPLAY AND FORMATTING

ChronLog uses a restricted layout to maintain consistency.

## 9.1 Visible Columns

Columns A, B, and C contain the operational logging interface.

Columns D through XFD are hidden.

## 9.2 Column Width Protection

Column widths are protected against ordinary user modification.

CAGs cannot resize protected columns without removing sheet protection.

## 9.3 Row Formatting

Protected formatting prevents ordinary users from changing row dimensions.

## 9.4 Horizontal Scrolling

The workbook layout was adjusted to reduce unwanted horizontal scrolling.

Freeze Panes was disabled after it caused column A to disappear from view during navigation.

The display should be configured so that columns A, B, and C fit within the Excel window.

Excel's zoom setting may be adjusted when necessary.

---

# 10. MACRO SECURITY AND STARTUP BEHAVIOR

## 10.1 Opening Without Macros

When the workbook is saved in its intended startup configuration and opened without macros:

- The Enable Macros sheet is displayed.
- Sheet1 remains hidden.
- Normal log entry is unavailable.

This prevents ordinary users from accidentally entering information without timestamps or automatic locking.

## 10.2 Enabling Macros

After macros are enabled:

1. Sheet1 becomes visible.
2. The Enable Macros sheet is hidden.
3. The startup routine prepares the log.
4. Previously recorded entries remain available.
5. The next available entry cell is selected.

## 10.3 Trusted Locations

Excel may automatically permit macros in workbooks opened from previously trusted locations.

Trusted Locations are configured separately on each computer.

A Trusted Location on the administrator's computer does not automatically transfer to another computer.

Only approved folders should be designated as Trusted Locations.

**Current limitation:** Desktop-only execution restrictions were discussed but have not been implemented in v2.0.

## 10.4 Macro Security Limitations

The startup screen helps prevent accidental use without macros.

It does not provide tamper-proof protection against deliberate manipulation of the Excel file.

Authorized administrators should maintain control of the original template and its protection password.

---

# 11. SAVING AND DATA RETENTION

## 11.1 Normal Saving

CAGs should save their daily log regularly using:

**Ctrl + S**

The VBA saving procedure temporarily prepares the workbook so that the Enable Macros sheet is visible and Sheet1 is hidden in the saved file.

The working interface is then restored for continued use.

A brief screen flash may occur during saving.

## 11.2 Save As

Save As is enabled in v2.0.

This allows users to create a new daily log without overwriting the original template.

The new file must remain in `.xlsm` format.

## 11.3 Reopening Existing Logs

When macros are enabled, the workbook restores the ChronLog interface and existing entries.

Previously locked entries remain protected.

The startup routine selects the next available entry cell when one exists.

## 11.4 Daily Log Retention

The intended operating procedure is to create a separate file for each logging date.

The SSM may copy completed daily files to an authorized USB flash drive for review on another computer.

Daily log retention periods, approved storage media, and archival procedures should follow site security requirements.

---

# 12. ADMINISTRATOR ACCESS AND PASSWORD MANAGEMENT

The SSM or another authorized administrator may remove worksheet protection when maintenance or corrections are necessary.

The worksheet protection password should be restricted to authorized personnel.

The password should not be distributed to CAGs who do not require administrative access.

The administrator must distinguish between:

- Excel macro security.
- Sheet protection.
- File access permissions.
- Authorized administrative corrections.

These are separate controls.

## 12.1 Unprotecting Sheet1

1. Open ChronLog with macros enabled.
2. Select Sheet1.
3. Open the Review tab.
4. Select Unprotect Sheet.
5. Enter the administrator password.
6. Click OK.

The sheet can now be edited according to the administrator's permissions.

## 12.2 Editing a Protected Entry

1. Unprotect Sheet1.
2. Locate the entry requiring correction.
3. Make the authorized correction.
4. Verify the revised information.
5. Reapply sheet protection.
6. Save the workbook.

Administrators should document significant corrections according to site procedures.

**Important:** When macros are enabled, the existing VBA may still respond to changes in unlocked entry cells. Administrator corrections should be made carefully and verified afterward.

## 12.3 Restoring Protection

1. Open Review.
2. Select Protect Sheet.
3. Enter the administrator password.
4. Confirm the password if requested.
5. Ensure Select unlocked cells is permitted.
6. Keep Format columns and Format rows disabled.
7. Click OK.

Verify that completed entries cannot be modified through normal use.

---

# 13. RESETTING A LOG TEMPLATE

Only an authorized administrator should reset a workbook for reuse as a blank template.

Before resetting, verify that the file is the intended template and not an operational record requiring retention.

## 13.1 Remove Existing Entries

1. Open the workbook with macros enabled.
2. Select Sheet1.
3. Unprotect the sheet using the administrator password.
4. Use Excel's Name Box to select:

   `A3:C10000`

5. Press Delete to clear the selected cell contents.

This removes existing timestamps, descriptions, and initials without deleting the rows or their formatting.

## 13.2 Restore Cell Locking

After clearing the entries:

1. Select `B3:C10000`.
2. Press Ctrl + 1.
3. Open the Protection tab.
4. Uncheck Locked.
5. Click OK.

Then:

1. Select `A3:A10000`.
2. Press Ctrl + 1.
3. Open the Protection tab.
4. Check Locked.
5. Click OK.

This is essential.

**Column A must remain locked. Columns B and C must be available for new entries.**

## 13.3 Restore Worksheet Protection

Reapply protection using the administrator password.

Confirm that:

- Column A cannot be edited.
- Empty cells in B and C can accept new information.
- Column widths remain protected.
- Row formatting remains protected.

## 13.4 Restore Row Visibility

Save, close, and reopen the workbook with macros enabled.

The startup routine should restore the intended row visibility and select the first available entry cell.

## 13.5 Validate the Blank Template

Before releasing a reset template:

1. Confirm that no previous log entries remain.
2. Confirm that A3 is locked.
3. Confirm that B3 and C3 are unlocked.
4. Enter a temporary test entry in a disposable copy.
5. Verify automatic timestamping and navigation.
6. Confirm that completed cells lock.
7. Verify that the next row becomes available.

Do not retain test entries in the blank master template.

---

# 14. SSM REVIEW PROCEDURES

The SSM is responsible for reviewing daily log entries and identifying omissions, inconsistencies, or entries requiring correction.

## 14.1 Obtaining Daily Logs

The intended procedure is:

1. Locate the completed daily `.xlsm` file on the CAG computer's Desktop.
2. Copy it to an authorized USB flash drive.
3. Transfer it to the designated review computer.
4. Open the file in Microsoft Excel.
5. Review the recorded entries.

## 14.2 Macro Requirements for Review

In the current v2.0 configuration, Sheet1 is normally hidden in the saved file.

Therefore, the SSM may need to enable approved macros to display the log.

The workbook does not currently provide a separate macro-free review mode.

## 14.3 Correcting Entries

If a correction is required:

1. Open the daily log.
2. Unprotect Sheet1 using the administrator password.
3. Make the authorized correction.
4. Reapply sheet protection.
5. Save the file.

The SSM should maintain an appropriate record of substantive corrections.

No separate review workbook is generated by ChronLog v2.0.

---

# 15. TROUBLESHOOTING

## 15.1 Automatic Timestamps Are Missing

Possible causes:

- Macros are disabled.
- The workbook is not operating in its intended startup state.
- The timestamp cell already contains information.
- VBA execution was interrupted.

Recommended action:

1. Confirm that macros are enabled.
2. Verify that column A is locked.
3. Check the affected row.
4. Notify the administrator if the issue continues.

Do not manually overwrite timestamps to conceal an error.

## 15.2 Excel Does Not Move to the Next Cell

Verify that macros are enabled.

Normal Excel navigation does not reproduce the complete ChronLog workflow.

If automatic navigation fails, save the workbook and notify the administrator.

## 15.3 New Rows Do Not Appear

Verify that both columns B and C contain information.

The next row is revealed after the current row is complete.

## 15.4 A Completed Cell Cannot Be Edited

This is expected behavior.

Completed cells are automatically locked.

Only authorized administrators should remove protection to make corrections.

## 15.5 Column A Disappears

Check the workbook's horizontal scroll position.

Ensure Freeze Panes is disabled.

Adjust the Excel zoom level or display layout so columns A, B, and C fit on the screen.

## 15.6 The Enable Macros Screen Appears

This indicates that the workbook is in its saved startup configuration.

Enable approved macros to access Sheet1.

If macros cannot be enabled, contact the administrator.

## 15.7 Save As Does Not Work

Verify that the workbook is running the current v2.0 VBA configuration.

Confirm that the selected destination is writable.

Save the new file using the `.xlsm` format.

## 15.8 A File Opens Without Expected Entries

Verify that the correct daily file was opened.

Check the filename and date.

Do not overwrite the file until the issue is investigated.

## 15.9 Excel Reports a VBA Error

Record the exact error message.

Avoid modifying VBA without authorization.

Contact the administrator responsible for maintaining ChronLog.

---

# 16. ADMINISTRATOR MAINTENANCE

Administrators should periodically verify:

- The original template remains blank.
- Daily Save As functionality works.
- Macros execute correctly.
- Timestamps are generated.
- Column A remains locked.
- Columns B and C lock after entry.
- Empty entry cells remain unlocked.
- New rows appear automatically.
- Existing entries remain intact.
- Worksheet protection remains active.
- The startup instruction screen operates correctly.
- Daily files are stored and retained according to site requirements.

Administrators should maintain a backup of the current working template before modifying VBA.

Any future modification should be tested using a separate copy before deployment.

---

# 17. SECURITY CONSIDERATIONS

ChronLog v2.0 includes multiple safeguards intended to reduce accidental alteration of security records.

These include:

- Automatic timestamps.
- Protected timestamp cells.
- Automatic locking of populated entries.
- Worksheet protection.
- Restricted formatting.
- Hidden unused columns and rows.
- Macro-required startup behavior.
- Controlled administrator corrections.

These measures establish a consistent workflow and make ordinary accidental changes less likely.

However, ChronLog does not provide cryptographic integrity verification, digital signatures, independent audit trails of every edit, or tamper-proof storage.

Worksheet protection alone cannot establish who made an unauthorized modification.

The SSM should maintain appropriate custody and retention procedures for completed records.

---

# 18. VERSION 2.0 CHANGELOG

**Release Date:** November 19, 2026

## Logging Enhancements

- Improved automatic timestamp handling.
- Added permanent protection to the timestamp column.
- Improved navigation between log descriptions and initials.
- Added automatic navigation to the next available log row.
- Preserved existing timestamps when additional row information is entered.

## Data Protection Enhancements

- Automatic locking of populated entry cells.
- Preservation of unlocked empty entry cells.
- Verification of protection after saving and reopening.
- Protection of column widths and row formatting.

## Interface Enhancements

- Hidden unused columns.
- Dynamic display of required log rows.
- Preservation of previously completed rows.
- Improved startup positioning.
- Corrected horizontal scrolling behavior.
- Removed problematic Freeze Panes configuration.

## Macro and Startup Enhancements

- Added the Enable Macros instruction sheet.
- Added automatic switching to Sheet1 when macros are enabled.
- Added startup preparation for existing logs.
- Added saved-file startup configuration to discourage entry without macros.

## File Management Enhancements

- Updated VBA save handling.
- Restored Save As functionality.
- Enabled creation of separate date-named daily logs.
- Preserved the original template during daily log creation.
- Verified retention of existing entries and protection after reopening.

## Testing Completed

- Timestamp generation.
- Automatic navigation.
- Cell locking.
- Worksheet protection.
- Dynamic row visibility.
- Macro startup behavior.
- Normal saving.
- Save As.
- Reopening existing logs.
- Preservation of recorded information.

---

# 19. KNOWN LIMITATIONS AND FUTURE IMPROVEMENTS

The following features are not currently implemented:

- Desktop-only macro execution restrictions.
- Cryptographic verification of log integrity.
- Independent logging of administrator corrections.
- Automatic finalization of completed daily logs.
- A separate macro-free SSM review mode.

These features may be evaluated in future releases.

They are not required for the currently established daily logging workflow.

---

# 20. ADMINISTRATOR QUICK REFERENCE

| Task | Procedure |
|---|---|
| Start a daily log | Open blank template, enable macros, Save As |
| Save current log | Ctrl + S |
| Enter log description | Column B |
| Enter CAG name/initials | Column C |
| View timestamp | Column A |
| Correct a locked entry | Review → Unprotect Sheet |
| Restore protection | Review → Protect Sheet |
| Clear template entries | Select A3:C10000, press Delete |
| Unlock blank entry cells | Format Cells → Protection → Uncheck Locked for B3:C10000 |
| Lock timestamps | Format Cells → Protection → Check Locked for A3:A10000 |
| Review a daily file | Open original daily `.xlsm` file |
| Create a new daily file | File → Save As → Desktop |
| Preserve original template | Never use the blank master for daily entries |

---

# 21. FINAL ADMINISTRATOR CHECKLIST

Before releasing ChronLog v2.0 for operational use:

- [ ] Original template is blank.
- [ ] Worksheet protection is enabled.
- [ ] A3:A10000 is locked.
- [ ] Empty B3:C10000 cells are unlocked.
- [ ] Automatic timestamps work.
- [ ] Automatic B-to-C navigation works.
- [ ] Automatic next-row navigation works.
- [ ] Completed entries lock automatically.
- [ ] Unused rows are hidden appropriately.
- [ ] Previous entries remain visible.
- [ ] Enable Macros startup screen functions.
- [ ] Save As creates a separate daily file.
- [ ] Original template remains unchanged.
- [ ] Saved daily files reopen correctly.
- [ ] Administrator password is controlled.
- [ ] Approved daily log storage and retention procedures are established.

---

**CAG CHRONLOG v2.0**

**Administrator User Manual**

**Release Date: November 19, 2026**

**End of Document**

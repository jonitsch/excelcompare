# excelcompare

A PowerShell utility for comparing two Excel worksheets, grouping identical differences, and optionally merging changes between files.

## Features

* Compare two Excel workbooks (`.xlsx`, `.xlsm`, etc.)
* Select worksheets interactively
* Detect all cell differences
* Group identical differences together
* Generate a unique difference ID for each difference group
* View all affected rows for a specific difference
* Merge differences from one workbook into the other
* Save merged changes directly to the target workbook

## Requirements

* Windows
* Microsoft Excel installed
* PowerShell 5.1 or later

## Usage

Run the script:

```powershell
.\ExcelCompare.ps1
```

### Step 1: Select Files

Enter the full path to both Excel workbooks:

```text
Enter path to first Excel file
Enter path to second Excel file
```

### Step 2: Select Worksheets

The script displays all available worksheet names and prompts you to select one worksheet from each workbook.

If both workbooks use the same worksheet name, simply press Enter when prompted for the second sheet.

### Step 3: Compare Data

The script reads both worksheets and compares all used rows and columns.

Example:

```text
Comparing 1200 rows and 15 columns...
```

Differences are grouped by:

* Column
* Value in File 1
* Value in File 2

This prevents hundreds of identical changes from being displayed individually.

## Difference Overview

Example output:

```text
ID           Count Column      File1_Value    File2_Value      Rows
--           ----- ------      -----------    -----------      ----
R2wL1nA9T3Y=   120 Department  Engineering    ENG              {12,13,14,15...}
cY6mPq8XaEk=    55 Status      Active         Enabled          {42,43,44,45...}
8Xb4JtVnR0Q=    12 Priority    High           Critical         {88,89,90...}
```

### Columns

| Column      | Description                                 |
| ----------- | ------------------------------------------- |
| ID          | Unique identifier for this difference group |
| Count       | Number of affected rows                     |
| Column      | Worksheet column containing the difference  |
| File1_Value | Value found in the first workbook           |
| File2_Value | Value found in the second workbook          |
| Rows        | Affected row numbers                        |

## Inspecting a Difference

Choose **List all differences by difference ID** and enter an ID:

```text
Enter an id:
R2wL1nA9T3Y=
```

Example output:

```text
Column: Department

Row  File1_Value  File2_Value
---  -----------  -----------
12   Engineering  ENG
13   Engineering  ENG
14   Engineering  ENG
15   Engineering  ENG
```

## Merging Differences

Choose **Merge differences by difference ID** and enter an ID.

The script then allows you to:

```text
1   Accept File1 Value 'Engineering'
2   Accept File2 Value 'ENG'
```

Selecting an option updates all affected rows and saves the target workbook automatically.

After saving, the worksheets are re-evaluated so the difference overview remains current.

## Notes

* Only cells with differing values are evaluated.
* Difference IDs are generated from the column and differing values and remain consistent during a comparison session.
* Workbooks are modified directly when merging.
* Always keep a backup of important files before performing merges.

## Cleanup

When exiting, the script:

* Closes all opened workbooks
* Terminates the Excel COM session
* Releases COM resources
* Performs garbage collection

This helps prevent orphaned Excel processes from remaining in memory.

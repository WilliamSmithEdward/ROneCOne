# Getting started

You can prove ROneCOne in a disposable demo before touching your own workbook. When you are ready,
installation is one class import.

## Option 1: see it working first

1. Open the [latest release](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest).
2. Download the demo workbook for the capability you want to explore.
3. Save it to a trusted local folder and open it in Excel.
4. Enable macros only after confirming that the file came from the ROneCOne release.
5. Press `Alt+F8`, select the demo macro, and choose **Run**.

| Workbook | Macro |
|---|---|
| Collections and LINQ | `RunROneCOneCollectionsDemo` |
| Delegates and expressions | `RunROneCOneDemo` |
| Typed events | `RunROneCOneEventsDemo` |
| Structured exceptions | `RunROneCOneExceptionsDemo` |
| Tasks and async | `RunROneCOneTasksDemo` |
| Data and providers | `RunROneCOneDataDemo` |
| HTTP and web data | `RunROneCOneHttpDemo` |
| JSON | `RunROneCOneJsonDemo` |
| Files and CSV | `RunROneCOneFilesDemo` |
| Processes | `RunROneCOneProcessDemo` |
| Text and hashing | `RunROneCOneTextDemo` |
| Dates and times | `RunROneCOneDateTimeDemo` |
| XML | `RunROneCOneXmlDemo` |

Each workbook writes its results into visible worksheets. Start with Collections and LINQ if you
want the broadest tour of everyday features.

## Option 2: add ROneCOne to your workbook

> [!TIP]
> Work on a copy of the workbook until you are comfortable with the result.

1. Download `ROneCOne.cls` from the
   [latest release](https://github.com/WilliamSmithEdward/ROneCOne/releases/latest).
2. Save your workbook as `.xlsm`, `.xlsb`, or `.xlam`.
3. Press `Alt+F11` to open the Visual Basic Editor.
4. Choose **File > Import File**.
5. Select `ROneCOne.cls`.
6. Confirm that **ROneCOne** appears under **Class Modules**.
7. Save the workbook.

No reference needs to be added under **Tools > References**.

## Your first check

Create a standard module, paste this procedure, and run it with `F5`:

```vba
Option Explicit

Public Sub TryROneCOne()
    Dim strongScores As ROneCOne
    Dim scores As ROneCOne

    Set scores = ROneCOne.ListOf(vbLong, 90, 72, 88, 95)

    Set strongScores = scores _
        .Where(scores.Element.AtLeast(85)) _
        .OrderDescending _
        .ToList

    MsgBox "Scores at or above 85: " & strongScores.JoinText(", ")
End Sub
```

The message should display `Scores at or above 85: 95, 90, 88`. `Element` means "the current
number" while the list is being filtered. The same pattern works with object properties such as
`Where("Age").AtLeast(40)`.

You wrote `90`, not `CLng(90)`. A typed list accepts any number that fits its type without loss
and stores it as the declared type, so plain integer literals go straight into a `List<Long>`.
A value that would lose information, such as a decimal into a whole-number list, is still
refused.

You have created a checked list, filtered it, sorted it, and displayed the result without writing a
loop, counter, or temporary array.

## If Excel blocks the file

Files downloaded from the internet can carry Windows' Mark of the Web. If Excel blocks macros:

1. Close the workbook.
2. In File Explorer, right-click the file and choose **Properties**.
3. If an **Unblock** checkbox appears, review the file source, select it, and choose **OK**.
4. Reopen the workbook.

> [!WARNING]
> Do not weaken Excel's global macro settings. Trust only the specific file you reviewed.

## Upgrading from an earlier release

Remove the old class before importing the new one:

1. Press `Alt+F11` and find **ROneCOne** under **Class Modules**.
2. Right-click it, choose **Remove ROneCOne**, and answer **No** when asked to export it first.
3. Choose **File > Import File** and select the new `ROneCOne.cls`.
4. Save the workbook.

Releases up to 1.9.0 changed the spelling of names in your own modules, turning `.Value` into
`.value` ([issue #6](https://github.com/WilliamSmithEdward/ROneCOne/issues/6)). VBA keeps one
spelling per name across a project, and the change stays after the old class is removed.
Importing 1.9.1 or later restores the names the runtime declares itself, such as `Value`,
`Count`, and `Item`. Names that an older release declared and 1.9.1 does not, 72 in all, can keep
the old spelling wherever your code uses them: `.cells`, `.names`, and `.width`, for example.

To restore them, paste this procedure at the end of any module, then delete it:

```vba
Private Sub RestoreSpellings()
    Dim Address, Alignment, Amount, ApplicationName, Cells, Character, Charset, Child, ChildNodes
    Dim Clone, ColumnIndex, Comparison, Context, Cursor, DateValue, Depth, Description
    Dim DocumentElement, ErrorText, Errors, Expression, Field, Fields, FileNumber, Filename, Group
    Dim HResult, InsertAfter, LeftColumn, Level, Levels, Limit, Line, LoadXML, Lookup, Markers, Mask
    Dim Method, Names, Negative, NodeType, Operation, Part, Position, Proper, PropertyName
    Dim Protection, Query, Reason, Record, Recordset, Reference, Resolved, Results, ReturnType, Root
    Dim RowCount, RowIndex, RowOffset, Scope, Selected, Sequence, SetProperty, Sheet, Sql, Stream
    Dim Symbol, Total, UniqueValues, Version, View, Width
End Sub
```

Declaring a name sets its spelling everywhere in the project, and the spelling stays after the
procedure is deleted, through saving and reopening. Nothing runs. A variable of your own with one
of these names, such as `total`, takes the listed spelling as well.

## Where next

- [Collections and LINQ](collections-and-linq.md) continues with everyday queries.
- [Guide index](README.md) shows the full learning path.

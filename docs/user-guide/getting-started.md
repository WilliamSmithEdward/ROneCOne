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

    Set scores = ROneCOne.ListOf(vbLong, _
        CLng(90), CLng(72), CLng(88), CLng(95))

    Set strongScores = scores _
        .Where(scores.Element.AtLeast(CLng(85))) _
        .OrderDescending _
        .ToList

    MsgBox "Scores at or above 85: " & strongScores.JoinText(", ")
End Sub
```

The message should display `Scores at or above 85: 95, 90, 88`. `Element` means "the current
number" while the list is being filtered. The same pattern works with object properties such as
`Where("Age").AtLeast(40)`.

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

## Where next

- [Collections and LINQ](collections-and-linq.md) continues with everyday queries.
- [Guide index](README.md) shows the full learning path.

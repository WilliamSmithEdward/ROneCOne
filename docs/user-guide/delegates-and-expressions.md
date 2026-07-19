# Delegates and expressions

A delegate is a value that represents behavior. It can be stored, passed to another feature,
combined with other behavior, inspected, and invoked later. ROneCOne gives functions and actions
one consistent shape whether they come from an expression, an object, or a workbook procedure.

## Build a function without another procedure

This creates a typed function that squares a `Long`:

```vba
Dim square As ROneCOne
Dim x As ROneCOne

Set x = ROneCOne.Var(vbLong)
Set square = x.Multiply(x).AsFunc

Debug.Print square(CLng(9))
```

The result is `81`. The function can be passed anywhere ROneCOne expects a compatible `Func`.

## Wrap an existing object method

Object methods can use the same contract:

```vba
Dim maximum As ROneCOne

Set maximum = ROneCOne.Func( _
    Application.WorksheetFunction, "Max") _
    .Takes(vbLong, vbLong) _
    .Returns(vbDouble)

Debug.Print maximum(CLng(4), CLng(7))
```

`Takes` and `Returns` make the expected signature explicit and catch mismatches before the target
is invoked.

## Wrap a workbook procedure

Given a public procedure named `Transform` in a standard module named `TextTools`:

```vba
Dim transform As ROneCOne

Set transform = ROneCOne.Func("TextTools.Transform") _
    .Takes(vbString) _
    .Returns(vbString)

Debug.Print transform("ready")
```

Workbook procedures need a qualified name because VBA does not expose a general first-class
procedure reference. Expression-based functions remain string-free.

## Use actions for work with no return value

```vba
Dim writeAudit As ROneCOne

Set writeAudit = ROneCOne.Action("Audit.WriteEntry") _
    .Takes(vbString)

writeAudit.Execute "Workbook opened"
```

`Execute` is the statement form for an Action. It validates the same signature as a function call
and does not require a dummy result variable.

## Combine actions

Compatible delegates form an immutable multicast chain:

```vba
Dim notify As ROneCOne

Set notify = ROneCOne.Combine(writeAudit, updateStatus)
notify.Execute "ready"

Set notify = notify.Remove(updateStatus)
```

Actions run in insertion order. Removing a handler returns a new delegate and leaves the original
one unchanged.

## Compose functions

`PipeTo` sends one function's result into the next one:

```vba
Dim normalizeAndMeasure As ROneCOne

Set normalizeAndMeasure = normalizeText.PipeTo(textLength)
Debug.Print normalizeAndMeasure("  Excel  ")
```

The two delegates must have compatible input and output contracts.

## Inspect a delegate

```vba
Debug.Print maximum.Signature
Debug.Print maximum.Arity
Debug.Print maximum.MethodName
Debug.Print maximum.InvocationCount
```

Metadata makes dynamically bound behavior easier to diagnose without changing the target code.

## Native and ByRef work

ROneCOne also supports signature-bound Windows x64 native calls and true native `ByRef` variables.
That surface is intentionally strict because an incorrect native signature can destabilize Excel.
Use it only after reading the
[native invocation contract](../delegates.md#native-invocation-and-true-byref).

## Where to go deeper

- [Delegates technical reference](../delegates.md) covers every target, signature rule, multicast
  behavior, dynamic invocation, native safety boundary, and canonical form.
- [Events and exceptions](events-and-exceptions.md) shows delegates coordinating real workbook
  flows.
- [Guide index](README.md) returns to the full learning path.

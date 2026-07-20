# Delegates and expressions

A delegate is a reusable piece of work. Instead of running a procedure immediately, you can save
it in a variable, pass it to another feature, combine it with other work, and run it later.
ROneCOne calls work that returns an answer a `Func`; work with no answer is an `Action`.

## Build a pricing rule without another procedure

This rule applies a ten percent discount to an order amount:

```vba
Dim applyDiscount As ROneCOne
Dim price As ROneCOne

Set price = ROneCOne.Var(vbDouble)
Set applyDiscount = price.Multiply(0.9).AsFunc

Debug.Print applyDiscount(CDbl(100))
```

The result is `90`. `price` is the input placeholder; `AsFunc` turns the expression into reusable
work. The rule can now be called directly or passed to collections, Tasks, and other features.

## Wrap an existing object method

Existing Excel methods can use the same checked shape. This example chooses the greater of two
stock targets:

```vba
Dim chooseHigherStock As ROneCOne

Set chooseHigherStock = ROneCOne.Func( _
    Application.WorksheetFunction, "Max") _
    .Takes(vbLong, vbLong) _
    .Returns(vbDouble)

Debug.Print chooseHigherStock(CLng(40), CLng(75))
```

`Takes` and `Returns` make the expected signature explicit and catch mismatches before the target
is invoked.

## Wrap a workbook procedure

Given a public function named `CalculateOrderTotal` in a standard module named `Orders`:

```vba
Dim calculateTotal As ROneCOne

Set calculateTotal = ROneCOne.Func("Orders.CalculateOrderTotal") _
    .Takes(vbLong, vbLong) _
    .Returns(vbLong)

Debug.Print calculateTotal(CLng(100), CLng(5))
```

Workbook procedures need a qualified name because VBA does not expose a general first-class
procedure reference. Expression-based functions remain string-free.

## Use actions for work with no return value

```vba
Dim writeAudit As ROneCOne

Set writeAudit = ROneCOne.Action("Audit.WriteEntry") _
    .Takes(vbString)

writeAudit.Execute "Order 1042 approved"
```

`Execute` is the statement form for an Action. It validates the same signature as a function call
and does not require a dummy result variable.

## Combine actions

Compatible delegates form an immutable multicast chain:

```vba
Dim notify As ROneCOne

Set notify = ROneCOne.Combine(writeAudit, updateStatus)
notify.Execute "Order 1042 approved"

Set notify = notify.Remove(updateStatus)
```

Actions run in insertion order. Removing a handler returns a new delegate and leaves the original
one unchanged.

## Compose functions

`PipeTo` sends one function's answer into the next function. Here the first rule applies a
discount and the second adds a handling charge:

```vba
Dim finalPrice As ROneCOne

Set finalPrice = applyDiscount.PipeTo(addHandling)
Debug.Print finalPrice(CDbl(100))
```

The two delegates must have compatible input and output contracts.

## Inspect a delegate

```vba
Debug.Print calculateTotal.Signature
Debug.Print calculateTotal.Arity
Debug.Print calculateTotal.MethodName
Debug.Print calculateTotal.InvocationCount
```

Metadata makes dynamically bound behavior easier to diagnose without changing the target code.

## Native and ByRef work

ROneCOne also supports signature-bound Windows x64 native calls and true native `ByRef` variables.

> [!WARNING]
> The native surface is intentionally strict because an incorrect native signature can
> destabilize Excel. Use it only after reading the
> [native invocation contract](../delegates.md#native-invocation-and-true-byref).

## Where next

- [Events and exceptions](events-and-exceptions.md) shows delegates coordinating real workbook
  flows.
- [Delegates technical reference](../delegates.md) covers every target, signature rule, multicast
  behavior, dynamic invocation, native safety boundary, and canonical form.
- [Guide index](README.md) returns to the full learning path.

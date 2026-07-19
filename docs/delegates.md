# Universal delegates

New to ROneCOne? Start with the
[Delegates and expressions user guide](user-guide/delegates-and-expressions.md).

ROneCOne represents expression lambdas, object methods, callable objects, workbook procedures,
multicast chains, compositions, and signature-bound native entry points with one immutable type.
The public model follows C# `Func`, `Action`, `Delegate`, and `DynamicInvoke` semantics where the
VBA language and COM host permit them.

## Expression lambdas

The shortest string-free form builds an immutable expression and infers its typed parameters:

```vba
Dim square As ROneCOne
Dim x As ROneCOne

Set x = ROneCOne.Var(vbLong)
Set square = x.Multiply(x).AsFunc
Debug.Print square(CLng(9))
```

The canonical expansion is:

```vba
Set x = ROneCOne.Parameter(vbLong)
Set square = ROneCOne.Lambda(x.Multiply(x), x)
```

`AsFunc` discovers unique parameters once in left-to-right order. `ROneCOne.Value(capturedValue)`
adds an immutable captured scalar or object reference to an expression tree.

## Universal factories

`Func` and `Action` normalize supported call targets behind the same invocation contract.

| Target | Construction | Dispatch boundary |
|---|---|---|
| Expression | `ROneCOne.Func(expression)` | ROneCOne expression evaluator |
| Object method | `ROneCOne.Func(target, "Transform")` | `CallByName` |
| Callable object | `ROneCOne.Func(target)` | `target.Run` |
| Workbook procedure | `ROneCOne.Func("Module.Transform")` | `Application.Run` |
| Native address | `ROneCOne.Native(address)` | `DispCallFunc` |

Object members and standard-module procedures require a name because VBA exposes neither a general
first-class procedure value nor `nameof`. Expression lambdas remain string-free. Omitting the
object member name binds `Run`, giving application classes a concise callable-object convention.

`Takes` and `Returns` add immutable runtime signature metadata:

```vba
Set transform = ROneCOne.Func(target, "Transform") _
    .Takes(vbLong) _
    .Returns(vbString)
```

Use a non-`Nothing` class prototype instead of a `VbVarType` token for an exact user-defined class:

```vba
Set transform = transform.Takes(customerPrototype).Returns(resultPrototype)
```

Signatures enforce arity, exact primitive `VarType`, exact class `TypeName`, object/scalar result
shape, and return type. Dynamic object and workbook delegates may omit metadata; native delegates
must provide a complete signature and fail before dispatch if it is missing.

## Invocation and metadata

The default member provides the natural C#-like call form. `Run` is the explicit equivalent, and
`DynamicInvoke` consumes a VBA array, `Collection`, or ROneCOne sequence.

```vba
result = transform(CLng(7))
result = transform.Run(CLng(7))
result = transform.DynamicInvoke(Array(CLng(7)))
```

VBA class modules inherit an `Invoke` dispatch member and reject a public method with that name at
compile time. `Run` is therefore the nearest valid explicit name; default invocation removes the
name entirely.

Every delegate exposes `Arity`, `Target`, `MethodName`, `IsAction`, `Signature`,
`InvocationCount`, and `GetInvocationList`. A signature is rendered in a C#-aligned form such as
`Func<Long, Long, Long>` or `Action<String>`.

## Multicast and composition

`Combine` flattens compatible delegates into an immutable invocation list. Actions run in insertion
order. A multicast function returns the last invocation's result, matching .NET multicast delegate
behavior. `Remove` removes the last matching invocation subsequence and leaves the source unchanged.

`ROneCOne.Action(existingFunc)` deliberately discards a result. The reverse conversion is invalid:
`ROneCOne.Func(existingAction)` and `existingAction.Returns(...)` raise `InvalidOperationError`.

```vba
Set changed = ROneCOne.Combine(firstHandler, secondHandler)
changed.Execute "ready"
Set withoutSecond = changed.Remove(secondHandler)
Set handlers = changed.GetInvocationList
```

`Execute` is the statement-form invocation surface for Actions. It validates the same signature as
`Run` and discards the Empty Action result, so application code never needs an `ignored` variable.

`PipeTo` is function composition: it passes the first delegate's result to a one-argument
continuation.

## Native invocation and true ByRef

Windows x64 native delegates use `DispCallFunc` from `oleaut32.dll`. `Native` and `NativeAction`
reject a zero address, require complete `Takes` / `Returns` metadata, and never guess a signature.
That fail-closed rule protects the Excel process from an ABI mismatch.

VBA's `Application.Run` and `CallByName` boundaries do not preserve a variable's `ByRef` identity.
ROneCOne therefore admits true `ByRef` only for signature-bound native delegates. `RefOf` marks it;
a typed reference factory captures the caller's original variable:

```vba
Dim increment As ROneCOne
Dim value As Long

value = 41
Set increment = ROneCOne.NativeAction(IncrementLongAddress) _
    .Takes(ROneCOne.RefOf(vbLong))
increment.Execute ROneCOne.RefLong(value)
Debug.Print value  ' 42
```

Typed variable wrappers are `RefByte`, `RefInteger`, `RefLong`, `RefLongLong`, `RefSingle`,
`RefDouble`, `RefCurrency`, `RefDate`, and `RefBoolean`. Passing a reference descriptor to an object
or workbook-procedure delegate raises `ByRefInvocationError` before the target runs.

A standard VBA module can expose an `AddressOf` value through an application-local pointer bridge;
the delegate demo contains a complete example. That bridge is application code, not another runtime
module. The shipped runtime remains the single `ROneCOne.cls` file and does not use VBIDE access or
generate trampolines.

## Limits that fail closed

- Object methods currently accept up to eight arguments through `CallByName`.
- Workbook procedures accept up to Excel's documented 30 `Application.Run` arguments.
- Native signatures are Windows x64 only and must match the actual procedure ABI exactly.
- True `ByRef` is native-only because late-bound VBA dispatch loses variable identity.
- `Invoke` cannot be declared by a VBA class; use natural call syntax, `Run`, or `DynamicInvoke`.

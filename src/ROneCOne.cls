VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ROneCOne"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Private Const ROLE_EMPTY As Long = 0
Private Const ROLE_VALUE As Long = 1
Private Const ROLE_PARAMETER As Long = 2
Private Const ROLE_EXPRESSION As Long = 3
Private Const ROLE_LAMBDA As Long = 4
Private Const ROLE_METHOD As Long = 5
Private Const ROLE_COMPOSITION As Long = 6
Private Const ROLE_LIST As Long = 7
Private Const ROLE_QUERY As Long = 8

Private Const OP_ADD As Long = 1
Private Const OP_SUBTRACT As Long = 2
Private Const OP_MULTIPLY As Long = 3
Private Const OP_DIVIDE As Long = 4
Private Const OP_MODULO As Long = 5
Private Const OP_NEGATE As Long = 6
Private Const OP_CONCAT As Long = 7
Private Const OP_EQUAL As Long = 8
Private Const OP_NOT_EQUAL As Long = 9
Private Const OP_LESS_THAN As Long = 10
Private Const OP_LESS_THAN_OR_EQUAL As Long = 11
Private Const OP_GREATER_THAN As Long = 12
Private Const OP_GREATER_THAN_OR_EQUAL As Long = 13
Private Const OP_AND_ALSO As Long = 14
Private Const OP_OR_ELSE As Long = 15
Private Const OP_NOT As Long = 16

Private Const QUERY_WHERE As Long = 1
Private Const QUERY_SELECT As Long = 2
Private Const QUERY_TAKE As Long = 3
Private Const QUERY_SKIP As Long = 4
Private Const QUERY_DISTINCT As Long = 5
Private Const QUERY_ORDER_BY As Long = 6
Private Const QUERY_ORDER_BY_DESCENDING As Long = 7
Private Const QUERY_APPEND As Long = 8
Private Const QUERY_PREPEND As Long = 9
Private Const QUERY_REVERSE As Long = 10

Private Const ERROR_INVALID_ARGUMENT As Long = vbObjectError + 4100
Private Const ERROR_INVALID_OPERATION As Long = vbObjectError + 4101
Private Const ERROR_ARITY_MISMATCH As Long = vbObjectError + 4102
Private Const ERROR_TYPE_MISMATCH As Long = vbObjectError + 4103
Private Const ERROR_UNBOUND_PARAMETER As Long = vbObjectError + 4104
Private Const ERROR_METHOD_INVOCATION As Long = vbObjectError + 4105

Private mRole As Long
Private mOperator As Long
Private mLeft As ROneCOne
Private mRight As ROneCOne
Private mNext As ROneCOne

Private mScalarValue As Variant
Private mObjectValue As Object
Private mHasObjectValue As Boolean

Private mExpectedType As VbVarType
Private mExpectedObjectType As String
Private mParameters As Collection

Private mTarget As Object
Private mMethodName As String
Private mMethodArity As Long
Private mMethodReturnsObject As Boolean

Private mItems As Collection
Private mEnumerationValues As Collection
Private mElementType As VbVarType
Private mElementObjectType As String
Private mSource As ROneCOne
Private mQueryDelegate As ROneCOne
Private mQueryValue As ROneCOne
Private mQueryCount As Long
Private mVersion As Long
Private mEnumerationVersion As Long

Public Function Parameter( _
    Optional ByVal expectedType As VbVarType = vbVariant _
) As ROneCOne
Attribute Parameter.VB_Description = "Creates a typed lambda parameter."
    Dim node As ROneCOne

    Set node = New ROneCOne
    node.ConfigureParameter expectedType, vbNullString
    Set Parameter = node
End Function

Public Function ParameterLike(ByVal example As Variant) As ROneCOne
Attribute ParameterLike.VB_Description = "Creates a parameter typed like an example value."
    Dim node As ROneCOne
    Dim exampleObject As Object

    Set node = New ROneCOne
    If IsObject(example) Then
        Set exampleObject = example
        If exampleObject Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ParameterLike", _
                "The example object cannot be Nothing."
        End If
        node.ConfigureParameter vbObject, TypeName(exampleObject)
    Else
        node.ConfigureParameter VarType(example), vbNullString
    End If
    Set ParameterLike = node
End Function

Public Function Value(ByVal capturedValue As Variant) As ROneCOne
Attribute Value.VB_Description = "Captures a scalar or object as an expression value."
    Dim node As ROneCOne
    Dim capturedObject As Object

    Set node = New ROneCOne
    If IsObject(capturedValue) Then
        Set capturedObject = capturedValue
        node.ConfigureObjectValue capturedObject
    Else
        node.ConfigureScalarValue capturedValue
    End If
    Set Value = node
End Function

Public Function ListOf(ByVal typeToken As Variant) As ROneCOne
Attribute ListOf.VB_Description = "Creates an empty, strictly typed List<T>."
    Dim listValue As ROneCOne
    Dim prototype As Object

    Set listValue = New ROneCOne
    If IsObject(typeToken) Then
        Set prototype = typeToken
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ListOf", _
                "The user-defined type prototype cannot be Nothing."
        End If
        listValue.ConfigureList vbObject, TypeName(prototype)
    Else
        ValidateTypeToken typeToken, "ListOf"
        listValue.ConfigureList CLng(typeToken), vbNullString
    End If
    Set ListOf = listValue
End Function

Public Function ListLike(ByVal example As Variant) As ROneCOne
Attribute ListLike.VB_Description = "Creates a List<T> typed like an example value."
    Dim listValue As ROneCOne
    Dim prototype As Object

    Set listValue = New ROneCOne
    If IsObject(example) Then
        Set prototype = example
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ListLike", _
                "The user-defined type prototype cannot be Nothing."
        End If
        listValue.ConfigureList vbObject, TypeName(prototype)
    Else
        listValue.ConfigureList VarType(example), vbNullString
    End If
    Set ListLike = listValue
End Function

Public Function Range(ByVal start As Long, ByVal count As Long) As ROneCOne
Attribute Range.VB_Description = "Creates a sequence of consecutive Long values."
    Dim index As Long
    Dim listValue As ROneCOne

    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Range", _
            "Count cannot be negative."
    End If
    If count > 0 Then
        If CDbl(start) + CDbl(count) - 1# > CDbl(2147483647) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Range", _
                "The requested range exceeds the Long data type."
        End If
    End If

    Set listValue = ROneCOne.ListOf(vbLong)
    For index = 0 To count - 1
        listValue.Add CLng(CDbl(start) + CDbl(index))
    Next index
    Set Range = listValue
End Function

Public Function Repeat(ByVal element As Variant, ByVal count As Long) As ROneCOne
Attribute Repeat.VB_Description = "Creates a sequence containing a repeated element."
    Dim index As Long
    Dim listValue As ROneCOne

    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Repeat", _
            "Count cannot be negative."
    End If
    Set listValue = ROneCOne.ListLike(element)
    For index = 1 To count
        listValue.Add element
    Next index
    Set Repeat = listValue
End Function

Public Function Lambda( _
    ByVal body As ROneCOne, _
    ParamArray parameters() As Variant _
) As ROneCOne
Attribute Lambda.VB_Description = "Builds an immutable expression-tree lambda delegate."
    Dim delegateValue As ROneCOne
    Dim parameterList As Collection
    Dim parameterValue As ROneCOne
    Dim rawParameter As Variant

    If body Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
            "The lambda body cannot be Nothing."
    End If

    Set parameterList = New Collection
    For Each rawParameter In parameters
        Set parameterValue = RequireROneCOne( _
            rawParameter, _
            "Lambda parameter")
        If parameterValue.InternalRole <> ROLE_PARAMETER Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "Every lambda parameter must be created by ROneCOne.Parameter."
        End If
        If CollectionContainsIdentity(parameterList, parameterValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "A lambda parameter cannot be declared more than once."
        End If
        parameterList.Add parameterValue
    Next rawParameter

    ValidateBindings body, parameterList

    Set delegateValue = New ROneCOne
    delegateValue.ConfigureLambda body, parameterList
    Set Lambda = delegateValue
End Function

Public Function FromMethod( _
    ByVal target As Object, _
    ByVal methodName As String, _
    Optional ByVal parameterCount As Long = -1, _
    Optional ByVal returnsObject As Boolean = False _
) As ROneCOne
Attribute FromMethod.VB_Description = "Creates a delegate for an object method."
    Dim delegateValue As ROneCOne

    If target Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromMethod", _
            "The method target cannot be Nothing."
    End If
    If Len(Trim$(methodName)) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromMethod", _
            "The method name cannot be empty."
    End If
    If parameterCount < -1 Or parameterCount > 8 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromMethod", _
            "The parameter count must be -1 or between 0 and 8."
    End If

    Set delegateValue = New ROneCOne
    delegateValue.ConfigureMethod target, Trim$(methodName), parameterCount, returnsObject
    Set FromMethod = delegateValue
End Function

Public Function Run(ParamArray arguments() As Variant) As Variant
Attribute Run.VB_Description = "Invokes a delegate or gets a zero-based sequence item."
Attribute Run.VB_UserMemId = 0
    Dim argumentValues As Collection
    Dim rawArgument As Variant
    Dim argumentObject As Object
    Dim argumentValue As ROneCOne
    Dim result As ROneCOne

    Set argumentValues = New Collection
    For Each rawArgument In arguments
        Set argumentValue = New ROneCOne
        If IsObject(rawArgument) Then
            Set argumentObject = rawArgument
            argumentValue.ConfigureObjectValue argumentObject
        Else
            argumentValue.ConfigureScalarValue rawArgument
        End If
        argumentValues.Add argumentValue
    Next rawArgument

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Then
        If argumentValues.Count <> 1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "A sequence indexer requires exactly one argument."
        End If
        Set argumentValue = argumentValues.Item(1)
        If argumentValue.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "A sequence index must be numeric."
        End If
        Set result = WrappedItem(CLng(argumentValue.InternalScalarValue))
        If result.InternalHasObjectValue Then
            Set Run = result.InternalObjectValue
        Else
            Run = result.InternalScalarValue
        End If
        Exit Function
    End If

    Set result = InvokeValue(argumentValues)

    If result.InternalHasObjectValue Then
        Set Run = result.InternalObjectValue
    Else
        Run = result.InternalScalarValue
    End If
End Function

Public Function PipeTo(ByVal nextDelegate As ROneCOne) As ROneCOne
Attribute PipeTo.VB_Description = "Pipes this result into a one-parameter delegate."
    Dim composition As ROneCOne

    RequireDelegate Me, "PipeTo source"
    RequireDelegate nextDelegate, "PipeTo continuation"
    If nextDelegate.Arity <> 1 And nextDelegate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "PipeTo", _
            "The continuation must accept one value."
    End If

    Set composition = New ROneCOne
    composition.ConfigureComposition Me, nextDelegate
    Set PipeTo = composition
End Function

Public Function Add(ByVal operand As Variant) As ROneCOne
Attribute Add.VB_Description = "Adds a list item or creates an addition expression."
    Dim wrapped As ROneCOne

    If mRole = ROLE_LIST Then
        Set wrapped = WrapElement(operand, "Add")
        mItems.Add wrapped
        AddUnwrappedValue mEnumerationValues, wrapped
        mVersion = mVersion + 1
        Set Add = Me
    Else
        Set Add = CreateBinaryExpression(OP_ADD, operand)
    End If
End Function

Public Function Subtract(ByVal operand As Variant) As ROneCOne
Attribute Subtract.VB_Description = "Creates a subtraction expression."
    Set Subtract = CreateBinaryExpression(OP_SUBTRACT, operand)
End Function

Public Function Multiply(ByVal operand As Variant) As ROneCOne
Attribute Multiply.VB_Description = "Creates a multiplication expression."
    Set Multiply = CreateBinaryExpression(OP_MULTIPLY, operand)
End Function

Public Function Divide(ByVal operand As Variant) As ROneCOne
Attribute Divide.VB_Description = "Creates a division expression."
    Set Divide = CreateBinaryExpression(OP_DIVIDE, operand)
End Function

Public Function Modulo(ByVal operand As Variant) As ROneCOne
Attribute Modulo.VB_Description = "Creates a modulo expression."
    Set Modulo = CreateBinaryExpression(OP_MODULO, operand)
End Function

Public Function Negate() As ROneCOne
Attribute Negate.VB_Description = "Creates a numeric negation expression."
    Set Negate = CreateUnaryExpression(OP_NEGATE)
End Function

Public Function Concat(ByVal operand As Variant) As ROneCOne
Attribute Concat.VB_Description = "Creates a string concatenation expression."
    Set Concat = CreateBinaryExpression(OP_CONCAT, operand)
End Function

Public Function EqualTo(ByVal operand As Variant) As ROneCOne
Attribute EqualTo.VB_Description = "Creates an equality expression."
    Set EqualTo = CreateBinaryExpression(OP_EQUAL, operand)
End Function

Public Function NotEqualTo(ByVal operand As Variant) As ROneCOne
Attribute NotEqualTo.VB_Description = "Creates an inequality expression."
    Set NotEqualTo = CreateBinaryExpression(OP_NOT_EQUAL, operand)
End Function

Public Function LessThan(ByVal operand As Variant) As ROneCOne
Attribute LessThan.VB_Description = "Creates a less-than expression."
    Set LessThan = CreateBinaryExpression(OP_LESS_THAN, operand)
End Function

Public Function LessThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute LessThanOrEqual.VB_Description = "Creates a less-than-or-equal expression."
    Set LessThanOrEqual = CreateBinaryExpression(OP_LESS_THAN_OR_EQUAL, operand)
End Function

Public Function GreaterThan(ByVal operand As Variant) As ROneCOne
Attribute GreaterThan.VB_Description = "Creates a greater-than expression."
    Set GreaterThan = CreateBinaryExpression(OP_GREATER_THAN, operand)
End Function

Public Function GreaterThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute GreaterThanOrEqual.VB_Description = "Creates a greater-than-or-equal expression."
    Set GreaterThanOrEqual = CreateBinaryExpression(OP_GREATER_THAN_OR_EQUAL, operand)
End Function

Public Function AndAlso(ByVal operand As Variant) As ROneCOne
Attribute AndAlso.VB_Description = "Creates a short-circuit Boolean AND expression."
    Set AndAlso = CreateBinaryExpression(OP_AND_ALSO, operand)
End Function

Public Function OrElse(ByVal operand As Variant) As ROneCOne
Attribute OrElse.VB_Description = "Creates a short-circuit Boolean OR expression."
    Set OrElse = CreateBinaryExpression(OP_OR_ELSE, operand)
End Function

Public Function NotExpression() As ROneCOne
Attribute NotExpression.VB_Description = "Creates a Boolean NOT expression."
    Set NotExpression = CreateUnaryExpression(OP_NOT)
End Function

Public Property Get Item(ByVal index As Long) As Variant
Attribute Item.VB_Description = "Gets or sets a zero-based typed list item."
    Dim wrapped As ROneCOne

    Set wrapped = WrappedItem(index)
    If wrapped.InternalHasObjectValue Then
        Set Item = wrapped.InternalObjectValue
    Else
        Item = wrapped.InternalScalarValue
    End If
End Property

Public Property Let Item(ByVal index As Long, ByVal value As Variant)
    ReplaceItem index, WrapElement(value, "Item")
End Property

Public Property Set Item(ByVal index As Long, ByVal value As Object)
    ReplaceItem index, WrapElement(value, "Item")
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "Gets the number of elements in the sequence."
    Dim values As Collection

    Set values = MaterializeWrapped
    Count = values.Count
End Property

Public Property Get GenericTypeName() As String
Attribute GenericTypeName.VB_Description = "Gets the C#-style generic sequence type name."
    RequireSequence Me, "GenericTypeName"
    If Len(mElementObjectType) <> 0 Then
        GenericTypeName = "List<" & mElementObjectType & ">"
    Else
        GenericTypeName = "List<" & TypeNameForVarType(mElementType) & ">"
    End If
End Property

Public Sub AddRange(ByVal values As ROneCOne)
Attribute AddRange.VB_Description = "Appends a typed sequence atomically to this list."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireList Me, "AddRange"
    RequireSequence values, "AddRange values"
    Set incoming = values.MaterializeWrapped
    For Each wrapped In incoming
        ValidateElementWrapped wrapped, "AddRange"
    Next wrapped
    For Each wrapped In incoming
        mItems.Add wrapped
        AddUnwrappedValue mEnumerationValues, wrapped
    Next wrapped
    If incoming.Count > 0 Then mVersion = mVersion + 1
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Removes every item from this typed list."
    RequireList Me, "Clear"
    Set mItems = New Collection
    Set mEnumerationValues = New Collection
    mVersion = mVersion + 1
End Sub

Public Function Contains(ByVal value As Variant) As Boolean
Attribute Contains.VB_Description = "Tests whether a typed sequence contains a value."
    Contains = (IndexOf(value) >= 0)
End Function

Public Function IndexOf(ByVal value As Variant) As Long
Attribute IndexOf.VB_Description = "Gets the first zero-based index of a value, or -1."
    Dim candidate As ROneCOne
    Dim index As Long
    Dim values As Collection

    RequireSequence Me, "IndexOf"
    Set candidate = WrapElement(value, "IndexOf")
    Set values = MaterializeWrapped
    IndexOf = -1
    For index = 1 To values.Count
        If ValuesEqual(values.Item(index), candidate) Then
            IndexOf = index - 1
            Exit Function
        End If
    Next index
End Function

Public Sub Insert(ByVal index As Long, ByVal value As Variant)
Attribute Insert.VB_Description = "Inserts a typed item at a zero-based list index."
    Dim wrapped As ROneCOne

    RequireList Me, "Insert"
    If index < 0 Or index > mItems.Count Then
        RaiseIndexError "Insert", index
    End If
    Set wrapped = WrapElement(value, "Insert")
    If index = mItems.Count Then
        mItems.Add wrapped
        AddUnwrappedValue mEnumerationValues, wrapped
    Else
        mItems.Add wrapped, Before:=index + 1
        AddUnwrappedValue mEnumerationValues, wrapped, index + 1
    End If
    mVersion = mVersion + 1
End Sub

Public Function Remove(ByVal value As Variant) As Boolean
Attribute Remove.VB_Description = "Removes the first matching typed list item."
    Dim index As Long

    RequireList Me, "Remove"
    index = IndexOf(value)
    If index >= 0 Then
        mItems.Remove index + 1
        mEnumerationValues.Remove index + 1
        mVersion = mVersion + 1
        Remove = True
    End If
End Function

Public Sub RemoveAt(ByVal index As Long)
Attribute RemoveAt.VB_Description = "Removes the item at a zero-based list index."
    RequireList Me, "RemoveAt"
    ValidateIndex index, mItems.Count, "RemoveAt"
    mItems.Remove index + 1
    mEnumerationValues.Remove index + 1
    mVersion = mVersion + 1
End Sub

Public Function Where(ByVal predicate As ROneCOne) As ROneCOne
Attribute Where.VB_Description = "Defers filtering until the sequence is consumed."
    ValidateUnaryDelegate predicate, "Where"
    Set Where = CreateQuery(QUERY_WHERE, predicate, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function SelectItems( _
    ByVal selector As ROneCOne, _
    ByVal resultTypeToken As Variant _
) As ROneCOne
Attribute SelectItems.VB_Description = "Defers a strongly typed LINQ Select projection."
    Dim objectType As String
    Dim valueType As VbVarType

    ValidateUnaryDelegate selector, "SelectItems"
    ResolveTypeDescriptor resultTypeToken, valueType, objectType, "SelectItems"
    Set SelectItems = CreateQuery(QUERY_SELECT, selector, Nothing, 0, _
        valueType, objectType)
End Function

Public Function Take(ByVal count As Long) As ROneCOne
Attribute Take.VB_Description = "Defers taking up to a number of leading elements."
    If count < 0 Then count = 0
    Set Take = CreateQuery(QUERY_TAKE, Nothing, Nothing, count, _
        mElementType, mElementObjectType)
End Function

Public Function Skip(ByVal count As Long) As ROneCOne
Attribute Skip.VB_Description = "Defers bypassing a number of leading elements."
    If count < 0 Then count = 0
    Set Skip = CreateQuery(QUERY_SKIP, Nothing, Nothing, count, _
        mElementType, mElementObjectType)
End Function

Public Function Distinct() As ROneCOne
Attribute Distinct.VB_Description = "Defers removal of duplicate sequence values."
    Set Distinct = CreateQuery(QUERY_DISTINCT, Nothing, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function OrderBy(ByVal keySelector As ROneCOne) As ROneCOne
Attribute OrderBy.VB_Description = "Defers stable ascending ordering by a selected key."
    ValidateUnaryDelegate keySelector, "OrderBy"
    Set OrderBy = CreateQuery(QUERY_ORDER_BY, keySelector, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function OrderByDescending(ByVal keySelector As ROneCOne) As ROneCOne
Attribute OrderByDescending.VB_Description = "Defers stable descending ordering by a selected key."
    ValidateUnaryDelegate keySelector, "OrderByDescending"
    Set OrderByDescending = CreateQuery( _
        QUERY_ORDER_BY_DESCENDING, keySelector, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function Append(ByVal element As Variant) As ROneCOne
Attribute Append.VB_Description = "Defers appending one typed element to a sequence."
    Dim wrapped As ROneCOne

    RequireSequence Me, "Append"
    Set wrapped = WrapElement(element, "Append")
    Set Append = CreateQuery(QUERY_APPEND, Nothing, wrapped, 0, _
        mElementType, mElementObjectType)
End Function

Public Function Prepend(ByVal element As Variant) As ROneCOne
Attribute Prepend.VB_Description = "Defers prepending one typed element to a sequence."
    Dim wrapped As ROneCOne

    RequireSequence Me, "Prepend"
    Set wrapped = WrapElement(element, "Prepend")
    Set Prepend = CreateQuery(QUERY_PREPEND, Nothing, wrapped, 0, _
        mElementType, mElementObjectType)
End Function

Public Function Reverse() As ROneCOne
Attribute Reverse.VB_Description = "Defers reversing a sequence."
    Set Reverse = CreateQuery(QUERY_REVERSE, Nothing, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function AnyItem(Optional ByVal predicate As Variant) As Boolean
Attribute AnyItem.VB_Description = "Tests whether any element exists or matches a predicate."
    Dim predicateValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If IsMissing(predicate) Then
        AnyItem = (values.Count <> 0)
        Exit Function
    End If
    Set predicateValue = RequireROneCOne(predicate, "Any predicate")
    ValidateUnaryDelegate predicateValue, "Any"
    For Each wrapped In values
        If InvokePredicate(predicateValue, wrapped, "Any") Then
            AnyItem = True
            Exit Function
        End If
    Next wrapped
End Function

Public Function All(ByVal predicate As ROneCOne) As Boolean
Attribute All.VB_Description = "Tests whether every element matches a predicate."
    Dim values As Collection
    Dim wrapped As ROneCOne

    ValidateUnaryDelegate predicate, "All"
    Set values = MaterializeWrapped
    All = True
    For Each wrapped In values
        If Not InvokePredicate(predicate, wrapped, "All") Then
            All = False
            Exit Function
        End If
    Next wrapped
End Function

Public Function First(Optional ByVal predicate As Variant) As Variant
Attribute First.VB_Description = "Returns the first element, optionally matching a predicate."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = RequireROneCOne(predicate, "First predicate")
        ValidateUnaryDelegate predicateValue, "First"
    End If
    Set wrapped = FindTerminalElement(predicateValue, False, "First")
    If wrapped.InternalHasObjectValue Then
        Set First = wrapped.InternalObjectValue
    Else
        First = wrapped.InternalScalarValue
    End If
End Function

Public Function Last(Optional ByVal predicate As Variant) As Variant
Attribute Last.VB_Description = "Returns the last element, optionally matching a predicate."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = RequireROneCOne(predicate, "Last predicate")
        ValidateUnaryDelegate predicateValue, "Last"
    End If
    Set wrapped = FindTerminalElement(predicateValue, True, "Last")
    If wrapped.InternalHasObjectValue Then
        Set Last = wrapped.InternalObjectValue
    Else
        Last = wrapped.InternalScalarValue
    End If
End Function

Public Function Sum(Optional ByVal selector As Variant) As Variant
Attribute Sum.VB_Description = "Sums numeric elements or selected numeric values."
    Dim values As Collection
    Dim wrapped As ROneCOne
    Dim selected As ROneCOne
    Dim selectorValue As ROneCOne
    Dim total As Variant

    Set values = MaterializeWrapped
    total = 0
    If Not IsMissing(selector) Then
        Set selectorValue = RequireROneCOne(selector, "Sum selector")
        ValidateUnaryDelegate selectorValue, "Sum"
    End If
    For Each wrapped In values
        Set selected = SelectTerminalValue(wrapped, selectorValue)
        total = total + RequireScalar(selected, "Sum")
    Next wrapped
    Sum = total
End Function

Public Function Average(Optional ByVal selector As Variant) As Double
Attribute Average.VB_Description = "Averages numeric elements or selected numeric values."
    Dim selected As ROneCOne
    Dim selectorValue As ROneCOne
    Dim total As Double
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence "Average"
    If Not IsMissing(selector) Then
        Set selectorValue = RequireROneCOne(selector, "Average selector")
        ValidateUnaryDelegate selectorValue, "Average"
    End If
    For Each wrapped In values
        Set selected = SelectTerminalValue(wrapped, selectorValue)
        total = total + CDbl(RequireScalar(selected, "Average"))
    Next wrapped
    Average = total / CDbl(values.Count)
End Function

Public Function Min(Optional ByVal selector As Variant) As Variant
Attribute Min.VB_Description = "Returns the minimum element or selected value."
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = RequireROneCOne(selector, "Min selector")
        ValidateUnaryDelegate selectorValue, "Min"
    End If
    Set wrapped = ExtremeValue(selectorValue, False, "Min")
    If wrapped.InternalHasObjectValue Then
        Set Min = wrapped.InternalObjectValue
    Else
        Min = wrapped.InternalScalarValue
    End If
End Function

Public Function Max(Optional ByVal selector As Variant) As Variant
Attribute Max.VB_Description = "Returns the maximum element or selected value."
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = RequireROneCOne(selector, "Max selector")
        ValidateUnaryDelegate selectorValue, "Max"
    End If
    Set wrapped = ExtremeValue(selectorValue, True, "Max")
    If wrapped.InternalHasObjectValue Then
        Set Max = wrapped.InternalObjectValue
    Else
        Max = wrapped.InternalScalarValue
    End If
End Function

Public Function ToList() As ROneCOne
Attribute ToList.VB_Description = "Immediately materializes a deferred sequence as List<T>."
    Dim listValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "ToList"
    Set listValue = New ROneCOne
    listValue.ConfigureList mElementType, mElementObjectType
    Set values = MaterializeWrapped
    For Each wrapped In values
        listValue.InternalAddWrapped wrapped
    Next wrapped
    Set ToList = listValue
End Function

Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Immediately materializes a sequence as a zero-based array."
    Dim index As Long
    Dim output() As Variant
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    ReDim output(0 To values.Count - 1)
    For index = 1 To values.Count
        Set wrapped = values.Item(index)
        If wrapped.InternalHasObjectValue Then
            Set output(index - 1) = wrapped.InternalObjectValue
        Else
            output(index - 1) = wrapped.InternalScalarValue
        End If
    Next index
    ToArray = output
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets the VBA For Each enumerator for this sequence."
Attribute NewEnum.VB_UserMemId = -4
    If mRole = ROLE_QUERY Then
        If mEnumerationValues Is Nothing Or _
            mEnumerationVersion <> InternalSequenceVersion Then
            Set mEnumerationValues = MaterializeValues
            mEnumerationVersion = InternalSequenceVersion
        End If
    End If
    Set NewEnum = mEnumerationValues.[_NewEnum]
End Property

Public Property Get Arity() As Long
Attribute Arity.VB_Description = "Gets the number of delegate parameters, or -1 when dynamic."
    Select Case mRole
        Case ROLE_LAMBDA
            Arity = mParameters.Count
        Case ROLE_METHOD
            Arity = mMethodArity
        Case ROLE_COMPOSITION
            Arity = mLeft.Arity
        Case Else
            Arity = -1
    End Select
End Property

Public Property Get TypeMismatchError() As Long
Attribute TypeMismatchError.VB_Description = "Gets the ROneCOne type-mismatch error number."
    TypeMismatchError = ERROR_TYPE_MISMATCH
End Property

Public Property Get UnboundParameterError() As Long
Attribute UnboundParameterError.VB_Description = "Gets the unbound-parameter error number."
    UnboundParameterError = ERROR_UNBOUND_PARAMETER
End Property

Friend Property Get InternalRole() As Long
Attribute InternalRole.VB_MemberFlags = "40"
    InternalRole = mRole
End Property

Friend Property Get InternalOperator() As Long
Attribute InternalOperator.VB_MemberFlags = "40"
    InternalOperator = mOperator
End Property

Friend Property Get InternalLeft() As ROneCOne
Attribute InternalLeft.VB_MemberFlags = "40"
    Set InternalLeft = mLeft
End Property

Friend Property Get InternalRight() As ROneCOne
Attribute InternalRight.VB_MemberFlags = "40"
    Set InternalRight = mRight
End Property

Friend Property Get InternalHasObjectValue() As Boolean
Attribute InternalHasObjectValue.VB_MemberFlags = "40"
    InternalHasObjectValue = mHasObjectValue
End Property

Friend Property Get InternalObjectValue() As Object
Attribute InternalObjectValue.VB_MemberFlags = "40"
    Set InternalObjectValue = mObjectValue
End Property

Friend Property Get InternalScalarValue() As Variant
Attribute InternalScalarValue.VB_MemberFlags = "40"
    InternalScalarValue = mScalarValue
End Property

Friend Property Get InternalExpectedType() As VbVarType
Attribute InternalExpectedType.VB_MemberFlags = "40"
    InternalExpectedType = mExpectedType
End Property

Friend Property Get InternalExpectedObjectType() As String
Attribute InternalExpectedObjectType.VB_MemberFlags = "40"
    InternalExpectedObjectType = mExpectedObjectType
End Property

Friend Sub ConfigureParameter( _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String _
)
    mRole = ROLE_PARAMETER
    mExpectedType = expectedType
    mExpectedObjectType = expectedObjectType
End Sub

Friend Sub ConfigureScalarValue(ByVal scalarValue As Variant)
    mRole = ROLE_VALUE
    mScalarValue = scalarValue
    mHasObjectValue = False
End Sub

Friend Sub ConfigureObjectValue(ByVal objectValue As Object)
    mRole = ROLE_VALUE
    Set mObjectValue = objectValue
    mHasObjectValue = True
End Sub

Friend Sub ConfigureExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    Set mRight = rightValue
End Sub

Friend Sub ConfigureLambda( _
    ByVal body As ROneCOne, _
    ByVal parameters As Collection _
)
    mRole = ROLE_LAMBDA
    Set mLeft = body
    Set mParameters = parameters
End Sub

Friend Sub ConfigureMethod( _
    ByVal target As Object, _
    ByVal methodName As String, _
    ByVal parameterCount As Long, _
    ByVal returnsObject As Boolean _
)
    mRole = ROLE_METHOD
    Set mTarget = target
    mMethodName = methodName
    mMethodArity = parameterCount
    mMethodReturnsObject = returnsObject
End Sub

Friend Sub ConfigureComposition( _
    ByVal firstDelegate As ROneCOne, _
    ByVal nextDelegate As ROneCOne _
)
    mRole = ROLE_COMPOSITION
    Set mLeft = firstDelegate
    Set mNext = nextDelegate
End Sub

Friend Sub ConfigureList( _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String _
)
    mRole = ROLE_LIST
    mElementType = elementType
    mElementObjectType = elementObjectType
    Set mItems = New Collection
    Set mEnumerationValues = New Collection
    mEnumerationVersion = 0
End Sub

Friend Sub ConfigureQuery( _
    ByVal source As ROneCOne, _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal queryValue As ROneCOne, _
    ByVal queryCount As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String _
)
    mRole = ROLE_QUERY
    Set mSource = source
    mOperator = queryOperator
    Set mQueryDelegate = queryDelegate
    Set mQueryValue = queryValue
    mQueryCount = queryCount
    mElementType = elementType
    mElementObjectType = elementObjectType
    mEnumerationVersion = -1
End Sub

Friend Sub InternalAddWrapped(ByVal wrapped As ROneCOne)
    ValidateElementWrapped wrapped, "ToList"
    mItems.Add wrapped
    AddUnwrappedValue mEnumerationValues, wrapped
    mVersion = mVersion + 1
End Sub

Friend Property Get InternalSequenceVersion() As Long
Attribute InternalSequenceVersion.VB_MemberFlags = "40"
    If mRole = ROLE_LIST Then
        InternalSequenceVersion = mVersion
    ElseIf mRole = ROLE_QUERY Then
        InternalSequenceVersion = mSource.InternalSequenceVersion
    End If
End Property

Friend Function MaterializeWrapped() As Collection
    Dim sourceValues As Collection
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "Materialize"
    If mRole = ROLE_LIST Then
        Set values = New Collection
        For Each wrapped In mItems
            values.Add wrapped
        Next wrapped
        Set MaterializeWrapped = values
        Exit Function
    End If

    Set sourceValues = mSource.MaterializeWrapped
    Select Case mOperator
        Case QUERY_WHERE
            Set values = ApplyWhere(sourceValues)
        Case QUERY_SELECT
            Set values = ApplySelect(sourceValues)
        Case QUERY_TAKE
            Set values = ApplyTake(sourceValues)
        Case QUERY_SKIP
            Set values = ApplySkip(sourceValues)
        Case QUERY_DISTINCT
            Set values = ApplyDistinct(sourceValues)
        Case QUERY_ORDER_BY
            Set values = ApplyOrder(sourceValues, False)
        Case QUERY_ORDER_BY_DESCENDING
            Set values = ApplyOrder(sourceValues, True)
        Case QUERY_APPEND
            Set values = CopyWrappedCollection(sourceValues)
            values.Add mQueryValue
        Case QUERY_PREPEND
            Set values = New Collection
            values.Add mQueryValue
            AddWrappedRange values, sourceValues
        Case QUERY_REVERSE
            Set values = ApplyReverse(sourceValues)
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Materialize", _
                "Unknown deferred query operator."
    End Select
    Set MaterializeWrapped = values
End Function

Private Function CreateQuery( _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal queryValue As ROneCOne, _
    ByVal queryCount As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String _
) As ROneCOne
    Dim query As ROneCOne

    RequireSequence Me, "LINQ"
    Set query = New ROneCOne
    query.ConfigureQuery Me, queryOperator, queryDelegate, queryValue, _
        queryCount, elementType, elementObjectType
    Set CreateQuery = query
End Function

Private Sub ResolveTypeDescriptor( _
    ByVal typeToken As Variant, _
    ByRef elementType As VbVarType, _
    ByRef elementObjectType As String, _
    ByVal memberName As String _
)
    Dim prototype As Object

    If IsObject(typeToken) Then
        Set prototype = typeToken
        If prototype Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The user-defined type prototype cannot be Nothing."
        End If
        elementType = vbObject
        elementObjectType = TypeName(prototype)
    Else
        ValidateTypeToken typeToken, memberName
        elementType = CLng(typeToken)
        elementObjectType = vbNullString
    End If
End Sub

Private Sub ValidateTypeToken(ByVal typeToken As Variant, ByVal memberName As String)
    Dim token As Long

    If Not IsNumeric(typeToken) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The type token must be a VbVarType value or an object prototype."
    End If
    token = CLng(typeToken)
    Select Case token
        Case vbEmpty, vbNull, vbInteger, vbLong, vbSingle, vbDouble
        Case vbCurrency, vbDate, vbString, vbObject, vbError, vbBoolean
        Case vbVariant, vbDecimal, vbByte, vbLongLong
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The type token is not a supported VbVarType value."
    End Select
End Sub

Private Function WrapElement( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidateObject As Object
    Dim wrapped As ROneCOne

    RequireSequence Me, memberName
    Set wrapped = New ROneCOne
    If IsObject(candidate) Then
        Set candidateObject = candidate
        wrapped.ConfigureObjectValue candidateObject
    Else
        wrapped.ConfigureScalarValue candidate
    End If
    ValidateElementWrapped wrapped, memberName
    Set WrapElement = wrapped
End Function

Private Sub ValidateElementWrapped( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim actualObject As Object

    If mElementType = vbVariant Then Exit Sub
    If mElementType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then
            RaiseElementTypeMismatch wrapped, memberName
        End If
        If Len(mElementObjectType) = 0 Then Exit Sub
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then Exit Sub
        If StrComp(TypeName(actualObject), mElementObjectType, vbBinaryCompare) <> 0 Then
            RaiseElementTypeMismatch wrapped, memberName
        End If
        Exit Sub
    End If

    If wrapped.InternalHasObjectValue Then
        RaiseElementTypeMismatch wrapped, memberName
    End If
    If VarType(wrapped.InternalScalarValue) <> mElementType Then
        RaiseElementTypeMismatch wrapped, memberName
    End If
End Sub

Private Sub RaiseElementTypeMismatch( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim actualName As String
    Dim actualObject As Object
    Dim expectedName As String

    If Len(mElementObjectType) <> 0 Then
        expectedName = mElementObjectType
    Else
        expectedName = TypeNameForVarType(mElementType)
    End If
    If wrapped.InternalHasObjectValue Then
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then
            actualName = "Nothing"
        Else
            actualName = TypeName(actualObject)
        End If
    Else
        actualName = TypeName(wrapped.InternalScalarValue)
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "List<" & expectedName & "> cannot contain " & actualName & "."
End Sub

Private Sub RequireSequence(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The sequence cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_LIST And candidate.InternalRole <> ROLE_QUERY Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            "This operation requires a typed sequence."
    End If
End Sub

Private Sub RequireList(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The list cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_LIST Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            "This operation requires a materialized List<T>."
    End If
End Sub

Private Sub ValidateIndex( _
    ByVal index As Long, _
    ByVal itemCount As Long, _
    ByVal memberName As String _
)
    If index < 0 Or index >= itemCount Then RaiseIndexError memberName, index
End Sub

Private Sub RaiseIndexError(ByVal memberName As String, ByVal index As Long)
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "Index " & CStr(index) & " is outside the sequence."
End Sub

Private Function WrappedItem(ByVal index As Long) As ROneCOne
    Dim values As Collection

    Set values = MaterializeWrapped
    ValidateIndex index, values.Count, "Item"
    Set WrappedItem = values.Item(index + 1)
End Function

Private Sub ReplaceItem(ByVal index As Long, ByVal wrapped As ROneCOne)
    Dim current As ROneCOne
    Dim position As Long
    Dim replacement As Collection

    RequireList Me, "Item"
    ValidateIndex index, mItems.Count, "Item"
    Set replacement = New Collection
    For position = 1 To mItems.Count
        If position = index + 1 Then
            replacement.Add wrapped
        Else
            Set current = mItems.Item(position)
            replacement.Add current
        End If
    Next position
    Set mItems = replacement
    Set mEnumerationValues = MaterializeValues
    mVersion = mVersion + 1
End Sub

Private Sub ValidateUnaryDelegate( _
    ByVal delegateValue As ROneCOne, _
    ByVal memberName As String _
)
    RequireSequence Me, memberName
    RequireDelegate delegateValue, memberName & " delegate"
    If delegateValue.Arity <> 1 And delegateValue.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The delegate must accept exactly one value."
    End If
End Sub

Private Function InvokeUnary( _
    ByVal delegateValue As ROneCOne, _
    ByVal wrapped As ROneCOne _
) As ROneCOne
    Dim arguments As Collection

    Set arguments = New Collection
    arguments.Add wrapped
    Set InvokeUnary = delegateValue.InvokeValue(arguments)
End Function

Private Function InvokePredicate( _
    ByVal predicate As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    Dim result As ROneCOne

    Set result = InvokeUnary(predicate, wrapped)
    InvokePredicate = CBool(RequireScalar(result, memberName))
End Function

Private Function FindTerminalElement( _
    ByVal predicate As ROneCOne, _
    ByVal takeLast As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim found As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        If predicate Is Nothing Then
            matches = True
        Else
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            Set found = wrapped
            If Not takeLast Then
                Set FindTerminalElement = found
                Exit Function
            End If
        End If
    Next wrapped
    If found Is Nothing Then RaiseEmptySequence memberName
    Set FindTerminalElement = found
End Function

Private Function SelectTerminalValue( _
    ByVal wrapped As ROneCOne, _
    ByVal selector As ROneCOne _
) As ROneCOne
    If selector Is Nothing Then
        Set SelectTerminalValue = wrapped
    Else
        Set SelectTerminalValue = InvokeUnary(selector, wrapped)
    End If
End Function

Private Function ExtremeValue( _
    ByVal selector As ROneCOne, _
    ByVal findMaximum As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim best As ROneCOne
    Dim candidate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence memberName
    For Each wrapped In values
        Set candidate = SelectTerminalValue(wrapped, selector)
        If best Is Nothing Then
            Set best = candidate
        ElseIf findMaximum And CompareWrapped(candidate, best, memberName) > 0 Then
            Set best = candidate
        ElseIf Not findMaximum And CompareWrapped(candidate, best, memberName) < 0 Then
            Set best = candidate
        End If
    Next wrapped
    Set ExtremeValue = best
End Function

Private Sub RaiseEmptySequence(ByVal memberName As String)
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "The sequence contains no matching elements."
End Sub

Private Function MaterializeValues() As Collection
    Dim values As Collection
    Dim wrapped As ROneCOne
    Dim wrappedValues As Collection

    Set values = New Collection
    Set wrappedValues = MaterializeWrapped
    For Each wrapped In wrappedValues
        AddUnwrappedValue values, wrapped
    Next wrapped
    Set MaterializeValues = values
End Function

Private Sub AddUnwrappedValue( _
    ByVal values As Collection, _
    ByVal wrapped As ROneCOne, _
    Optional ByVal beforePosition As Long = 0 _
)
    Dim actualObject As Object
    Dim actualValue As Variant

    If wrapped.InternalHasObjectValue Then
        Set actualObject = wrapped.InternalObjectValue
        Set actualValue = actualObject
    Else
        actualValue = wrapped.InternalScalarValue
    End If
    If beforePosition > 0 Then
        values.Add actualValue, Before:=beforePosition
    Else
        values.Add actualValue
    End If
End Sub

Private Function ApplyWhere(ByVal sourceValues As Collection) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If InvokePredicate(mQueryDelegate, wrapped, "Where") Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyWhere = output
End Function

Private Function ApplySelect(ByVal sourceValues As Collection) As Collection
    Dim output As Collection
    Dim selected As ROneCOne
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        Set selected = InvokeUnary(mQueryDelegate, wrapped)
        ValidateElementWrapped selected, "SelectItems"
        output.Add selected
    Next wrapped
    Set ApplySelect = output
End Function

Private Function ApplyTake(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection
    Dim takeCount As Long

    Set output = New Collection
    takeCount = mQueryCount
    If takeCount > sourceValues.Count Then takeCount = sourceValues.Count
    For index = 1 To takeCount
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyTake = output
End Function

Private Function ApplySkip(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = mQueryCount + 1 To sourceValues.Count
        output.Add sourceValues.Item(index)
    Next index
    Set ApplySkip = output
End Function

Private Function ApplyDistinct(ByVal sourceValues As Collection) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains(output, wrapped) Then output.Add wrapped
    Next wrapped
    Set ApplyDistinct = output
End Function

Private Function ApplyOrder( _
    ByVal sourceValues As Collection, _
    ByVal descending As Boolean _
) As Collection
    Dim comparison As Long
    Dim index As Long
    Dim key As ROneCOne
    Dim keys As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set keys = New Collection
    Set output = New Collection
    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        index = 1
        Do While index <= keys.Count
            comparison = CompareWrapped(key, keys.Item(index), "OrderBy")
            If descending Then comparison = -comparison
            If comparison < 0 Then Exit Do
            index = index + 1
        Loop
        If index > keys.Count Then
            keys.Add key
            output.Add wrapped
        Else
            keys.Add key, Before:=index
            output.Add wrapped, Before:=index
        End If
    Next wrapped
    Set ApplyOrder = output
End Function

Private Function ApplyReverse(ByVal sourceValues As Collection) As Collection
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = sourceValues.Count To 1 Step -1
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyReverse = output
End Function

Private Function CopyWrappedCollection( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection

    Set output = New Collection
    AddWrappedRange output, sourceValues
    Set CopyWrappedCollection = output
End Function

Private Sub AddWrappedRange( _
    ByVal target As Collection, _
    ByVal sourceValues As Collection _
)
    Dim wrapped As ROneCOne

    For Each wrapped In sourceValues
        target.Add wrapped
    Next wrapped
End Sub

Private Function WrappedCollectionContains( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne _
) As Boolean
    Dim wrapped As ROneCOne

    For Each wrapped In values
        If ValuesEqual(wrapped, candidate) Then
            WrappedCollectionContains = True
            Exit Function
        End If
    Next wrapped
End Function

Private Function CompareWrapped( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim leftScalar As Variant
    Dim rightScalar As Variant

    If leftValue.InternalHasObjectValue Or rightValue.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "Object values require a scalar key selector for ordering."
    End If
    leftScalar = leftValue.InternalScalarValue
    rightScalar = rightValue.InternalScalarValue
    If leftScalar = rightScalar Then
        CompareWrapped = 0
    ElseIf leftScalar < rightScalar Then
        CompareWrapped = -1
    Else
        CompareWrapped = 1
    End If
End Function

Private Function CreateBinaryExpression( _
    ByVal expressionOperator As Long, _
    ByVal operand As Variant _
) As ROneCOne
    Dim expression As ROneCOne
    Dim rightValue As ROneCOne

    RequireExpression Me, "Expression receiver"
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, Me, rightValue
    Set CreateBinaryExpression = expression
End Function

Private Function CreateUnaryExpression(ByVal expressionOperator As Long) As ROneCOne
    Dim expression As ROneCOne
    Dim emptyRight As ROneCOne

    RequireExpression Me, "Expression receiver"
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, Me, emptyRight
    Set CreateUnaryExpression = expression
End Function

Private Function CoerceExpression(ByVal candidate As Variant) As ROneCOne
    Dim candidateObject As Object
    Dim node As ROneCOne

    If IsObject(candidate) Then
        Set candidateObject = candidate
        If TypeOf candidateObject Is ROneCOne Then
            Set node = candidateObject
            RequireExpression node, "Expression operand"
            Set CoerceExpression = node
            Exit Function
        End If
    End If

    Set node = New ROneCOne
    If IsObject(candidate) Then
        node.ConfigureObjectValue candidateObject
    Else
        node.ConfigureScalarValue candidate
    End If
    Set CoerceExpression = node
End Function

Friend Function InvokeValue(ByVal arguments As Collection) As ROneCOne
    Dim firstResult As ROneCOne
    Dim nextArguments As Collection

    Select Case mRole
        Case ROLE_LAMBDA
            If arguments.Count <> mParameters.Count Then
                RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                    "Expected " & CStr(mParameters.Count) & " argument(s), received " & _
                    CStr(arguments.Count) & "."
            End If
            Set InvokeValue = EvaluateNode(mLeft, mParameters, arguments)

        Case ROLE_METHOD
            Set InvokeValue = InvokeMethod(arguments)

        Case ROLE_COMPOSITION
            Set firstResult = mLeft.InvokeValue(arguments)
            Set nextArguments = New Collection
            nextArguments.Add firstResult
            Set InvokeValue = mNext.InvokeValue(nextArguments)

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Run", _
                "Only delegates can be invoked."
    End Select
End Function

Private Function EvaluateNode( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection, _
    ByVal arguments As Collection _
) As ROneCOne
    Dim leftResult As ROneCOne
    Dim rightResult As ROneCOne
    Dim scalarResult As Variant
    Dim operatorCode As Long

    Select Case node.InternalRole
        Case ROLE_VALUE
            Set EvaluateNode = node

        Case ROLE_PARAMETER
            Set EvaluateNode = ResolveParameter(node, parameters, arguments)

        Case ROLE_EXPRESSION
            Set leftResult = EvaluateNode(node.InternalLeft, parameters, arguments)
            operatorCode = node.InternalOperator

            If operatorCode = OP_AND_ALSO Then
                If Not CBool(RequireScalar(leftResult, "AndAlso")) Then
                    Set EvaluateNode = CreateScalarResult(False)
                    Exit Function
                End If
            ElseIf operatorCode = OP_OR_ELSE Then
                If CBool(RequireScalar(leftResult, "OrElse")) Then
                    Set EvaluateNode = CreateScalarResult(True)
                    Exit Function
                End If
            End If

            If Not node.InternalRight Is Nothing Then
                Set rightResult = EvaluateNode(node.InternalRight, parameters, arguments)
            End If

            Select Case operatorCode
                Case OP_ADD
                    scalarResult = RequireScalar(leftResult, "Add") + _
                        RequireScalar(rightResult, "Add")
                Case OP_SUBTRACT
                    scalarResult = RequireScalar(leftResult, "Subtract") - _
                        RequireScalar(rightResult, "Subtract")
                Case OP_MULTIPLY
                    scalarResult = RequireScalar(leftResult, "Multiply") * _
                        RequireScalar(rightResult, "Multiply")
                Case OP_DIVIDE
                    scalarResult = RequireScalar(leftResult, "Divide") / _
                        RequireScalar(rightResult, "Divide")
                Case OP_MODULO
                    scalarResult = RequireScalar(leftResult, "Modulo") Mod _
                        RequireScalar(rightResult, "Modulo")
                Case OP_NEGATE
                    scalarResult = -RequireScalar(leftResult, "Negate")
                Case OP_CONCAT
                    scalarResult = CStr(RequireScalar(leftResult, "Concat")) & _
                        CStr(RequireScalar(rightResult, "Concat"))
                Case OP_EQUAL
                    scalarResult = ValuesEqual(leftResult, rightResult)
                Case OP_NOT_EQUAL
                    scalarResult = Not ValuesEqual(leftResult, rightResult)
                Case OP_LESS_THAN
                    scalarResult = RequireScalar(leftResult, "LessThan") < _
                        RequireScalar(rightResult, "LessThan")
                Case OP_LESS_THAN_OR_EQUAL
                    scalarResult = RequireScalar(leftResult, "LessThanOrEqual") <= _
                        RequireScalar(rightResult, "LessThanOrEqual")
                Case OP_GREATER_THAN
                    scalarResult = RequireScalar(leftResult, "GreaterThan") > _
                        RequireScalar(rightResult, "GreaterThan")
                Case OP_GREATER_THAN_OR_EQUAL
                    scalarResult = RequireScalar(leftResult, "GreaterThanOrEqual") >= _
                        RequireScalar(rightResult, "GreaterThanOrEqual")
                Case OP_AND_ALSO
                    scalarResult = CBool(RequireScalar(rightResult, "AndAlso"))
                Case OP_OR_ELSE
                    scalarResult = CBool(RequireScalar(rightResult, "OrElse"))
                Case OP_NOT
                    scalarResult = Not CBool(RequireScalar(leftResult, "NotExpression"))
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "EvaluateNode", _
                        "Unknown expression operator."
            End Select

            Set EvaluateNode = CreateScalarResult(scalarResult)

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "EvaluateNode", _
                "The lambda body contains a non-expression value."
    End Select
End Function

Private Function ResolveParameter( _
    ByVal parameterValue As ROneCOne, _
    ByVal parameters As Collection, _
    ByVal arguments As Collection _
) As ROneCOne
    Dim index As Long
    Dim argumentValue As ROneCOne

    For index = 1 To parameters.Count
        If parameters.Item(index) Is parameterValue Then
            Set argumentValue = arguments.Item(index)
            ValidateParameterValue parameterValue, argumentValue
            Set ResolveParameter = argumentValue
            Exit Function
        End If
    Next index

    RaiseContractError ERROR_UNBOUND_PARAMETER, "ResolveParameter", _
        "The expression contains an unbound parameter."
End Function

Private Sub ValidateParameterValue( _
    ByVal parameterValue As ROneCOne, _
    ByVal argumentValue As ROneCOne _
)
    Dim actualType As VbVarType
    Dim expectedObjectType As String
    Dim expectedType As VbVarType

    expectedType = parameterValue.InternalExpectedType
    expectedObjectType = parameterValue.InternalExpectedObjectType
    If expectedType = vbVariant Then Exit Sub

    If expectedType = vbObject Then
        If Not argumentValue.InternalHasObjectValue Then
            RaiseTypeMismatch parameterValue, argumentValue
        End If
        If Len(expectedObjectType) <> 0 Then
            If StrComp( _
                TypeName(argumentValue.InternalObjectValue), _
                expectedObjectType, _
                vbTextCompare _
            ) <> 0 Then
                RaiseTypeMismatch parameterValue, argumentValue
            End If
        End If
        Exit Sub
    End If

    If argumentValue.InternalHasObjectValue Then
        RaiseTypeMismatch parameterValue, argumentValue
    End If

    actualType = VarType(argumentValue.InternalScalarValue)
    If actualType <> expectedType Then
        RaiseTypeMismatch parameterValue, argumentValue
    End If
End Sub

Private Sub RaiseTypeMismatch( _
    ByVal parameterValue As ROneCOne, _
    ByVal argumentValue As ROneCOne _
)
    Dim expectedName As String
    Dim actualName As String

    If Len(parameterValue.InternalExpectedObjectType) <> 0 Then
        expectedName = parameterValue.InternalExpectedObjectType
    Else
        expectedName = TypeNameForVarType(parameterValue.InternalExpectedType)
    End If

    If argumentValue.InternalHasObjectValue Then
        actualName = TypeName(argumentValue.InternalObjectValue)
    Else
        actualName = TypeName(argumentValue.InternalScalarValue)
    End If

    RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
        "Expected " & expectedName & ", received " & actualName & "."
End Sub

Private Function InvokeMethod(ByVal arguments As Collection) As ROneCOne
    Dim returnedObject As Object
    Dim returnedValue As Variant

    If mMethodArity <> -1 And arguments.Count <> mMethodArity Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Expected " & CStr(mMethodArity) & " argument(s), received " & _
            CStr(arguments.Count) & "."
    End If
    If arguments.Count > 8 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Method delegates support at most 8 arguments."
    End If

    On Error GoTo InvocationFailed
    If mMethodReturnsObject Then
        Select Case arguments.Count
            Case 0
                Set returnedObject = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, Unwrap(arguments.Item(1)))
            Case 2
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)))
            Case 3
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)))
            Case 4
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)))
            Case Else
                RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                    "Object-returning method delegates support at most 4 arguments."
        End Select
        Set InvokeMethod = CreateObjectResult(returnedObject)
    Else
        Select Case arguments.Count
            Case 0
                returnedValue = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, Unwrap(arguments.Item(1)))
            Case 2
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)))
            Case 3
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)))
            Case 4
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)))
            Case 5
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)), _
                    Unwrap(arguments.Item(5)))
            Case 6
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)), _
                    Unwrap(arguments.Item(5)), Unwrap(arguments.Item(6)))
            Case 7
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)), _
                    Unwrap(arguments.Item(5)), Unwrap(arguments.Item(6)), _
                    Unwrap(arguments.Item(7)))
            Case 8
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    Unwrap(arguments.Item(1)), Unwrap(arguments.Item(2)), _
                    Unwrap(arguments.Item(3)), Unwrap(arguments.Item(4)), _
                    Unwrap(arguments.Item(5)), Unwrap(arguments.Item(6)), _
                    Unwrap(arguments.Item(7)), Unwrap(arguments.Item(8)))
        End Select
        Set InvokeMethod = CreateScalarResult(returnedValue)
    End If
    Exit Function

InvocationFailed:
    RaiseContractError ERROR_METHOD_INVOCATION, "Run", _
        "Method '" & mMethodName & "' failed with " & CStr(Err.Number) & ": " & _
        Err.Description
End Function

Private Function Unwrap(ByVal wrapped As ROneCOne) As Variant
    If wrapped.InternalHasObjectValue Then
        Set Unwrap = wrapped.InternalObjectValue
    Else
        Unwrap = wrapped.InternalScalarValue
    End If
End Function

Private Function CreateScalarResult(ByVal scalarValue As Variant) As ROneCOne
    Dim result As ROneCOne

    Set result = New ROneCOne
    result.ConfigureScalarValue scalarValue
    Set CreateScalarResult = result
End Function

Private Function CreateObjectResult(ByVal objectValue As Object) As ROneCOne
    Dim result As ROneCOne

    Set result = New ROneCOne
    result.ConfigureObjectValue objectValue
    Set CreateObjectResult = result
End Function

Private Function RequireScalar( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
) As Variant
    If wrapped Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The expression operand is missing."
    End If
    If wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "This expression requires scalar operands."
    End If
    RequireScalar = wrapped.InternalScalarValue
End Function

Private Function ValuesEqual( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne _
) As Boolean
    If leftValue.InternalHasObjectValue Or rightValue.InternalHasObjectValue Then
        If Not leftValue.InternalHasObjectValue Or Not rightValue.InternalHasObjectValue Then
            ValuesEqual = False
        Else
            ValuesEqual = leftValue.InternalObjectValue Is rightValue.InternalObjectValue
        End If
    Else
        ValuesEqual = (leftValue.InternalScalarValue = rightValue.InternalScalarValue)
    End If
End Function

Private Sub ValidateBindings( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
)
    Select Case node.InternalRole
        Case ROLE_VALUE
            Exit Sub
        Case ROLE_PARAMETER
            If Not CollectionContainsIdentity(parameters, node) Then
                RaiseContractError ERROR_UNBOUND_PARAMETER, "Lambda", _
                    "The expression contains a parameter not declared by this lambda."
            End If
        Case ROLE_EXPRESSION
            ValidateBindings node.InternalLeft, parameters
            If Not node.InternalRight Is Nothing Then
                ValidateBindings node.InternalRight, parameters
            End If
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, "Lambda", _
                "The lambda body must be an expression, parameter, or captured value."
    End Select
End Sub

Private Function RequireROneCOne( _
    ByVal candidate As Variant, _
    ByVal context As String _
) As ROneCOne
    Dim candidateObject As Object

    If Not IsObject(candidate) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "Expected a ROneCOne value."
    End If
    Set candidateObject = candidate
    If candidateObject Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The ROneCOne value cannot be Nothing."
    End If
    If Not TypeOf candidateObject Is ROneCOne Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "Expected a ROneCOne value."
    End If
    Set RequireROneCOne = candidateObject
End Function

Private Sub RequireExpression(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The expression cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            Exit Sub
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, context, _
                "This operation requires an expression value."
    End Select
End Sub

Private Sub RequireDelegate(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The delegate cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_LAMBDA, ROLE_METHOD, ROLE_COMPOSITION
            Exit Sub
        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, context, _
                "This operation requires a delegate."
    End Select
End Sub

Private Function CollectionContainsIdentity( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne _
) As Boolean
    Dim index As Long

    For index = 1 To values.Count
        If values.Item(index) Is candidate Then
            CollectionContainsIdentity = True
            Exit Function
        End If
    Next index
End Function

Private Function TypeNameForVarType(ByVal valueType As VbVarType) As String
    Select Case valueType
        Case vbEmpty
            TypeNameForVarType = "Empty"
        Case vbNull
            TypeNameForVarType = "Null"
        Case vbInteger
            TypeNameForVarType = "Integer"
        Case vbLong
            TypeNameForVarType = "Long"
        Case vbSingle
            TypeNameForVarType = "Single"
        Case vbDouble
            TypeNameForVarType = "Double"
        Case vbCurrency
            TypeNameForVarType = "Currency"
        Case vbDate
            TypeNameForVarType = "Date"
        Case vbString
            TypeNameForVarType = "String"
        Case vbObject
            TypeNameForVarType = "Object"
        Case vbError
            TypeNameForVarType = "Error"
        Case vbBoolean
            TypeNameForVarType = "Boolean"
        Case vbVariant
            TypeNameForVarType = "Variant"
        Case vbDecimal
            TypeNameForVarType = "Decimal"
        Case vbByte
            TypeNameForVarType = "Byte"
        Case vbLongLong
            TypeNameForVarType = "LongLong"
        Case Else
            TypeNameForVarType = "VarType " & CStr(valueType)
    End Select
End Function

Private Sub RaiseContractError( _
    ByVal errorNumber As Long, _
    ByVal memberName As String, _
    ByVal message As String _
)
    Err.Raise errorNumber, "ROneCOne." & memberName, message
End Sub

Private Sub Class_Terminate()
    Set mLeft = Nothing
    Set mRight = Nothing
    Set mNext = Nothing
    Set mObjectValue = Nothing
    Set mParameters = Nothing
    Set mTarget = Nothing
    Set mItems = Nothing
    Set mEnumerationValues = Nothing
    Set mSource = Nothing
    Set mQueryDelegate = Nothing
    Set mQueryValue = Nothing
End Sub

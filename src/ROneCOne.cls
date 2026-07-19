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

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal instancePointer As LongPtr, _
    ByVal functionPointer As LongPtr, _
    ByVal callingConvention As Long, _
    ByVal returnType As Integer, _
    ByVal argumentCount As Long, _
    ByVal argumentTypesPointer As LongPtr, _
    ByVal argumentPointersPointer As LongPtr, _
    ByVal resultPointer As LongPtr _
) As Long

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
Private Const ROLE_PROCEDURE As Long = 9
Private Const ROLE_MULTICAST As Long = 10
Private Const ROLE_NATIVE As Long = 11
Private Const ROLE_REFERENCE As Long = 12
Private Const ROLE_EVENT As Long = 13
Private Const ROLE_TRY As Long = 14
Private Const ROLE_ERROR As Long = 15
Private Const ROLE_WHERE_BUILDER As Long = 16

Private Const CALLCONV_STDCALL As Long = 4
Private Const HRESULT_SUCCESS As Long = 0

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
Private Const OP_MEMBER_VALUE As Long = 17
Private Const OP_MEMBER_OBJECT As Long = 18
Private Const OP_STARTS_WITH As Long = 19
Private Const OP_ENDS_WITH As Long = 20
Private Const OP_CONTAINS_TEXT As Long = 21
Private Const OP_IS_NOTHING As Long = 22
Private Const OP_IS_NOT_NOTHING As Long = 23
Private Const OP_IS_NULL_OR_EMPTY As Long = 24
Private Const OP_MATCHES_PATTERN As Long = 25
Private Const OP_SAFE_MEMBER_VALUE As Long = 26
Private Const OP_SAFE_MEMBER_OBJECT As Long = 27
Private Const OP_STRING_EQUAL As Long = 28
Private Const OP_STRING_NOT_EQUAL As Long = 29
Private Const OP_IN_SEQUENCE As Long = 30
Private Const OP_NOT_IN_SEQUENCE As Long = 31
Private Const OP_ANY_MATCH As Long = 32
Private Const OP_ALL_MATCH As Long = 33
Private Const OP_NONE_MATCH As Long = 34

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
Private Const QUERY_DISTINCT_BY As Long = 11
Private Const QUERY_WHERE_NOT As Long = 12

Private Const ERROR_INVALID_ARGUMENT As Long = vbObjectError + 4100
Private Const ERROR_INVALID_OPERATION As Long = vbObjectError + 4101
Private Const ERROR_ARITY_MISMATCH As Long = vbObjectError + 4102
Private Const ERROR_TYPE_MISMATCH As Long = vbObjectError + 4103
Private Const ERROR_UNBOUND_PARAMETER As Long = vbObjectError + 4104
Private Const ERROR_METHOD_INVOCATION As Long = vbObjectError + 4105
Private Const ERROR_MEMBER_ACCESS As Long = vbObjectError + 4106
Private Const ERROR_NATIVE_INVOCATION As Long = vbObjectError + 4107
Private Const ERROR_BYREF_INVOCATION As Long = vbObjectError + 4108

Private mRole As Long
Private mOperator As Long
Private mExpressionOption As Long
Private mLeft As ROneCOne
Private mRight As ROneCOne
Private mNext As ROneCOne

Private mScalarValue As Variant
Private mObjectValue As Object
Private mHasObjectValue As Boolean

Private mExpectedType As VbVarType
Private mExpectedObjectType As String
Private mParameters As Collection
Private mIsByRef As Boolean

Private mTarget As Object
Private mMethodName As String
Private mMethodArity As Long
Private mIsActionDelegate As Boolean
Private mHasParameterSignature As Boolean
Private mHasReturnSignature As Boolean
Private mReturnType As VbVarType
Private mReturnObjectType As String
Private mInvocationList As Collection
Private mNativeAddress As LongPtr
Private mReferenceAddress As LongPtr

Private mTryBody As ROneCOne
Private mCatchActions As Collection
Private mCatchNumbers As Collection
Private mFinallyAction As ROneCOne
Private mErrorNumber As Long
Private mErrorSource As String
Private mErrorDescription As String
Private mErrorHelpFile As String
Private mErrorHelpContext As Long

Private mItems As Collection
Private mEnumerationValues As Collection
Private mElementType As VbVarType
Private mElementObjectType As String
Private mSource As ROneCOne
Private mQueryDelegate As ROneCOne
Private mQueryComparer As ROneCOne
Private mQueryValue As ROneCOne
Private mQueryCount As Long
Private mVersion As Long
Private mEnumerationVersion As Long
Private mElementParameter As ROneCOne
Private mContextMemberPath As String

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

Public Function Var( _
    Optional ByVal expectedType As VbVarType = vbVariant _
) As ROneCOne
Attribute Var.VB_Description = "Concise C#-style alias for a typed Parameter."
    Set Var = Parameter(expectedType)
End Function

Public Function VarLike(ByVal example As Variant) As ROneCOne
Attribute VarLike.VB_Description = "Creates a concise parameter typed like an example value."
    Set VarLike = ParameterLike(example)
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

Public Function ListOf( _
    ByVal typeToken As Variant, _
    ParamArray initialItems() As Variant _
) As ROneCOne
Attribute ListOf.VB_Description = "Creates a strict List<T> with optional initial items."
    Dim initialItem As Variant
    Dim listValue As ROneCOne
    Dim prototype As Object
    Dim wrappedItems As Collection

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
    Set wrappedItems = New Collection
    For Each initialItem In initialItems
        wrappedItems.Add listValue.InternalWrapElement(initialItem, "ListOf")
    Next initialItem
    CommitInitialItems listValue, wrappedItems
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

Public Function ListFrom( _
    ByVal firstItem As Variant, _
    ParamArray remainingItems() As Variant _
) As ROneCOne
Attribute ListFrom.VB_Description = "Infers T and creates a populated List<T>."
    Dim remainingItem As Variant
    Dim listValue As ROneCOne
    Dim wrappedItems As Collection

    Set listValue = ListLike(firstItem)
    Set wrappedItems = New Collection
    wrappedItems.Add listValue.InternalWrapElement(firstItem, "ListFrom")
    For Each remainingItem In remainingItems
        wrappedItems.Add listValue.InternalWrapElement(remainingItem, "ListFrom")
    Next remainingItem
    CommitInitialItems listValue, wrappedItems
    Set ListFrom = listValue
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

    If parameterList.Count = 0 Then
        CollectExpressionParameters body, parameterList
    End If

    ValidateBindings body, parameterList

    Set delegateValue = New ROneCOne
    delegateValue.ConfigureLambda body, parameterList
    Set Lambda = delegateValue
End Function

Public Function AsFunc() As ROneCOne
Attribute AsFunc.VB_Description = "Infers parameters and turns this expression into a Func."
    RequireExpression Me, "AsFunc"
    Set AsFunc = ROneCOne.Lambda(Me)
End Function

Public Function Func( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Func.VB_Description = "Creates a Func from an expression, object, or procedure."
    Set Func = CreateCallableDelegate(callable, memberName, False)
End Function

Public Function Action( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Action.VB_Description = "Creates an Action from an expression, object, or procedure."
    Set Action = CreateCallableDelegate(callable, memberName, True)
End Function

Public Function EventOf(ParamArray typeTokens() As Variant) As ROneCOne
Attribute EventOf.VB_Description = "Creates a strongly typed mutable event."
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim eventValue As ROneCOne
    Dim typeToken As Variant

    Set descriptors = New Collection
    For Each typeToken In typeTokens
        Set descriptor = CreateParameterDescriptor(typeToken, "EventOf")
        If descriptor.InternalIsByRef Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "EventOf", _
                "Events support ByVal parameters only."
        End If
        descriptors.Add descriptor
    Next typeToken
    Set eventValue = New ROneCOne
    eventValue.ConfigureEvent descriptors
    Set EventOf = eventValue
End Function

Public Function Try(ByVal work As ROneCOne) As ROneCOne
Attribute Try.VB_Description = "Begins a structured Try/Catch/Finally operation."
    Dim operation As ROneCOne

    If mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Try", _
            "Try must be called from the ROneCOne factory."
    End If
    ValidateAction work, "Try", True
    Set operation = New ROneCOne
    operation.ConfigureTry work
    Set Try = operation
End Function

Public Property Get Exception() As ROneCOne
Attribute Exception.VB_Description = "Gets the strongly typed exception prototype."
    Dim errorValue As ROneCOne

    Set errorValue = New ROneCOne
    errorValue.ConfigureError 0, vbNullString, vbNullString, vbNullString, 0
    Set Exception = errorValue
End Property

Public Function Native(ByVal procedureAddress As LongPtr) As ROneCOne
Attribute Native.VB_Description = "Creates a signature-bound native function delegate."
    Dim delegateValue As ROneCOne

    If procedureAddress = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Native", _
            "The procedure address cannot be zero."
    End If
    Set delegateValue = New ROneCOne
    delegateValue.ConfigureNative procedureAddress, False
    Set Native = delegateValue
End Function

Public Function NativeAction(ByVal procedureAddress As LongPtr) As ROneCOne
Attribute NativeAction.VB_Description = "Creates a signature-bound native Action delegate."
    Dim delegateValue As ROneCOne

    If procedureAddress = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "NativeAction", _
            "The procedure address cannot be zero."
    End If
    Set delegateValue = New ROneCOne
    delegateValue.ConfigureNative procedureAddress, True
    Set NativeAction = delegateValue
End Function

Public Function RefOf(ByVal typeToken As Variant) As ROneCOne
Attribute RefOf.VB_Description = "Marks a numeric delegate parameter as ByRef."
    Dim descriptor As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "RefOf"
    ValidateReferenceType valueType, "RefOf"
    Set descriptor = New ROneCOne
    descriptor.ConfigureParameter valueType, objectType, True
    Set RefOf = descriptor
End Function

Public Function RefLong(ByRef value As Long) As ROneCOne
Attribute RefLong.VB_Description = "Captures a Long variable for true ByRef invocation."
    Set RefLong = CreateReference(VarPtr(value), vbLong)
End Function

Public Function RefByte(ByRef value As Byte) As ROneCOne
Attribute RefByte.VB_Description = "Captures a Byte variable for true ByRef invocation."
    Set RefByte = CreateReference(VarPtr(value), vbByte)
End Function

Public Function RefInteger(ByRef value As Integer) As ROneCOne
Attribute RefInteger.VB_Description = "Captures an Integer for true ByRef invocation."
    Set RefInteger = CreateReference(VarPtr(value), vbInteger)
End Function

Public Function RefLongLong(ByRef value As LongLong) As ROneCOne
Attribute RefLongLong.VB_Description = "Captures a LongLong for true ByRef invocation."
    Set RefLongLong = CreateReference(VarPtr(value), vbLongLong)
End Function

Public Function RefSingle(ByRef value As Single) As ROneCOne
Attribute RefSingle.VB_Description = "Captures a Single variable for true ByRef invocation."
    Set RefSingle = CreateReference(VarPtr(value), vbSingle)
End Function

Public Function RefDouble(ByRef value As Double) As ROneCOne
Attribute RefDouble.VB_Description = "Captures a Double variable for true ByRef invocation."
    Set RefDouble = CreateReference(VarPtr(value), vbDouble)
End Function

Public Function RefCurrency(ByRef value As Currency) As ROneCOne
Attribute RefCurrency.VB_Description = "Captures a Currency value for true ByRef invocation."
    Set RefCurrency = CreateReference(VarPtr(value), vbCurrency)
End Function

Public Function RefDate(ByRef value As Date) As ROneCOne
Attribute RefDate.VB_Description = "Captures a Date variable for true ByRef invocation."
    Set RefDate = CreateReference(VarPtr(value), vbDate)
End Function

Public Function RefBoolean(ByRef value As Boolean) As ROneCOne
Attribute RefBoolean.VB_Description = "Captures a Boolean for true ByRef invocation."
    Set RefBoolean = CreateReference(VarPtr(value), vbBoolean)
End Function

Public Function Takes(ParamArray typeTokens() As Variant) As ROneCOne
Attribute Takes.VB_Description = "Returns a delegate with an immutable parameter signature."
    Dim clone As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim rawToken As Variant

    RequireDelegate Me, "Takes"
    Set descriptors = New Collection
    For Each rawToken In typeTokens
        Set descriptor = CreateParameterDescriptor(rawToken, "Takes")
        descriptors.Add descriptor
    Next rawToken

    Set clone = CloneDelegate(Me)
    clone.ConfigureParameterSignature descriptors
    Set Takes = clone
End Function

Public Function Returns(ByVal typeToken As Variant) As ROneCOne
Attribute Returns.VB_Description = "Returns a Func with an immutable result type."
    Dim clone As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    RequireDelegate Me, "Returns"
    If mIsActionDelegate Then
        RaiseContractError ERROR_INVALID_OPERATION, "Returns", _
            "An Action has no result; create a Func for a return value."
    End If
    ResolveTypeDescriptor typeToken, valueType, objectType, "Returns"
    Set clone = CloneDelegate(Me)
    clone.ConfigureReturnSignature valueType, objectType
    Set Returns = clone
End Function

Public Function DynamicInvoke(ByVal arguments As Variant) As Variant
Attribute DynamicInvoke.VB_Description = "Invokes from an array or sequence of arguments."
    Dim argumentValues As Collection
    Dim result As ROneCOne

    RequireDelegate Me, "DynamicInvoke"
    Set argumentValues = BuildDynamicArguments(arguments)
    Set result = InvokeValue(argumentValues)
    AssignPublicResult DynamicInvoke, result
End Function

Public Function Combine(ParamArray delegates() As Variant) As ROneCOne
Attribute Combine.VB_Description = "Returns an immutable multicast delegate."
    Dim combined As ROneCOne
    Dim invocationList As Collection
    Dim rawDelegate As Variant
    Dim delegateValue As ROneCOne

    Set invocationList = New Collection
    If IsDelegateRole(mRole) Then
        AppendInvocationList invocationList, Me
    ElseIf mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Combine", _
            "Combine must be called on a delegate or the ROneCOne factory."
    End If

    For Each rawDelegate In delegates
        Set delegateValue = RequireROneCOne(rawDelegate, "Combine")
        RequireDelegate delegateValue, "Combine"
        AppendInvocationList invocationList, delegateValue
    Next rawDelegate
    If invocationList.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Combine", _
            "At least one delegate is required."
    End If
    ValidateInvocationCompatibility invocationList

    If invocationList.Count = 1 Then
        Set Combine = invocationList.Item(1)
        Exit Function
    End If
    Set combined = New ROneCOne
    combined.ConfigureMulticast invocationList
    Set Combine = combined
End Function

Public Function Run(ParamArray arguments() As Variant) As Variant
Attribute Run.VB_Description = "Invokes, indexes, or selects a member from an expression."
Attribute Run.VB_UserMemId = 0
    Dim argumentValues As Collection
    Dim rawArgument As Variant

    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument

    AssignInvocationResult Run, argumentValues
End Function

Public Sub Execute(ParamArray arguments() As Variant)
Attribute Execute.VB_Description = "Executes an Action or structured operation."
    Dim argumentValues As Collection
    Dim ignored As ROneCOne
    Dim rawArgument As Variant

    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument

    If IsDelegateRole(mRole) Then
        If Not mIsActionDelegate Then
            RaiseContractError ERROR_INVALID_OPERATION, "Execute", _
                "Execute requires an Action delegate."
        End If
        Set ignored = InvokeValue(argumentValues)
        Exit Sub
    End If
    If mRole = ROLE_TRY Then
        If argumentValues.Count <> 0 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Execute", _
                "A Try operation does not accept arguments."
        End If
        ExecuteTryOperation
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Execute", _
        "Execute requires an Action or Try operation."
End Sub

Public Function Subscribe(ByVal handler As ROneCOne) As ROneCOne
Attribute Subscribe.VB_Description = "Adds an Action handler to this event."
    RequireEvent Me, "Subscribe"
    ValidateEventHandler handler, "Subscribe"
    mInvocationList.Add handler
    Set Subscribe = Me
End Function

Public Function Unsubscribe(ByVal handler As ROneCOne) As Boolean
Attribute Unsubscribe.VB_Description = "Removes the last matching event handler."
    Dim index As Long

    RequireEvent Me, "Unsubscribe"
    ValidateEventHandler handler, "Unsubscribe"
    For index = mInvocationList.Count To 1 Step -1
        If DelegateIdentityEquals(mInvocationList.Item(index), handler) Then
            mInvocationList.Remove index
            Unsubscribe = True
            Exit Function
        End If
    Next index
End Function

Public Sub Emit(ParamArray arguments() As Variant)
Attribute Emit.VB_Description = "Raises this event for a snapshot of its handlers."
    Dim argumentValues As Collection
    Dim handler As ROneCOne
    Dim handlers As Collection
    Dim ignored As ROneCOne
    Dim rawArgument As Variant

    RequireEvent Me, "Emit"
    Set argumentValues = New Collection
    For Each rawArgument In arguments
        argumentValues.Add WrapInvocationArgument(rawArgument)
    Next rawArgument
    ValidateDelegateArguments argumentValues
    Set handlers = CopyWrappedCollection(mInvocationList)
    For Each handler In handlers
        Set ignored = handler.InvokeValue(argumentValues)
    Next handler
End Sub

Public Property Get HandlerCount() As Long
Attribute HandlerCount.VB_Description = "Gets the number of subscribed event handlers."
    RequireEvent Me, "HandlerCount"
    HandlerCount = mInvocationList.Count
End Property

Public Function Catch( _
    ByVal handlerOrErrorNumber As Variant, _
    Optional ByVal handler As Variant _
) As ROneCOne
Attribute Catch.VB_Description = "Adds a catch-all or error-number Catch handler."
    Dim catchAction As ROneCOne
    Dim errorNumber As Long
    Dim operation As ROneCOne

    RequireTry Me, "Catch"
    If IsMissing(handler) Then
        Set catchAction = RequireROneCOne(handlerOrErrorNumber, "Catch")
    Else
        If Not IsNumeric(handlerOrErrorNumber) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Catch", _
                "A filtered Catch requires an error number."
        End If
        errorNumber = CLng(handlerOrErrorNumber)
        Set catchAction = RequireROneCOne(handler, "Catch")
    End If
    ValidateCatchAction catchAction
    Set operation = CloneTryOperation(Me)
    operation.ConfigureCatch errorNumber, catchAction
    Set Catch = operation
End Function

Public Function Finally(ByVal cleanup As ROneCOne) As ROneCOne
Attribute Finally.VB_Description = "Adds the Action that always runs after Try and Catch."
    Dim operation As ROneCOne

    RequireTry Me, "Finally"
    ValidateAction cleanup, "Finally", True
    If Not mFinallyAction Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Finally", _
            "A Try operation can have only one Finally block."
    End If
    Set operation = CloneTryOperation(Me)
    operation.ConfigureFinally cleanup
    Set Finally = operation
End Function

Private Sub AssignInvocationResult( _
    ByRef destination As Variant, _
    ByVal argumentValues As Collection _
)
    Dim argumentValue As ROneCOne
    Dim result As ROneCOne

    If mRole = ROLE_VALUE Or mRole = ROLE_PARAMETER Or _
        mRole = ROLE_EXPRESSION Then
        If argumentValues.Count = 1 Then
            Set argumentValue = argumentValues.Item(1)
            If Not argumentValue.InternalHasObjectValue Then
                If VarType(argumentValue.InternalScalarValue) = vbString Then
                    Set result = Member(CStr(argumentValue.InternalScalarValue))
                    Set destination = result
                    Exit Sub
                End If
            End If
        End If
    End If

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Then
        If argumentValues.Count <> 1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "A sequence indexer or member selector requires one argument."
        End If
        Set argumentValue = argumentValues.Item(1)
        If argumentValue.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "A sequence index must be numeric or a member name."
        End If
        If VarType(argumentValue.InternalScalarValue) = vbString Then
            Set result = Condition(CStr(argumentValue.InternalScalarValue))
            Set destination = result
            Exit Sub
        End If
        Set result = WrappedItem(CLng(argumentValue.InternalScalarValue))
        If result.InternalHasObjectValue Then
            Set destination = result.InternalObjectValue
        Else
            destination = result.InternalScalarValue
        End If
        Exit Sub
    End If

    Set result = InvokeValue(argumentValues)
    AssignPublicResult destination, result
End Sub

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
    If mRole = ROLE_WHERE_BUILDER Then
        Set EqualTo = CompleteWhereBinary(OP_EQUAL, operand)
    Else
        Set EqualTo = CreateBinaryExpression(OP_EQUAL, operand)
    End If
End Function

Public Function NotEqualTo(ByVal operand As Variant) As ROneCOne
Attribute NotEqualTo.VB_Description = "Creates an inequality expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set NotEqualTo = CompleteWhereBinary(OP_NOT_EQUAL, operand)
    Else
        Set NotEqualTo = CreateBinaryExpression(OP_NOT_EQUAL, operand)
    End If
End Function

Public Function EqualToIgnoreCase(ByVal operand As String) As ROneCOne
Attribute EqualToIgnoreCase.VB_Description = "Tests string equality without case sensitivity."
    Set EqualToIgnoreCase = CompleteStringCondition( _
        OP_STRING_EQUAL, operand, vbTextCompare)
End Function

Public Function NotEqualToIgnoreCase(ByVal operand As String) As ROneCOne
Attribute NotEqualToIgnoreCase.VB_Description = "Tests string inequality without case sensitivity."
    Set NotEqualToIgnoreCase = CompleteStringCondition( _
        OP_STRING_NOT_EQUAL, operand, vbTextCompare)
End Function

Public Function LessThan(ByVal operand As Variant) As ROneCOne
Attribute LessThan.VB_Description = "Creates a less-than expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set LessThan = CompleteWhereBinary(OP_LESS_THAN, operand)
    Else
        Set LessThan = CreateBinaryExpression(OP_LESS_THAN, operand)
    End If
End Function

Public Function LessThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute LessThanOrEqual.VB_Description = "Creates a less-than-or-equal expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set LessThanOrEqual = CompleteWhereBinary( _
            OP_LESS_THAN_OR_EQUAL, operand)
    Else
        Set LessThanOrEqual = CreateBinaryExpression( _
            OP_LESS_THAN_OR_EQUAL, operand)
    End If
End Function

Public Function GreaterThan(ByVal operand As Variant) As ROneCOne
Attribute GreaterThan.VB_Description = "Creates a greater-than expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set GreaterThan = CompleteWhereBinary(OP_GREATER_THAN, operand)
    Else
        Set GreaterThan = CreateBinaryExpression(OP_GREATER_THAN, operand)
    End If
End Function

Public Function GreaterThanOrEqual(ByVal operand As Variant) As ROneCOne
Attribute GreaterThanOrEqual.VB_Description = "Creates a greater-than-or-equal expression."
    If mRole = ROLE_WHERE_BUILDER Then
        Set GreaterThanOrEqual = CompleteWhereBinary( _
            OP_GREATER_THAN_OR_EQUAL, operand)
    Else
        Set GreaterThanOrEqual = CreateBinaryExpression( _
            OP_GREATER_THAN_OR_EQUAL, operand)
    End If
End Function

Public Function AtLeast(ByVal operand As Variant) As ROneCOne
Attribute AtLeast.VB_Description = "Sugar for a greater-than-or-equal expression."
    Set AtLeast = GreaterThanOrEqual(operand)
End Function

Public Function AtMost(ByVal operand As Variant) As ROneCOne
Attribute AtMost.VB_Description = "Sugar for a less-than-or-equal expression."
    Set AtMost = LessThanOrEqual(operand)
End Function

Public Function Between( _
    ByVal lowerBound As Variant, _
    ByVal upperBound As Variant, _
    Optional ByVal inclusive As Boolean = True _
) As ROneCOne
Attribute Between.VB_Description = "Tests whether a value lies between two bounds."
    Dim receiver As ROneCOne
    Dim expression As ROneCOne

    If mRole = ROLE_WHERE_BUILDER Then
        Set receiver = ContextMemberExpression(False)
    Else
        RequireExpression Me, "Between receiver"
        Set receiver = Me
    End If
    If inclusive Then
        Set expression = receiver.GreaterThanOrEqual(lowerBound) _
            .AndAlso(receiver.LessThanOrEqual(upperBound))
    Else
        Set expression = receiver.GreaterThan(lowerBound) _
            .AndAlso(receiver.LessThan(upperBound))
    End If
    Set Between = CompleteWhereOrExpression(expression)
End Function

Public Function OneOf(ParamArray values() As Variant) As ROneCOne
Attribute OneOf.VB_Description = "Tests whether a value equals any supplied candidate."
    Dim candidate As Variant
    Dim expression As ROneCOne
    Dim hasSingleValue As Boolean
    Dim receiver As ROneCOne

    On Error Resume Next
    hasSingleValue = (UBound(values) = LBound(values))
    On Error GoTo 0
    If hasSingleValue Then
        If IsSequenceContainer(values(LBound(values))) Then
            Set OneOf = CreateMembershipCondition( _
                values(LBound(values)), False)
            Exit Function
        End If
    End If
    If mRole = ROLE_WHERE_BUILDER Then
        Set receiver = ContextMemberExpression(False)
    Else
        RequireExpression Me, "OneOf receiver"
        Set receiver = Me
    End If
    For Each candidate In values
        If expression Is Nothing Then
            Set expression = receiver.EqualTo(candidate)
        Else
            Set expression = expression.OrElse(receiver.EqualTo(candidate))
        End If
    Next candidate
    If expression Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "OneOf", _
            "At least one candidate value is required."
    End If
    Set OneOf = CompleteWhereOrExpression(expression)
End Function

Public Function IsIn( _
    ByVal values As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute IsIn.VB_Description = "Tests membership in a sequence, array, or Collection."
    If IsMissing(comparer) Then
        Set IsIn = CreateMembershipCondition(values, False)
    Else
        Set IsIn = CreateMembershipCondition(values, False, comparer)
    End If
End Function

Public Function NotIn( _
    ByVal values As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute NotIn.VB_Description = "Tests non-membership in a sequence, array, or Collection."
    If IsMissing(comparer) Then
        Set NotIn = CreateMembershipCondition(values, True)
    Else
        Set NotIn = CreateMembershipCondition(values, True, comparer)
    End If
End Function

Public Function StartsWith( _
    ByVal prefix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute StartsWith.VB_Description = "Tests a string prefix using an explicit comparison mode."
    Set StartsWith = CompleteStringCondition( _
        OP_STARTS_WITH, prefix, compareMode)
End Function

Public Function EndsWith( _
    ByVal suffix As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute EndsWith.VB_Description = "Tests a string suffix using an explicit comparison mode."
    Set EndsWith = CompleteStringCondition( _
        OP_ENDS_WITH, suffix, compareMode)
End Function

Public Function ContainsText( _
    ByVal fragment As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute ContainsText.VB_Description = "Tests whether a string contains a fragment."
    Set ContainsText = CompleteStringCondition( _
        OP_CONTAINS_TEXT, fragment, compareMode)
End Function

Public Function StartsWithIgnoreCase(ByVal prefix As String) As ROneCOne
Attribute StartsWithIgnoreCase.VB_Description = "Tests a string prefix without case sensitivity."
    Set StartsWithIgnoreCase = StartsWith(prefix, vbTextCompare)
End Function

Public Function EndsWithIgnoreCase(ByVal suffix As String) As ROneCOne
Attribute EndsWithIgnoreCase.VB_Description = "Tests a string suffix without case sensitivity."
    Set EndsWithIgnoreCase = EndsWith(suffix, vbTextCompare)
End Function

Public Function ContainsIgnoreCase(ByVal fragment As String) As ROneCOne
Attribute ContainsIgnoreCase.VB_Description = "Tests string containment without case sensitivity."
    Set ContainsIgnoreCase = ContainsText(fragment, vbTextCompare)
End Function

Public Function MatchesPattern( _
    ByVal pattern As String, _
    Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare _
) As ROneCOne
Attribute MatchesPattern.VB_Description = "Tests a value with VBA's Like pattern syntax."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    Set receiver = ConditionReceiver(False, "MatchesPattern")
    Set rightValue = CoerceExpression(pattern)
    Set expression = New ROneCOne
    ValidateStringCompareMode compareMode, "MatchesPattern"
    expression.ConfigureExpression OP_MATCHES_PATTERN, receiver, _
        rightValue, compareMode
    Set MatchesPattern = CompleteWhereOrExpression(expression)
End Function

Public Function MatchesPatternIgnoreCase(ByVal pattern As String) As ROneCOne
Attribute MatchesPatternIgnoreCase.VB_Description = "Tests a Like pattern without case sensitivity."
    Set MatchesPatternIgnoreCase = MatchesPattern(pattern, vbTextCompare)
End Function

Public Function IsNothing() As ROneCOne
Attribute IsNothing.VB_Description = "Tests whether an object reference is Nothing."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "IsNothing")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NOTHING, receiver, Nothing
    Set IsNothing = CompleteWhereOrExpression(expression)
End Function

Public Function IsNotNothing() As ROneCOne
Attribute IsNotNothing.VB_Description = "Tests whether an object reference is not Nothing."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "IsNotNothing")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NOT_NOTHING, receiver, Nothing
    Set IsNotNothing = CompleteWhereOrExpression(expression)
End Function

Public Function IsNullOrEmpty() As ROneCOne
Attribute IsNullOrEmpty.VB_Description = "Tests whether a scalar is Null or an empty string."
    Dim expression As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(False, "IsNullOrEmpty")
    Set expression = New ROneCOne
    expression.ConfigureExpression OP_IS_NULL_OR_EMPTY, receiver, Nothing
    Set IsNullOrEmpty = CompleteWhereOrExpression(expression)
End Function

Public Function IsTrue() As ROneCOne
Attribute IsTrue.VB_Description = "Tests whether a Boolean value is True."
    If mRole = ROLE_WHERE_BUILDER Then
        Set IsTrue = CompleteWhereBinary(OP_EQUAL, True)
    Else
        Set IsTrue = EqualTo(True)
    End If
End Function

Public Function IsFalse() As ROneCOne
Attribute IsFalse.VB_Description = "Tests whether a Boolean value is False."
    If mRole = ROLE_WHERE_BUILDER Then
        Set IsFalse = CompleteWhereBinary(OP_EQUAL, False)
    Else
        Set IsFalse = EqualTo(False)
    End If
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

Public Function Member( _
    ByVal memberName As String, _
    Optional ByVal returnsObject As Boolean = False, _
    Optional ByVal nullSafe As Boolean = False _
) As ROneCOne
Attribute Member.VB_Description = "Creates an object property-access expression."
    Dim expression As ROneCOne
    Dim operatorCode As Long

    RequireExpression Me, "Member receiver"
    memberName = Trim$(memberName)
    If Len(memberName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Member", _
            "The member name cannot be empty."
    End If

    If nullSafe Then
        operatorCode = OP_SAFE_MEMBER_VALUE
        If returnsObject Then operatorCode = OP_SAFE_MEMBER_OBJECT
    Else
        operatorCode = OP_MEMBER_VALUE
        If returnsObject Then operatorCode = OP_MEMBER_OBJECT
    End If
    Set expression = New ROneCOne
    expression.ConfigureMemberExpression operatorCode, Me, memberName
    Set Member = expression
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

Public Property Get Count(Optional ByVal predicate As Variant) As Long
Attribute Count.VB_Description = "Counts every element or only predicate matches."
    Dim predicateValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If IsMissing(predicate) Then
        Count = values.Count
        Exit Property
    End If
    Set predicateValue = NormalizeContextUnary(predicate, "Count", False)
    For Each wrapped In values
        If InvokePredicate(predicateValue, wrapped, "Count") Then Count = Count + 1
    Next wrapped
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

Public Sub AddRange(ByVal values As Variant)
Attribute AddRange.VB_Description = "Atomically appends a sequence, array, or Collection."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireList Me, "AddRange"
    Set incoming = BuildIncomingItems(values, "AddRange")
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

Public Function Contains( _
    ByVal value As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Contains.VB_Description = "Tests sequence membership or string containment."
    Dim expression As ROneCOne

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Then
        If IsExpressionVariant(value) Then
            Set expression = RequireROneCOne(value, "Contains")
            If IsMissing(comparer) Then
                Set Contains = expression.IsIn(Me)
            Else
                Set Contains = expression.IsIn(Me, comparer)
            End If
        ElseIf IsMissing(comparer) Then
            Contains = (IndexOf(value) >= 0)
        Else
            Contains = (IndexOf(value, comparer) >= 0)
        End If
        Exit Function
    End If
    Set Contains = ContainsText(CStr(value))
End Function

Public Function Both(ByVal operand As Variant) As ROneCOne
Attribute Both.VB_Description = "Concise alias for short-circuit Boolean AndAlso."
    Set Both = AndAlso(operand)
End Function

Public Function Either(ByVal operand As Variant) As ROneCOne
Attribute Either.VB_Description = "Concise alias for short-circuit Boolean OrElse."
    Set Either = OrElse(operand)
End Function

Public Function Negated() As ROneCOne
Attribute Negated.VB_Description = "Concise alias for a Boolean NOT expression."
    Set Negated = NotExpression
End Function

Public Function IndexOf( _
    ByVal value As Variant, _
    Optional ByVal comparer As Variant _
) As Long
Attribute IndexOf.VB_Description = "Gets the first zero-based index of a value, or -1."
    Dim candidate As ROneCOne
    Dim comparerValue As ROneCOne
    Dim index As Long
    Dim values As Collection

    RequireSequence Me, "IndexOf"
    Set candidate = WrapElement(value, "IndexOf")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "IndexOf")
    End If
    Set values = MaterializeWrapped
    IndexOf = -1
    For index = 1 To values.Count
        If ValuesEqualWithComparer( _
            values.Item(index), candidate, comparerValue, "IndexOf") Then
            IndexOf = index - 1
            Exit Function
        End If
    Next index
End Function

Public Function ContainsMember( _
    ByVal source As ROneCOne, _
    ByVal memberPath As String _
) As ROneCOne
Attribute ContainsMember.VB_Description = "Builds a reusable collection-membership expression."
    RequireSequence Me, "ContainsMember"
    RequireSequence source, "ContainsMember"
    Set ContainsMember = source.Condition(memberPath).IsIn(Me)
End Function

Public Function SequenceEqual( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As Boolean
Attribute SequenceEqual.VB_Description = "Tests ordered equality with an optional comparer."
    Dim comparerValue As ROneCOne
    Dim index As Long
    Dim leftValues As Collection
    Dim rightSequence As ROneCOne
    Dim rightValues As Collection

    Set rightSequence = NormalizeMembershipSequence(other, "SequenceEqual")
    Set leftValues = MaterializeWrapped
    Set rightValues = rightSequence.MaterializeWrapped
    If leftValues.Count <> rightValues.Count Then Exit Function
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "SequenceEqual")
    End If
    For index = 1 To leftValues.Count
        If Not ValuesEqualWithComparer( _
            leftValues.Item(index), rightValues.Item(index), _
            comparerValue, "SequenceEqual") Then Exit Function
    Next index
    SequenceEqual = True
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

Public Function Remove( _
    ByVal value As Variant, _
    Optional ByVal removal As Variant _
) As Variant
Attribute Remove.VB_Description = "Removes a list item or delegate invocation."
    Dim delegateValue As ROneCOne
    Dim result As ROneCOne

    If mRole = ROLE_LIST Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "List removal accepts exactly one value."
        End If
        Remove = RemoveListItem(value)
        Exit Function
    End If

    If IsDelegateRole(mRole) Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Instance delegate removal accepts one delegate."
        End If
        Set delegateValue = RequireROneCOne(value, "Remove")
        Set result = RemoveDelegateValue(Me, delegateValue)
    ElseIf mRole = ROLE_EMPTY Then
        If IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Static delegate removal requires source and value delegates."
        End If
        Set delegateValue = RequireROneCOne(removal, "Remove")
        Set result = RemoveDelegateValue(RequireROneCOne(value, "Remove"), delegateValue)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Remove", _
            "Remove requires a list or delegate."
    End If

    Set Remove = result
End Function

Public Sub RemoveAt(ByVal index As Long)
Attribute RemoveAt.VB_Description = "Removes the item at a zero-based list index."
    RequireList Me, "RemoveAt"
    ValidateIndex index, mItems.Count, "RemoveAt"
    mItems.Remove index + 1
    mEnumerationValues.Remove index + 1
    mVersion = mVersion + 1
End Sub

Public Property Get Element() As ROneCOne
Attribute Element.VB_Description = "Gets a typed parameter representing each sequence element."
    RequireSequence Me, "Element"
    If mElementParameter Is Nothing Then
        Set mElementParameter = New ROneCOne
        mElementParameter.ConfigureParameter mElementType, mElementObjectType
    End If
    Set Element = mElementParameter
End Property

Public Function Condition( _
    ByVal memberPath As String, _
    Optional ByVal returnsObject As Boolean = False _
) As ROneCOne
Attribute Condition.VB_Description = "Creates a reusable sequence-member expression."
    RequireSequence Me, "Condition"
    Set Condition = BuildMemberPath(memberPath, returnsObject)
End Function

Public Function Match( _
    ByVal memberPath As String, _
    ByVal value As Variant _
) As ROneCOne
Attribute Match.VB_Description = "Creates a reusable member-equality predicate expression."
    Set Match = Condition(memberPath).EqualTo(value)
End Function

Public Function NotMatch( _
    ByVal memberPath As String, _
    ByVal value As Variant _
) As ROneCOne
Attribute NotMatch.VB_Description = "Creates a reusable member-inequality predicate expression."
    Set NotMatch = Condition(memberPath).NotEqualTo(value)
End Function

Public Function Always() As ROneCOne
Attribute Always.VB_Description = "Creates a reusable predicate that always returns True."
    Set Always = ROneCOne.Lambda(ROneCOne.Value(True), Element)
End Function

Public Function Never() As ROneCOne
Attribute Never.VB_Description = "Creates a reusable predicate that always returns False."
    Set Never = ROneCOne.Lambda(ROneCOne.Value(False), Element)
End Function

Public Function EqualityComparer( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute EqualityComparer.VB_Description = "Creates a typed Func<T, T, Boolean> equality comparer."
    Set EqualityComparer = CreateSequenceComparer( _
        callable, memberName, True, "EqualityComparer")
End Function

Public Function Comparer( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Comparer.VB_Description = "Creates a typed Func<T, T, Long> ordering comparer."
    Set Comparer = CreateSequenceComparer( _
        callable, memberName, False, "Comparer")
End Function

Public Function Predicate( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute Predicate.VB_Description = "Infers a typed Func<T, Boolean> for this sequence."
    Dim predicateDelegate As ROneCOne

    RequireSequence Me, "Predicate"
    Set predicateDelegate = ROneCOne.Func(callable, memberName)
    Set Predicate = BindContextPredicate(predicateDelegate)
End Function

Public Function WhereMethod( _
    ByVal callable As Variant, _
    Optional ByVal memberName As String = vbNullString _
) As ROneCOne
Attribute WhereMethod.VB_Description = "Filters with an inferred method or procedure predicate."
    Set WhereMethod = Where(Predicate(callable, memberName))
End Function

Public Function Where(ByVal predicateOrMember As Variant) As ROneCOne
Attribute Where.VB_Description = "Defers filtering or begins a fluent member condition."
    Dim builder As ROneCOne
    Dim predicateDelegate As ROneCOne

    RequireSequence Me, "Where"
    If Not IsObject(predicateOrMember) Then
        If VarType(predicateOrMember) = vbString Then
            Set builder = New ROneCOne
            builder.ConfigureWhereBuilder Me, CStr(predicateOrMember)
            Set Where = builder
            Exit Function
        End If
    End If
    Set predicateDelegate = NormalizeUnary( _
        RequireROneCOne(predicateOrMember, "Where"), "Where")
    Set Where = CreateQuery(QUERY_WHERE, predicateDelegate, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function WhereNot(ByVal predicate As Variant) As ROneCOne
Attribute WhereNot.VB_Description = "Filters out elements matching a predicate."
    Dim predicateValue As ROneCOne

    Set predicateValue = NormalizeContextUnary(predicate, "WhereNot", False)
    Set WhereNot = CreateQuery( _
        QUERY_WHERE_NOT, predicateValue, Nothing, 0, _
        mElementType, mElementObjectType)
End Function

Public Function WhereAny( _
    ByVal collectionPath As String, _
    Optional ByVal predicate As Variant _
) As ROneCOne
Attribute WhereAny.VB_Description = "Filters when a nested sequence has any matching element."
    Dim expression As ROneCOne

    Set expression = Condition(collectionPath, True)
    If IsMissing(predicate) Then
        Set WhereAny = Where(expression.AnyMatch)
    Else
        Set WhereAny = Where(expression.AnyMatch(predicate))
    End If
End Function

Public Function WhereAll( _
    ByVal collectionPath As String, _
    ByVal predicate As Variant _
) As ROneCOne
Attribute WhereAll.VB_Description = "Filters when all nested-sequence elements match."
    Set WhereAll = Where( _
        Condition(collectionPath, True).AllMatch(predicate))
End Function

Public Function WhereNone( _
    ByVal collectionPath As String, _
    Optional ByVal predicate As Variant _
) As ROneCOne
Attribute WhereNone.VB_Description = "Filters when no nested-sequence element matches."
    Dim expression As ROneCOne

    Set expression = Condition(collectionPath, True)
    If IsMissing(predicate) Then
        Set WhereNone = Where(expression.NoneMatch)
    Else
        Set WhereNone = Where(expression.NoneMatch(predicate))
    End If
End Function

Public Function SelectItems( _
    ByVal selector As Variant, _
    ByVal resultTypeToken As Variant _
) As ROneCOne
Attribute SelectItems.VB_Description = "Defers a strongly typed LINQ Select projection."
    Dim objectType As String
    Dim selectorDelegate As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor resultTypeToken, valueType, objectType, "SelectItems"
    Set selectorDelegate = NormalizeContextUnary( _
        selector, "SelectItems", valueType = vbObject)
    Set SelectItems = CreateQuery(QUERY_SELECT, selectorDelegate, Nothing, 0, _
        valueType, objectType)
End Function

Public Function Map( _
    ByVal selector As Variant, _
    ByVal resultTypeToken As Variant _
) As ROneCOne
Attribute Map.VB_Description = "Concise sugar for a strongly typed SelectItems projection."
    Set Map = SelectItems(selector, resultTypeToken)
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

Public Function Distinct(Optional ByVal comparer As Variant) As ROneCOne
Attribute Distinct.VB_Description = "Defers removal of duplicate sequence values."
    Dim comparerValue As ROneCOne

    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "Distinct")
    End If
    Set Distinct = CreateQuery(QUERY_DISTINCT, Nothing, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function DistinctBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute DistinctBy.VB_Description = "Defers removal of elements with duplicate selected keys."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "DistinctBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "DistinctBy")
    End If
    Set DistinctBy = CreateQuery(QUERY_DISTINCT_BY, selectorDelegate, _
        Nothing, 0, mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderBy.VB_Description = "Defers stable ascending ordering by a selected key."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "OrderBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "OrderBy")
    End If
    Set OrderBy = CreateQuery(QUERY_ORDER_BY, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderByDescending( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderByDescending.VB_Description = "Defers stable descending ordering by a selected key."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "OrderByDescending", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, False, "OrderByDescending")
    End If
    Set OrderByDescending = CreateQuery( _
        QUERY_ORDER_BY_DESCENDING, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function Sorted(Optional ByVal comparer As Variant) As ROneCOne
Attribute Sorted.VB_Description = "Orders a sequence ascending by each element itself."
    If IsMissing(comparer) Then
        Set Sorted = OrderBy(Element)
    Else
        Set Sorted = OrderBy(Element, comparer)
    End If
End Function

Public Function SortedDescending(Optional ByVal comparer As Variant) As ROneCOne
Attribute SortedDescending.VB_Description = "Orders a sequence descending by each element itself."
    If IsMissing(comparer) Then
        Set SortedDescending = OrderByDescending(Element)
    Else
        Set SortedDescending = OrderByDescending(Element, comparer)
    End If
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

Public Function AnyMatch(Optional ByVal predicate As Variant) As ROneCOne
Attribute AnyMatch.VB_Description = "Tests whether a nested sequence has any matching element."
    If IsMissing(predicate) Then
        Set AnyMatch = CreateQuantifierCondition(OP_ANY_MATCH, Empty, False)
    Else
        Set AnyMatch = CreateQuantifierCondition( _
            OP_ANY_MATCH, predicate, True)
    End If
End Function

Public Function AllMatch(ByVal predicate As Variant) As ROneCOne
Attribute AllMatch.VB_Description = "Tests whether all nested-sequence elements match."
    Set AllMatch = CreateQuantifierCondition(OP_ALL_MATCH, predicate, True)
End Function

Public Function NoneMatch(Optional ByVal predicate As Variant) As ROneCOne
Attribute NoneMatch.VB_Description = "Tests whether no nested-sequence element matches."
    If IsMissing(predicate) Then
        Set NoneMatch = CreateQuantifierCondition(OP_NONE_MATCH, Empty, False)
    Else
        Set NoneMatch = CreateQuantifierCondition( _
            OP_NONE_MATCH, predicate, True)
    End If
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
    Set predicateValue = NormalizeContextUnary( _
        predicate, "Any", False)
    For Each wrapped In values
        If InvokePredicate(predicateValue, wrapped, "Any") Then
            AnyItem = True
            Exit Function
        End If
    Next wrapped
End Function

Public Function Exists(Optional ByVal predicate As Variant) As Boolean
Attribute Exists.VB_Description = "Tests whether any element exists or matches."
    If IsMissing(predicate) Then
        Exists = AnyItem
    Else
        Exists = AnyItem(predicate)
    End If
End Function

Public Function None(Optional ByVal predicate As Variant) As Boolean
Attribute None.VB_Description = "Tests whether no element exists or matches a predicate."
    If IsMissing(predicate) Then
        None = Not AnyItem
    Else
        None = Not AnyItem(predicate)
    End If
End Function

Public Function All(ByVal predicate As Variant) As Boolean
Attribute All.VB_Description = "Tests whether every element matches a predicate."
    Dim predicateDelegate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set predicateDelegate = NormalizeContextUnary(predicate, "All", False)
    Set values = MaterializeWrapped
    All = True
    For Each wrapped In values
        If Not InvokePredicate(predicateDelegate, wrapped, "All") Then
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
        Set predicateValue = NormalizeContextUnary( _
            predicate, "First", False)
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
        Set predicateValue = NormalizeContextUnary( _
            predicate, "Last", False)
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
        Set selectorValue = NormalizeContextUnary( _
            selector, "Sum", False)
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
        Set selectorValue = NormalizeContextUnary( _
            selector, "Average", False)
    End If
    For Each wrapped In values
        Set selected = SelectTerminalValue(wrapped, selectorValue)
        total = total + CDbl(RequireScalar(selected, "Average"))
    Next wrapped
    Average = total / CDbl(values.Count)
End Function

Public Function Min( _
    Optional ByVal selector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Min.VB_Description = "Returns the minimum element or selected value."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Min", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "Min")
    End If
    Set wrapped = ExtremeValue(selectorValue, comparerValue, False, "Min")
    If wrapped.InternalHasObjectValue Then
        Set Min = wrapped.InternalObjectValue
    Else
        Min = wrapped.InternalScalarValue
    End If
End Function

Public Function Max( _
    Optional ByVal selector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute Max.VB_Description = "Returns the maximum element or selected value."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(selector) Then
        Set selectorValue = NormalizeContextUnary( _
            selector, "Max", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "Max")
    End If
    Set wrapped = ExtremeValue(selectorValue, comparerValue, True, "Max")
    If wrapped.InternalHasObjectValue Then
        Set Max = wrapped.InternalObjectValue
    Else
        Max = wrapped.InternalScalarValue
    End If
End Function

Public Function FirstOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute FirstOrDefault.VB_Description = "Returns the first match or default(T)."
    Dim found As Boolean
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "FirstOrDefault", False)
    End If
    Set wrapped = TryFindTerminalElement( _
        predicateValue, False, "FirstOrDefault", found)
    If found Then
        AssignPublicResult FirstOrDefault, wrapped
    Else
        AssignDefaultPublicResult FirstOrDefault
    End If
End Function

Public Function LastOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute LastOrDefault.VB_Description = "Returns the last match or default(T)."
    Dim found As Boolean
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "LastOrDefault", False)
    End If
    Set wrapped = TryFindTerminalElement( _
        predicateValue, True, "LastOrDefault", found)
    If found Then
        AssignPublicResult LastOrDefault, wrapped
    Else
        AssignDefaultPublicResult LastOrDefault
    End If
End Function

Public Function SingleItem(Optional ByVal predicate As Variant) As Variant
Attribute SingleItem.VB_Description = "Returns the sole matching element or raises."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary(predicate, "Single", False)
    End If
    Set wrapped = FindSingleElement(predicateValue, False, "Single")
    AssignPublicResult SingleItem, wrapped
End Function

Public Function SingleOrDefault(Optional ByVal predicate As Variant) As Variant
Attribute SingleOrDefault.VB_Description = "Returns one match, default(T), or raises for many."
    Dim predicateValue As ROneCOne
    Dim wrapped As ROneCOne

    If Not IsMissing(predicate) Then
        Set predicateValue = NormalizeContextUnary( _
            predicate, "SingleOrDefault", False)
    End If
    Set wrapped = FindSingleElement( _
        predicateValue, True, "SingleOrDefault")
    If wrapped Is Nothing Then
        AssignDefaultPublicResult SingleOrDefault
    Else
        AssignPublicResult SingleOrDefault, wrapped
    End If
End Function

Public Function MinBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute MinBy.VB_Description = "Returns the element with the minimum selected key."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "MinBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "MinBy")
    End If
    Set wrapped = ExtremeElement( _
        selectorValue, comparerValue, False, "MinBy")
    AssignPublicResult MinBy, wrapped
End Function

Public Function MaxBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As Variant
Attribute MaxBy.VB_Description = "Returns the element with the maximum selected key."
    Dim comparerValue As ROneCOne
    Dim selectorValue As ROneCOne
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "MaxBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "MaxBy")
    End If
    Set wrapped = ExtremeElement( _
        selectorValue, comparerValue, True, "MaxBy")
    AssignPublicResult MaxBy, wrapped
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
        Case ROLE_METHOD, ROLE_PROCEDURE, ROLE_NATIVE
            Arity = mMethodArity
        Case ROLE_COMPOSITION
            Arity = mLeft.Arity
        Case ROLE_MULTICAST
            Arity = mInvocationList.Item(1).Arity
        Case Else
            Arity = -1
    End Select
End Property

Public Property Get InvocationCount() As Long
Attribute InvocationCount.VB_Description = "Gets the immutable delegate invocation count."
    RequireDelegate Me, "InvocationCount"
    If mRole = ROLE_MULTICAST Then
        InvocationCount = mInvocationList.Count
    Else
        InvocationCount = 1
    End If
End Property

Public Property Get Target() As Object
Attribute Target.VB_Description = "Gets the bound object target, or Nothing."
    RequireDelegate Me, "Target"
    Set Target = mTarget
End Property

Public Property Get MethodName() As String
Attribute MethodName.VB_Description = "Gets the bound method or workbook procedure name."
    RequireDelegate Me, "MethodName"
    MethodName = mMethodName
End Property

Public Property Get IsAction() As Boolean
Attribute IsAction.VB_Description = "Gets whether the delegate has no result value."
    RequireDelegate Me, "IsAction"
    IsAction = mIsActionDelegate
End Property

Public Property Get Signature() As String
Attribute Signature.VB_Description = "Gets the C#-style delegate signature description."
    RequireDelegate Me, "Signature"
    Signature = BuildSignature(Me)
End Property

Public Function GetInvocationList() As ROneCOne
Attribute GetInvocationList.VB_Description = "Returns a typed copy of the invocation list."
    Dim delegateValue As ROneCOne
    Dim result As ROneCOne

    RequireDelegate Me, "GetInvocationList"
    Set result = ROneCOne.ListOf(Me)
    If mRole = ROLE_MULTICAST Then
        For Each delegateValue In mInvocationList
            result.Add delegateValue
        Next delegateValue
    Else
        result.Add Me
    End If
    Set GetInvocationList = result
End Function

Public Sub ForEach(ByVal action As ROneCOne)
Attribute ForEach.VB_Description = "Executes an Action once for every sequence element."
    Dim actionDelegate As ROneCOne
    Dim ignored As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set actionDelegate = NormalizeUnaryAction(action, "ForEach")
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set ignored = InvokeUnary(actionDelegate, wrapped)
    Next wrapped
End Sub

Public Function JoinText( _
    ByVal separator As String, _
    Optional ByVal selector As Variant _
) As String
Attribute JoinText.VB_Description = "Joins scalar sequence values into one String."
    Dim index As Long
    Dim parts() As String
    Dim selected As ROneCOne
    Dim selectorDelegate As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence Me, "JoinText"
    If Not IsMissing(selector) Then
        Set selectorDelegate = NormalizeContextUnary( _
            selector, "JoinText", False)
    End If
    Set values = MaterializeWrapped
    If values.Count = 0 Then Exit Function

    ReDim parts(0 To values.Count - 1)
    For Each wrapped In values
        If selectorDelegate Is Nothing Then
            Set selected = wrapped
        Else
            Set selected = InvokeUnary(selectorDelegate, wrapped)
        End If
        parts(index) = CStr(RequireScalar(selected, "JoinText"))
        index = index + 1
    Next wrapped
    JoinText = VBA.Join(parts, separator)
End Function

Public Property Get TypeMismatchError() As Long
Attribute TypeMismatchError.VB_Description = "Gets the ROneCOne type-mismatch error number."
    TypeMismatchError = ERROR_TYPE_MISMATCH
End Property

Public Property Get UnboundParameterError() As Long
Attribute UnboundParameterError.VB_Description = "Gets the unbound-parameter error number."
    UnboundParameterError = ERROR_UNBOUND_PARAMETER
End Property

Public Property Get MemberAccessError() As Long
Attribute MemberAccessError.VB_Description = "Gets the object-member access error number."
    MemberAccessError = ERROR_MEMBER_ACCESS
End Property

Public Property Get NativeInvocationError() As Long
Attribute NativeInvocationError.VB_Description = "Gets the native invocation error number."
    NativeInvocationError = ERROR_NATIVE_INVOCATION
End Property

Public Property Get ByRefInvocationError() As Long
Attribute ByRefInvocationError.VB_Description = "Gets the unsupported ByRef boundary error."
    ByRefInvocationError = ERROR_BYREF_INVOCATION
End Property

Public Property Get InvalidOperationError() As Long
Attribute InvalidOperationError.VB_Description = "Gets the invalid delegate operation error."
    InvalidOperationError = ERROR_INVALID_OPERATION
End Property

Public Property Get ErrorNumber() As Long
Attribute ErrorNumber.VB_Description = "Gets the captured VBA error number."
    RequireError Me, "ErrorNumber"
    ErrorNumber = mErrorNumber
End Property

Public Property Get ErrorSource() As String
Attribute ErrorSource.VB_Description = "Gets the captured VBA error source."
    RequireError Me, "ErrorSource"
    ErrorSource = mErrorSource
End Property

Public Property Get Message() As String
Attribute Message.VB_Description = "Gets the captured VBA error description."
    RequireError Me, "Message"
    Message = mErrorDescription
End Property

Public Property Get HelpFile() As String
Attribute HelpFile.VB_Description = "Gets the captured VBA error help file."
    RequireError Me, "HelpFile"
    HelpFile = mErrorHelpFile
End Property

Public Property Get HelpContext() As Long
Attribute HelpContext.VB_Description = "Gets the captured VBA error help context."
    RequireError Me, "HelpContext"
    HelpContext = mErrorHelpContext
End Property

Friend Property Get InternalRole() As Long
Attribute InternalRole.VB_MemberFlags = "40"
    InternalRole = mRole
End Property

Friend Property Get InternalOperator() As Long
Attribute InternalOperator.VB_MemberFlags = "40"
    InternalOperator = mOperator
End Property

Friend Property Get InternalExpressionOption() As Long
Attribute InternalExpressionOption.VB_MemberFlags = "40"
    InternalExpressionOption = mExpressionOption
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

Friend Property Get InternalIsByRef() As Boolean
Attribute InternalIsByRef.VB_MemberFlags = "40"
    InternalIsByRef = mIsByRef
End Property

Friend Property Get InternalNext() As ROneCOne
Attribute InternalNext.VB_MemberFlags = "40"
    Set InternalNext = mNext
End Property

Friend Property Get InternalParameters() As Collection
Attribute InternalParameters.VB_MemberFlags = "40"
    Set InternalParameters = mParameters
End Property

Friend Property Get InternalTarget() As Object
Attribute InternalTarget.VB_MemberFlags = "40"
    Set InternalTarget = mTarget
End Property

Friend Property Get InternalMethodName() As String
Attribute InternalMethodName.VB_MemberFlags = "40"
    InternalMethodName = mMethodName
End Property

Friend Property Get InternalQueryDelegate() As ROneCOne
Attribute InternalQueryDelegate.VB_MemberFlags = "40"
    Set InternalQueryDelegate = mQueryDelegate
End Property

Friend Property Get InternalQueryValue() As ROneCOne
Attribute InternalQueryValue.VB_MemberFlags = "40"
    Set InternalQueryValue = mQueryValue
End Property

Friend Property Get InternalQueryComparer() As ROneCOne
Attribute InternalQueryComparer.VB_MemberFlags = "40"
    Set InternalQueryComparer = mQueryComparer
End Property

Friend Property Get InternalMethodArity() As Long
Attribute InternalMethodArity.VB_MemberFlags = "40"
    InternalMethodArity = mMethodArity
End Property

Friend Property Get InternalIsActionDelegate() As Boolean
Attribute InternalIsActionDelegate.VB_MemberFlags = "40"
    InternalIsActionDelegate = mIsActionDelegate
End Property

Friend Property Get InternalHasParameterSignature() As Boolean
Attribute InternalHasParameterSignature.VB_MemberFlags = "40"
    InternalHasParameterSignature = mHasParameterSignature
End Property

Friend Property Get InternalHasReturnSignature() As Boolean
Attribute InternalHasReturnSignature.VB_MemberFlags = "40"
    InternalHasReturnSignature = mHasReturnSignature
End Property

Friend Property Get InternalReturnType() As VbVarType
Attribute InternalReturnType.VB_MemberFlags = "40"
    InternalReturnType = mReturnType
End Property

Friend Property Get InternalReturnObjectType() As String
Attribute InternalReturnObjectType.VB_MemberFlags = "40"
    InternalReturnObjectType = mReturnObjectType
End Property

Friend Property Get InternalInvocationList() As Collection
Attribute InternalInvocationList.VB_MemberFlags = "40"
    Set InternalInvocationList = mInvocationList
End Property

Friend Property Get InternalNativeAddress() As LongPtr
Attribute InternalNativeAddress.VB_MemberFlags = "40"
    InternalNativeAddress = mNativeAddress
End Property

Friend Property Get InternalReferenceAddress() As LongPtr
Attribute InternalReferenceAddress.VB_MemberFlags = "40"
    InternalReferenceAddress = mReferenceAddress
End Property

Friend Property Get InternalTryBody() As ROneCOne
Attribute InternalTryBody.VB_MemberFlags = "40"
    Set InternalTryBody = mTryBody
End Property

Friend Property Get InternalCatchActions() As Collection
Attribute InternalCatchActions.VB_MemberFlags = "40"
    Set InternalCatchActions = mCatchActions
End Property

Friend Property Get InternalCatchNumbers() As Collection
Attribute InternalCatchNumbers.VB_MemberFlags = "40"
    Set InternalCatchNumbers = mCatchNumbers
End Property

Friend Property Get InternalFinallyAction() As ROneCOne
Attribute InternalFinallyAction.VB_MemberFlags = "40"
    Set InternalFinallyAction = mFinallyAction
End Property

Friend Sub ConfigureParameter( _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String, _
    Optional ByVal isByRef As Boolean = False _
)
    mRole = ROLE_PARAMETER
    mExpectedType = expectedType
    mExpectedObjectType = expectedObjectType
    mIsByRef = isByRef
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
    ByVal rightValue As ROneCOne, _
    Optional ByVal expressionOption As Long = 0 _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    mExpressionOption = expressionOption
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
    mMethodArity = parameters.Count
    mHasParameterSignature = True
    mIsActionDelegate = False
End Sub

Friend Sub ConfigureMemberExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal memberName As String _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    mMethodName = memberName
End Sub

Friend Sub ConfigureMembershipExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal sequenceValue As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    Set mQueryValue = sequenceValue
    Set mQueryComparer = comparer
End Sub

Friend Sub ConfigureQuantifierExpression( _
    ByVal expressionOperator As Long, _
    ByVal leftValue As ROneCOne, _
    ByVal predicate As ROneCOne _
)
    mRole = ROLE_EXPRESSION
    mOperator = expressionOperator
    Set mLeft = leftValue
    Set mQueryDelegate = predicate
End Sub

Friend Sub ConfigureMethod( _
    ByVal target As Object, _
    ByVal methodName As String, _
    ByVal isAction As Boolean _
)
    mRole = ROLE_METHOD
    Set mTarget = target
    mMethodName = methodName
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureProcedure(ByVal procedureName As String, ByVal isAction As Boolean)
    mRole = ROLE_PROCEDURE
    mMethodName = procedureName
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureNative(ByVal procedureAddress As LongPtr, ByVal isAction As Boolean)
    mRole = ROLE_NATIVE
    mNativeAddress = procedureAddress
    mMethodArity = -1
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
    End If
End Sub

Friend Sub ConfigureReference(ByVal variableAddress As LongPtr, ByVal valueType As VbVarType)
    mRole = ROLE_REFERENCE
    mReferenceAddress = variableAddress
    mExpectedType = valueType
End Sub

Friend Sub ConfigureEvent(ByVal descriptors As Collection)
    mRole = ROLE_EVENT
    Set mParameters = descriptors
    mMethodArity = descriptors.Count
    mHasParameterSignature = True
    mHasReturnSignature = True
    mIsActionDelegate = True
    mReturnType = vbEmpty
    Set mInvocationList = New Collection
End Sub

Friend Sub ConfigureTry(ByVal work As ROneCOne)
    mRole = ROLE_TRY
    Set mTryBody = work
    Set mCatchActions = New Collection
    Set mCatchNumbers = New Collection
End Sub

Friend Sub ConfigureCatch(ByVal errorNumber As Long, ByVal handler As ROneCOne)
    mCatchNumbers.Add errorNumber
    mCatchActions.Add handler
End Sub

Friend Sub ConfigureFinally(ByVal cleanup As ROneCOne)
    Set mFinallyAction = cleanup
End Sub

Friend Sub ConfigureError( _
    ByVal errorNumber As Long, _
    ByVal errorSource As String, _
    ByVal errorDescription As String, _
    ByVal errorHelpFile As String, _
    ByVal errorHelpContext As Long _
)
    mRole = ROLE_ERROR
    mErrorNumber = errorNumber
    mErrorSource = errorSource
    mErrorDescription = errorDescription
    mErrorHelpFile = errorHelpFile
    mErrorHelpContext = errorHelpContext
End Sub

Friend Sub ConfigureParameterSignature(ByVal descriptors As Collection)
    Set mParameters = descriptors
    mMethodArity = descriptors.Count
    mHasParameterSignature = True
End Sub

Friend Sub ConfigureReturnSignature( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mReturnType = valueType
    mReturnObjectType = objectType
    mHasReturnSignature = True
    mIsActionDelegate = False
End Sub

Friend Sub ConfigureActionSignature(ByVal isAction As Boolean)
    mIsActionDelegate = isAction
    If isAction Then
        mHasReturnSignature = True
        mReturnType = vbEmpty
        mReturnObjectType = vbNullString
    End If
End Sub

Friend Sub ConfigureMulticast(ByVal invocationList As Collection)
    Dim firstDelegate As ROneCOne

    mRole = ROLE_MULTICAST
    Set mInvocationList = invocationList
    Set firstDelegate = invocationList.Item(1)
    CopySignatureFrom firstDelegate
End Sub

Friend Sub ConfigureComposition( _
    ByVal firstDelegate As ROneCOne, _
    ByVal nextDelegate As ROneCOne _
)
    mRole = ROLE_COMPOSITION
    Set mLeft = firstDelegate
    Set mNext = nextDelegate
    mMethodArity = firstDelegate.Arity
    mIsActionDelegate = nextDelegate.IsAction
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
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
)
    mRole = ROLE_QUERY
    Set mSource = source
    mOperator = queryOperator
    Set mQueryDelegate = queryDelegate
    Set mQueryComparer = queryComparer
    Set mQueryValue = queryValue
    mQueryCount = queryCount
    mElementType = elementType
    mElementObjectType = elementObjectType
    mEnumerationVersion = -1
End Sub

Friend Sub ConfigureWhereBuilder( _
    ByVal source As ROneCOne, _
    ByVal memberPath As String _
)
    mRole = ROLE_WHERE_BUILDER
    Set mSource = source
    mContextMemberPath = NormalizeMemberPath(memberPath, "Where")
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
        Case QUERY_WHERE_NOT
            Set values = ApplyWhere(sourceValues, True)
        Case QUERY_SELECT
            Set values = ApplySelect(sourceValues)
        Case QUERY_TAKE
            Set values = ApplyTake(sourceValues)
        Case QUERY_SKIP
            Set values = ApplySkip(sourceValues)
        Case QUERY_DISTINCT
            Set values = ApplyDistinct(sourceValues)
        Case QUERY_DISTINCT_BY
            Set values = ApplyDistinctBy(sourceValues)
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

Private Function CreateCallableDelegate( _
    ByVal callable As Variant, _
    ByVal memberName As String, _
    ByVal isAction As Boolean _
) As ROneCOne
    Dim candidate As Object
    Dim created As ROneCOne
    Dim runtimeValue As ROneCOne
    Dim resolvedMember As String
    Dim procedureName As String

    If IsObject(callable) Then
        Set candidate = callable
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
                "The callable object cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            Set runtimeValue = candidate
            If IsDelegateRole(runtimeValue.InternalRole) Then
                If Not isAction And runtimeValue.InternalIsActionDelegate Then
                    RaiseContractError ERROR_INVALID_OPERATION, "Func", _
                        "An Action cannot be converted to a Func without a result."
                End If
                Set created = CloneDelegate(runtimeValue)
                created.ConfigureActionSignature isAction
                Set CreateCallableDelegate = created
                Exit Function
            End If
            RequireExpression runtimeValue, "Func"
            Set created = ROneCOne.Lambda(runtimeValue)
            created.ConfigureActionSignature isAction
            Set CreateCallableDelegate = created
            Exit Function
        End If

        resolvedMember = Trim$(memberName)
        If Len(resolvedMember) = 0 Then resolvedMember = "Run"
        Set created = New ROneCOne
        created.ConfigureMethod candidate, resolvedMember, isAction
        Set CreateCallableDelegate = created
        Exit Function
    End If

    If VarType(callable) <> vbString Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "A callable must be an expression, object, or workbook procedure name."
    End If
    If Len(Trim$(memberName)) <> 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "A workbook procedure does not accept a separate member name."
    End If
    procedureName = Trim$(CStr(callable))
    If Len(procedureName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Func", _
            "The workbook procedure name cannot be empty."
    End If
    Set created = New ROneCOne
    created.ConfigureProcedure procedureName, isAction
    Set CreateCallableDelegate = created
End Function

Private Function CloneDelegate(ByVal source As ROneCOne) As ROneCOne
    Dim clone As ROneCOne

    RequireDelegate source, "CloneDelegate"
    Set clone = New ROneCOne
    Select Case source.InternalRole
        Case ROLE_LAMBDA
            clone.ConfigureLambda source.InternalLeft, source.InternalParameters
        Case ROLE_METHOD
            clone.ConfigureMethod source.InternalTarget, source.InternalMethodName, _
                source.InternalIsActionDelegate
        Case ROLE_PROCEDURE
            clone.ConfigureProcedure source.InternalMethodName, _
                source.InternalIsActionDelegate
        Case ROLE_NATIVE
            clone.ConfigureNative source.InternalNativeAddress, _
                source.InternalIsActionDelegate
        Case ROLE_COMPOSITION
            clone.ConfigureComposition source.InternalLeft, source.InternalNext
        Case ROLE_MULTICAST
            clone.ConfigureMulticast source.InternalInvocationList
    End Select
    If source.InternalHasParameterSignature Then
        clone.ConfigureParameterSignature source.InternalParameters
    End If
    If source.InternalHasReturnSignature Then
        If source.InternalIsActionDelegate Then
            clone.ConfigureActionSignature True
        Else
            clone.ConfigureReturnSignature source.InternalReturnType, _
                source.InternalReturnObjectType
        End If
    End If
    Set CloneDelegate = clone
End Function

Private Function CloneTryOperation(ByVal source As ROneCOne) As ROneCOne
    Dim clone As ROneCOne
    Dim index As Long

    RequireTry source, "CloneTryOperation"
    Set clone = New ROneCOne
    clone.ConfigureTry source.InternalTryBody
    For index = 1 To source.InternalCatchActions.Count
        clone.ConfigureCatch source.InternalCatchNumbers.Item(index), _
            source.InternalCatchActions.Item(index)
    Next index
    If Not source.InternalFinallyAction Is Nothing Then
        clone.ConfigureFinally source.InternalFinallyAction
    End If
    Set CloneTryOperation = clone
End Function

Private Sub ValidateAction( _
    ByVal actionValue As ROneCOne, _
    ByVal memberName As String, _
    ByVal requiresNoArguments As Boolean _
)
    RequireDelegate actionValue, memberName
    If Not actionValue.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires an Action delegate."
    End If
    If requiresNoArguments Then
        If actionValue.Arity <> 0 And actionValue.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The Action must accept no arguments."
        End If
    End If
End Sub

Friend Function InternalWrapElement( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Set InternalWrapElement = WrapElement(candidate, memberName)
End Function

Private Sub ValidateCatchAction(ByVal actionValue As ROneCOne)
    Dim descriptor As ROneCOne

    ValidateAction actionValue, "Catch", False
    If Not actionValue.InternalHasParameterSignature Then
        RaiseContractError ERROR_INVALID_OPERATION, "Catch", _
            "A Catch Action requires a complete Takes signature."
    End If
    If actionValue.Arity <> 0 And actionValue.Arity <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Catch", _
            "A Catch Action must accept zero arguments or one exception."
    End If
    If actionValue.Arity = 1 And actionValue.InternalHasParameterSignature Then
        Set descriptor = actionValue.InternalParameters.Item(1)
        If descriptor.InternalExpectedType <> vbObject Or _
            StrComp(descriptor.InternalExpectedObjectType, "ROneCOne", _
                vbTextCompare) <> 0 Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Catch", _
                "The Catch argument must use ROneCOne.Exception as its type."
        End If
    End If
End Sub

Private Sub ValidateEventHandler( _
    ByVal handler As ROneCOne, _
    ByVal memberName As String _
)
    Dim index As Long

    ValidateAction handler, memberName, False
    If Not handler.InternalHasParameterSignature Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "An event handler requires a complete Takes signature."
    End If
    If handler.Arity <> mMethodArity Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The event handler arity does not match the event."
    End If
    For index = 1 To mParameters.Count
        If Not ParameterDescriptorsMatch( _
            mParameters.Item(index), handler.InternalParameters.Item(index) _
        ) Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "The event handler signature does not match the event."
        End If
    Next index
End Sub

Private Sub CopySignatureFrom(ByVal source As ROneCOne)
    Set mParameters = source.InternalParameters
    mMethodArity = source.InternalMethodArity
    mIsActionDelegate = source.InternalIsActionDelegate
    mHasParameterSignature = source.InternalHasParameterSignature
    mHasReturnSignature = source.InternalHasReturnSignature
    mReturnType = source.InternalReturnType
    mReturnObjectType = source.InternalReturnObjectType
End Sub

Private Function CreateParameterDescriptor( _
    ByVal typeToken As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim descriptor As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    If IsObject(typeToken) Then
        Set candidate = typeToken
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set descriptor = candidate
                If descriptor.InternalRole = ROLE_PARAMETER And _
                    descriptor.InternalIsByRef Then
                    Set CreateParameterDescriptor = descriptor
                    Exit Function
                End If
            End If
        End If
    End If

    ResolveTypeDescriptor typeToken, valueType, objectType, memberName
    Set descriptor = New ROneCOne
    descriptor.ConfigureParameter valueType, objectType
    Set CreateParameterDescriptor = descriptor
End Function

Private Sub ValidateReferenceType( _
    ByVal valueType As VbVarType, _
    ByVal memberName As String _
)
    Select Case valueType
        Case vbByte, vbInteger, vbLong, vbLongLong
        Case vbSingle, vbDouble, vbCurrency, vbDate, vbBoolean
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "ByRef currently supports numeric, Date, and Boolean variables."
    End Select
End Sub

Private Function CreateReference( _
    ByVal variableAddress As LongPtr, _
    ByVal valueType As VbVarType _
) As ROneCOne
    Dim reference As ROneCOne

    Set reference = New ROneCOne
    reference.ConfigureReference variableAddress, valueType
    Set CreateReference = reference
End Function

Private Function BuildDynamicArguments(ByVal arguments As Variant) As Collection
    Dim candidate As Object
    Dim rawArgument As Variant
    Dim result As Collection
    Dim values As Variant

    Set result = New Collection
    If IsObject(arguments) Then
        Set candidate = arguments
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "DynamicInvoke", _
                "The argument collection cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            values = candidate.ToArray
            For Each rawArgument In values
                result.Add WrapInvocationArgument(rawArgument)
            Next rawArgument
            Set BuildDynamicArguments = result
            Exit Function
        End If
        If TypeName(candidate) = "Collection" Then
            For Each rawArgument In candidate
                result.Add WrapInvocationArgument(rawArgument)
            Next rawArgument
            Set BuildDynamicArguments = result
            Exit Function
        End If
    ElseIf IsArray(arguments) Then
        For Each rawArgument In arguments
            result.Add WrapInvocationArgument(rawArgument)
        Next rawArgument
        Set BuildDynamicArguments = result
        Exit Function
    ElseIf IsEmpty(arguments) Then
        Set BuildDynamicArguments = result
        Exit Function
    End If

    RaiseContractError ERROR_INVALID_ARGUMENT, "DynamicInvoke", _
        "Arguments must be a VBA array, Collection, or ROneCOne sequence."
End Function

Private Function WrapInvocationArgument(ByVal rawArgument As Variant) As ROneCOne
    Dim argumentObject As Object
    Dim wrapped As ROneCOne

    If IsObject(rawArgument) Then
        Set argumentObject = rawArgument
        If Not argumentObject Is Nothing Then
            If TypeOf argumentObject Is ROneCOne Then
                Set wrapped = argumentObject
                If wrapped.InternalRole = ROLE_REFERENCE Then
                    Set WrapInvocationArgument = wrapped
                    Exit Function
                End If
            End If
        End If
    End If

    Set wrapped = New ROneCOne
    If IsObject(rawArgument) Then
        wrapped.ConfigureObjectValue argumentObject
    Else
        wrapped.ConfigureScalarValue rawArgument
    End If
    Set WrapInvocationArgument = wrapped
End Function

Private Sub AssignPublicResult( _
    ByRef destination As Variant, _
    ByVal result As ROneCOne _
)
    If result.InternalHasObjectValue Then
        Set destination = result.InternalObjectValue
    Else
        destination = result.InternalScalarValue
    End If
End Sub

Private Sub AssignDefaultPublicResult(ByRef destination As Variant)
    If mElementType = vbObject Then
        Set destination = Nothing
        Exit Sub
    End If
    Select Case mElementType
        Case vbString
            destination = vbNullString
        Case vbBoolean
            destination = False
        Case vbDate
            destination = CDate(0)
        Case vbByte
            destination = CByte(0)
        Case vbInteger
            destination = CInt(0)
        Case vbLong
            destination = CLng(0)
        Case vbLongLong
            destination = CLngLng(0)
        Case vbSingle
            destination = CSng(0)
        Case vbDouble
            destination = CDbl(0)
        Case vbCurrency
            destination = CCur(0)
        Case Else
            destination = Empty
    End Select
End Sub

Private Function IsDelegateRole(ByVal roleValue As Long) As Boolean
    Select Case roleValue
        Case ROLE_LAMBDA, ROLE_METHOD, ROLE_COMPOSITION
            IsDelegateRole = True
        Case ROLE_PROCEDURE, ROLE_MULTICAST, ROLE_NATIVE
            IsDelegateRole = True
    End Select
End Function

Private Sub AppendInvocationList( _
    ByVal destination As Collection, _
    ByVal delegateValue As ROneCOne _
)
    Dim item As ROneCOne

    If delegateValue.InternalRole = ROLE_MULTICAST Then
        For Each item In delegateValue.InternalInvocationList
            destination.Add item
        Next item
    Else
        destination.Add delegateValue
    End If
End Sub

Private Sub ValidateInvocationCompatibility(ByVal invocationList As Collection)
    Dim firstDelegate As ROneCOne
    Dim index As Long

    Set firstDelegate = invocationList.Item(1)
    For index = 2 To invocationList.Count
        If Not DelegateSignaturesMatch(firstDelegate, invocationList.Item(index)) Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Combine", _
                "Multicast delegates must have matching signatures."
        End If
    Next index
End Sub

Private Function DelegateSignaturesMatch( _
    ByVal leftDelegate As ROneCOne, _
    ByVal rightDelegate As ROneCOne _
) As Boolean
    Dim index As Long

    If leftDelegate.IsAction <> rightDelegate.IsAction Then Exit Function
    If leftDelegate.Arity <> rightDelegate.Arity Then Exit Function
    If leftDelegate.InternalHasParameterSignature <> _
        rightDelegate.InternalHasParameterSignature Then Exit Function
    If leftDelegate.InternalHasReturnSignature <> _
        rightDelegate.InternalHasReturnSignature Then Exit Function
    If leftDelegate.InternalHasReturnSignature Then
        If leftDelegate.InternalReturnType <> _
            rightDelegate.InternalReturnType Then Exit Function
        If StrComp( _
            leftDelegate.InternalReturnObjectType, _
            rightDelegate.InternalReturnObjectType, _
            vbTextCompare _
        ) <> 0 Then Exit Function
    End If
    If leftDelegate.InternalHasParameterSignature Then
        For index = 1 To leftDelegate.InternalParameters.Count
            If Not ParameterDescriptorsMatch( _
                leftDelegate.InternalParameters.Item(index), _
                rightDelegate.InternalParameters.Item(index) _
            ) Then Exit Function
        Next index
    End If
    DelegateSignaturesMatch = True
End Function

Private Function ParameterDescriptorsMatch( _
    ByVal leftParameter As ROneCOne, _
    ByVal rightParameter As ROneCOne _
) As Boolean
    If leftParameter.InternalExpectedType <> _
        rightParameter.InternalExpectedType Then Exit Function
    If leftParameter.InternalIsByRef <> _
        rightParameter.InternalIsByRef Then Exit Function
    ParameterDescriptorsMatch = (StrComp( _
        leftParameter.InternalExpectedObjectType, _
        rightParameter.InternalExpectedObjectType, _
        vbTextCompare) = 0)
End Function

Private Function RemoveListItem(ByVal value As Variant) As Boolean
    Dim index As Long

    RequireList Me, "Remove"
    index = IndexOf(value)
    If index >= 0 Then
        mItems.Remove index + 1
        mEnumerationValues.Remove index + 1
        mVersion = mVersion + 1
        RemoveListItem = True
    End If
End Function

Private Function RemoveDelegateValue( _
    ByVal source As ROneCOne, _
    ByVal value As ROneCOne _
) As ROneCOne
    Dim candidate As ROneCOne
    Dim index As Long
    Dim matchIndex As Long
    Dim offset As Long
    Dim remaining As Collection
    Dim sourceList As Collection
    Dim valueList As Collection

    RequireDelegate source, "Remove"
    RequireDelegate value, "Remove"
    Set sourceList = New Collection
    Set valueList = New Collection
    AppendInvocationList sourceList, source
    AppendInvocationList valueList, value
    matchIndex = 0
    For index = sourceList.Count - valueList.Count + 1 To 1 Step -1
        For offset = 1 To valueList.Count
            If Not DelegateIdentityEquals( _
                sourceList.Item(index + offset - 1), _
                valueList.Item(offset) _
            ) Then Exit For
        Next offset
        If offset > valueList.Count Then
            matchIndex = index
            Exit For
        End If
    Next index
    If matchIndex = 0 Then
        Set RemoveDelegateValue = source
        Exit Function
    End If

    Set remaining = New Collection
    For index = 1 To sourceList.Count
        If index < matchIndex Or index >= matchIndex + valueList.Count Then
            remaining.Add sourceList.Item(index)
        End If
    Next index
    If remaining.Count = 0 Then Exit Function
    If remaining.Count = 1 Then
        Set RemoveDelegateValue = remaining.Item(1)
        Exit Function
    End If
    Set candidate = New ROneCOne
    candidate.ConfigureMulticast remaining
    Set RemoveDelegateValue = candidate
End Function

Private Function DelegateIdentityEquals( _
    ByVal leftDelegate As ROneCOne, _
    ByVal rightDelegate As ROneCOne _
) As Boolean
    If leftDelegate Is rightDelegate Then
        DelegateIdentityEquals = True
        Exit Function
    End If
    If leftDelegate.InternalRole <> rightDelegate.InternalRole Then Exit Function
    If Not DelegateSignaturesMatch(leftDelegate, rightDelegate) Then Exit Function

    Select Case leftDelegate.InternalRole
        Case ROLE_METHOD
            DelegateIdentityEquals = _
                (leftDelegate.InternalTarget Is rightDelegate.InternalTarget) And _
                (StrComp(leftDelegate.InternalMethodName, _
                    rightDelegate.InternalMethodName, _
                    vbTextCompare) = 0)
        Case ROLE_PROCEDURE
            DelegateIdentityEquals = (StrComp( _
                leftDelegate.InternalMethodName, _
                rightDelegate.InternalMethodName, _
                vbTextCompare) = 0)
        Case ROLE_NATIVE
            DelegateIdentityEquals = _
                (leftDelegate.InternalNativeAddress = _
                    rightDelegate.InternalNativeAddress)
        Case ROLE_COMPOSITION
            DelegateIdentityEquals = _
                (leftDelegate.InternalLeft Is rightDelegate.InternalLeft) And _
                (leftDelegate.InternalNext Is rightDelegate.InternalNext)
    End Select
End Function

Private Function BuildSignature(ByVal delegateValue As ROneCOne) As String
    Dim descriptor As ROneCOne
    Dim parts As String

    If delegateValue.InternalHasParameterSignature Then
        For Each descriptor In delegateValue.InternalParameters
            If Len(parts) > 0 Then parts = parts & ", "
            If descriptor.InternalIsByRef Then parts = parts & "ref "
            If Len(descriptor.InternalExpectedObjectType) > 0 Then
                parts = parts & descriptor.InternalExpectedObjectType
            Else
                parts = parts & TypeNameForVarType(descriptor.InternalExpectedType)
            End If
        Next descriptor
    Else
        parts = "dynamic"
    End If

    If delegateValue.InternalIsActionDelegate Then
        If parts = vbNullString Then
            BuildSignature = "Action"
        Else
            BuildSignature = "Action<" & parts & ">"
        End If
        Exit Function
    End If
    If Len(parts) > 0 Then parts = parts & ", "
    If delegateValue.InternalHasReturnSignature Then
        If Len(delegateValue.InternalReturnObjectType) > 0 Then
            parts = parts & delegateValue.InternalReturnObjectType
        Else
            parts = parts & TypeNameForVarType(delegateValue.InternalReturnType)
        End If
    Else
        parts = parts & "dynamic"
    End If
    BuildSignature = "Func<" & parts & ">"
End Function

Private Function CreateQuery( _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal queryValue As ROneCOne, _
    ByVal queryCount As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
) As ROneCOne
    Dim query As ROneCOne

    RequireSequence Me, "LINQ"
    Set query = New ROneCOne
    query.ConfigureQuery Me, queryOperator, queryDelegate, queryValue, _
        queryCount, elementType, elementObjectType, queryComparer
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

Private Sub CommitInitialItems( _
    ByVal listValue As ROneCOne, _
    ByVal wrappedItems As Collection _
)
    Dim wrapped As ROneCOne

    For Each wrapped In wrappedItems
        listValue.InternalAddWrapped wrapped
    Next wrapped
End Sub

Private Function BuildIncomingItems( _
    ByVal values As Variant, _
    ByVal memberName As String _
) As Collection
    Dim candidate As Object
    Dim incoming As Collection
    Dim rawValue As Variant
    Dim sequence As ROneCOne

    Set incoming = New Collection
    If IsObject(values) Then
        Set candidate = values
        If candidate Is Nothing Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The incoming values cannot be Nothing."
        End If
        If TypeOf candidate Is ROneCOne Then
            Set sequence = candidate
            RequireSequence sequence, memberName
            Set BuildIncomingItems = sequence.MaterializeWrapped
            Exit Function
        End If
        If TypeName(candidate) = "Collection" Then
            For Each rawValue In candidate
                incoming.Add WrapElement(rawValue, memberName)
            Next rawValue
            Set BuildIncomingItems = incoming
            Exit Function
        End If
    ElseIf IsArray(values) Then
        For Each rawValue In values
            incoming.Add WrapElement(rawValue, memberName)
        Next rawValue
        Set BuildIncomingItems = incoming
        Exit Function
    End If
    RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
        "Expected a ROneCOne sequence, VBA array, or Collection."
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

Private Function NormalizeUnary( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim normalized As ROneCOne
    Dim parameters As Collection

    RequireSequence Me, memberName
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The expression or delegate cannot be Nothing."
    End If

    If IsDelegateRole(candidate.InternalRole) Then
        If candidate.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "A query selector or predicate requires a Func."
        End If
        If candidate.Arity <> 1 And candidate.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The delegate must accept exactly one value."
        End If
        Set NormalizeUnary = candidate
        Exit Function
    End If

    Select Case candidate.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            Set parameters = New Collection
            CollectExpressionParameters candidate, parameters
            If parameters.Count <> 1 Then
                RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                    "The expression must reference exactly one parameter."
            End If
            Set normalized = New ROneCOne
            normalized.ConfigureLambda candidate, parameters
            Set NormalizeUnary = normalized

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "Expected a unary expression or delegate."
    End Select
End Function

Private Sub CollectExpressionParameters( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection _
)
    Select Case node.InternalRole
        Case ROLE_VALUE
            Exit Sub

        Case ROLE_PARAMETER
            If Not CollectionContainsIdentity(parameters, node) Then
                parameters.Add node
            End If

        Case ROLE_EXPRESSION
            CollectExpressionParameters node.InternalLeft, parameters
            If Not node.InternalRight Is Nothing Then
                CollectExpressionParameters node.InternalRight, parameters
            End If

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, _
                "CollectExpressionParameters", _
                "The expression contains an unsupported node."
    End Select
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
    ByVal comparer As ROneCOne, _
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
        ElseIf findMaximum And _
            CompareWithComparer(candidate, best, comparer, memberName) > 0 Then
            Set best = candidate
        ElseIf Not findMaximum And _
            CompareWithComparer(candidate, best, comparer, memberName) < 0 Then
            Set best = candidate
        End If
    Next wrapped
    Set ExtremeValue = best
End Function

Private Function TryFindTerminalElement( _
    ByVal predicate As ROneCOne, _
    ByVal takeLast As Boolean, _
    ByVal memberName As String, _
    ByRef found As Boolean _
) As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        matches = predicate Is Nothing
        If Not predicate Is Nothing Then
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            found = True
            Set TryFindTerminalElement = wrapped
            If Not takeLast Then Exit Function
        End If
    Next wrapped
End Function

Private Function FindSingleElement( _
    ByVal predicate As ROneCOne, _
    ByVal allowDefault As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim found As ROneCOne
    Dim matches As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    For Each wrapped In values
        matches = predicate Is Nothing
        If Not predicate Is Nothing Then
            matches = InvokePredicate(predicate, wrapped, memberName)
        End If
        If matches Then
            If Not found Is Nothing Then
                RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                    "The sequence contains more than one matching element."
            End If
            Set found = wrapped
        End If
    Next wrapped
    If found Is Nothing And Not allowDefault Then RaiseEmptySequence memberName
    Set FindSingleElement = found
End Function

Private Function InvokeBinary( _
    ByVal delegateValue As ROneCOne, _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne _
) As ROneCOne
    Dim arguments As Collection

    Set arguments = New Collection
    arguments.Add leftValue
    arguments.Add rightValue
    Set InvokeBinary = delegateValue.InvokeValue(arguments)
End Function

Private Function NormalizeStandaloneUnary( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    Dim normalized As ROneCOne
    Dim parameters As Collection

    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The expression or delegate cannot be Nothing."
    End If
    If IsDelegateRole(candidate.InternalRole) Then
        If candidate.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "A predicate requires a Func."
        End If
        If candidate.Arity <> 1 And candidate.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The predicate must accept exactly one value."
        End If
        Set NormalizeStandaloneUnary = candidate
        Exit Function
    End If
    RequireExpression candidate, memberName
    Set parameters = New Collection
    CollectExpressionParameters candidate, parameters
    If parameters.Count <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The expression must reference exactly one parameter."
    End If
    Set normalized = New ROneCOne
    normalized.ConfigureLambda candidate, parameters
    Set NormalizeStandaloneUnary = normalized
End Function

Private Function CreateSequenceComparer( _
    ByVal callable As Variant, _
    ByVal memberName As String, _
    ByVal equality As Boolean, _
    ByVal context As String _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection
    Dim expectedType As VbVarType
    Dim objectType As String

    Set comparerValue = CreateCallableDelegate(callable, memberName, False)
    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Then
        expectedType = mElementType
        objectType = mElementObjectType
    Else
        expectedType = vbVariant
    End If
    If Not comparerValue.InternalHasParameterSignature Then
        Set descriptors = New Collection
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter expectedType, objectType
        descriptors.Add descriptor
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter expectedType, objectType
        descriptors.Add descriptor
        comparerValue.ConfigureParameterSignature descriptors
    End If
    If Not comparerValue.InternalHasReturnSignature Then
        If equality Then
            comparerValue.ConfigureReturnSignature vbBoolean, vbNullString
        Else
            comparerValue.ConfigureReturnSignature vbLong, vbNullString
        End If
    End If
    Set CreateSequenceComparer = NormalizeComparer( _
        comparerValue, equality, context)
End Function

Private Function NormalizeComparer( _
    ByVal candidate As Variant, _
    ByVal equality As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim comparerValue As ROneCOne

    Set comparerValue = RequireROneCOne(candidate, memberName)
    RequireDelegate comparerValue, memberName
    If comparerValue.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "A comparer requires a Func result."
    End If
    If comparerValue.Arity <> 2 And comparerValue.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "A comparer must accept exactly two values."
    End If
    If comparerValue.InternalHasReturnSignature Then
        If equality And comparerValue.InternalReturnType <> vbBoolean Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "An equality comparer must return Boolean."
        ElseIf Not equality And comparerValue.InternalReturnType <> vbLong Then
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "An ordering comparer must return Long."
        End If
    End If
    Set NormalizeComparer = comparerValue
End Function

Private Function ExtremeElement( _
    ByVal selector As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal findMaximum As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim bestElement As ROneCOne
    Dim bestKey As ROneCOne
    Dim candidateKey As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set values = MaterializeWrapped
    If values.Count = 0 Then RaiseEmptySequence memberName
    For Each wrapped In values
        Set candidateKey = InvokeUnary(selector, wrapped)
        If bestElement Is Nothing Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        ElseIf findMaximum And _
            CompareWithComparer( _
                candidateKey, bestKey, comparer, memberName) > 0 Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        ElseIf Not findMaximum And _
            CompareWithComparer( _
                candidateKey, bestKey, comparer, memberName) < 0 Then
            Set bestElement = wrapped
            Set bestKey = candidateKey
        End If
    Next wrapped
    Set ExtremeElement = bestElement
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

Private Function ApplyWhere( _
    ByVal sourceValues As Collection, _
    Optional ByVal negate As Boolean = False _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If InvokePredicate(mQueryDelegate, wrapped, "Where") Xor negate Then
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
        If Not WrappedCollectionContains( _
            output, wrapped, mQueryComparer) Then output.Add wrapped
    Next wrapped
    Set ApplyDistinct = output
End Function

Private Function ApplyDistinctBy( _
    ByVal sourceValues As Collection _
) As Collection
    Dim key As ROneCOne
    Dim keys As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set keys = New Collection
    Set output = New Collection
    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains(keys, key, mQueryComparer) Then
            keys.Add key
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyDistinctBy = output
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
            comparison = CompareWithComparer( _
                key, keys.Item(index), mQueryComparer, "OrderBy")
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
    ByVal candidate As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
) As Boolean
    Dim wrapped As ROneCOne

    For Each wrapped In values
        If ValuesEqualWithComparer( _
            wrapped, candidate, comparer, "Contains") Then
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
    If IsNull(leftScalar) Or IsNull(rightScalar) Then
        If IsNull(leftScalar) And IsNull(rightScalar) Then
            CompareWrapped = 0
        ElseIf IsNull(leftScalar) Then
            CompareWrapped = -1
        Else
            CompareWrapped = 1
        End If
        Exit Function
    End If
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

Private Function ValuesEqualWithComparer( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    Dim result As ROneCOne

    If comparer Is Nothing Then
        ValuesEqualWithComparer = ValuesEqual(leftValue, rightValue)
    Else
        Set result = InvokeBinary(comparer, leftValue, rightValue)
        ValuesEqualWithComparer = CBool(RequireScalar(result, memberName))
    End If
End Function

Private Function CompareWithComparer( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal comparer As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim result As ROneCOne

    If comparer Is Nothing Then
        CompareWithComparer = CompareWrapped( _
            leftValue, rightValue, memberName)
    Else
        Set result = InvokeBinary(comparer, leftValue, rightValue)
        CompareWithComparer = CLng(RequireScalar(result, memberName))
    End If
End Function

Private Function NormalizeContextUnary( _
    ByVal candidate As Variant, _
    ByVal memberName As String, _
    ByVal returnsObject As Boolean _
) As ROneCOne
    Dim expression As ROneCOne

    RequireSequence Me, memberName
    If Not IsObject(candidate) Then
        If VarType(candidate) = vbString Then
            Set expression = BuildMemberPath(CStr(candidate), returnsObject)
            Set NormalizeContextUnary = NormalizeUnary(expression, memberName)
            Exit Function
        End If
    End If
    Set NormalizeContextUnary = NormalizeUnary( _
        RequireROneCOne(candidate, memberName), memberName)
End Function

Private Function BindContextPredicate( _
    ByVal predicateDelegate As ROneCOne _
) As ROneCOne
    Dim bound As ROneCOne
    Dim descriptor As ROneCOne
    Dim descriptors As Collection

    RequireDelegate predicateDelegate, "Predicate"
    If predicateDelegate.IsAction Then
        RaiseContractError ERROR_INVALID_OPERATION, "Predicate", _
            "A predicate requires a Func result."
    End If
    If predicateDelegate.Arity <> 1 And predicateDelegate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Predicate", _
            "The predicate must accept exactly one sequence element."
    End If
    If predicateDelegate.InternalHasReturnSignature Then
        If predicateDelegate.InternalReturnType <> vbBoolean Or _
            Len(predicateDelegate.InternalReturnObjectType) > 0 Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
                "The predicate result must be Boolean."
        End If
    End If
    If predicateDelegate.InternalHasParameterSignature Then
        Set descriptor = predicateDelegate.InternalParameters.Item(1)
        ValidateContextPredicateParameter descriptor
    End If

    Set bound = CloneDelegate(predicateDelegate)
    If Not bound.InternalHasParameterSignature Then
        Set descriptors = New Collection
        Set descriptor = New ROneCOne
        descriptor.ConfigureParameter mElementType, mElementObjectType
        descriptors.Add descriptor
        bound.ConfigureParameterSignature descriptors
    End If
    If Not bound.InternalHasReturnSignature Then
        bound.ConfigureReturnSignature vbBoolean, vbNullString
    End If
    Set BindContextPredicate = bound
End Function

Private Sub ValidateContextPredicateParameter( _
    ByVal descriptor As ROneCOne _
)
    Dim expectedObjectType As String
    Dim expectedType As VbVarType

    If descriptor.InternalIsByRef Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
            "A sequence predicate parameter must be passed ByVal."
    End If
    expectedType = descriptor.InternalExpectedType
    expectedObjectType = descriptor.InternalExpectedObjectType
    If expectedType = vbVariant Then Exit Sub

    If mElementType = vbObject Then
        If expectedType <> vbObject Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
                "The predicate parameter does not accept the sequence type."
        End If
        If Len(expectedObjectType) = 0 Then Exit Sub
        If StrComp(expectedObjectType, mElementObjectType, _
            vbTextCompare) = 0 Then Exit Sub
    ElseIf expectedType = mElementType Then
        Exit Sub
    End If
    RaiseContractError ERROR_TYPE_MISMATCH, "Predicate", _
        "The predicate parameter does not accept the sequence type."
End Sub

Private Function BuildMemberPath( _
    ByVal memberPath As String, _
    ByVal returnsObject As Boolean _
) As ROneCOne
    Dim expression As ROneCOne
    Dim index As Long
    Dim normalized As String
    Dim parts As Variant
    Dim propagateNullSafe As Boolean
    Dim segment As String
    Dim nullSafe As Boolean

    normalized = NormalizeMemberPath(memberPath, "Condition")
    parts = VBA.Split(normalized, ".")
    Set expression = Element
    For index = LBound(parts) To UBound(parts)
        segment = CStr(parts(index))
        nullSafe = (Left$(segment, 1) = "?")
        If nullSafe Then
            segment = Mid$(segment, 2)
            propagateNullSafe = True
        End If
        Set expression = expression.Member( _
            segment, index < UBound(parts) Or returnsObject, _
            propagateNullSafe)
    Next index
    Set BuildMemberPath = expression
End Function

Private Function NormalizeMemberPath( _
    ByVal memberPath As String, _
    ByVal memberName As String _
) As String
    Dim index As Long
    Dim normalized As String
    Dim parts As Variant
    Dim segment As String

    memberPath = Replace(memberPath, "?.", ".?")
    memberPath = Trim$(memberPath)
    If Len(memberPath) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The member path cannot be empty."
    End If
    parts = VBA.Split(memberPath, ".")
    For index = LBound(parts) To UBound(parts)
        segment = Trim$(CStr(parts(index)))
        If Left$(segment, 1) = "?" Then
            segment = "?" & Trim$(Mid$(segment, 2))
        End If
        If Len(segment) = 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "The member path cannot contain an empty segment."
        End If
        If segment = "?" Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A null-safe member segment must include a member name."
        End If
        If Len(normalized) > 0 Then normalized = normalized & "."
        normalized = normalized & segment
    Next index
    NormalizeMemberPath = normalized
End Function

Private Function ContextMemberExpression( _
    ByVal returnsObject As Boolean _
) As ROneCOne
    If mRole <> ROLE_WHERE_BUILDER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Where", _
            "A contextual member condition must begin with Where(memberName)."
    End If
    Set ContextMemberExpression = mSource.Condition( _
        mContextMemberPath, returnsObject)
End Function

Private Function ConditionReceiver( _
    ByVal returnsObject As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    If mRole = ROLE_WHERE_BUILDER Then
        Set ConditionReceiver = ContextMemberExpression(returnsObject)
    Else
        RequireExpression Me, memberName & " receiver"
        Set ConditionReceiver = Me
    End If
End Function

Private Function CompleteWhereBinary( _
    ByVal expressionOperator As Long, _
    ByVal operand As Variant _
) As ROneCOne
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    Set receiver = ContextMemberExpression(False)
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, receiver, rightValue
    Set CompleteWhereBinary = mSource.Where(expression)
End Function

Private Function CompleteStringCondition( _
    ByVal expressionOperator As Long, _
    ByVal operand As String, _
    ByVal compareMode As VbCompareMethod _
) As ROneCOne
    Dim expression As ROneCOne
    Dim receiver As ROneCOne
    Dim rightValue As ROneCOne

    ValidateStringCompareMode compareMode, "String condition"
    Set receiver = ConditionReceiver(False, "String condition")
    Set rightValue = CoerceExpression(operand)
    Set expression = New ROneCOne
    expression.ConfigureExpression expressionOperator, receiver, _
        rightValue, compareMode
    Set CompleteStringCondition = CompleteWhereOrExpression(expression)
End Function

Private Sub ValidateStringCompareMode( _
    ByVal compareMode As VbCompareMethod, _
    ByVal memberName As String _
)
    If compareMode <> vbBinaryCompare And compareMode <> vbTextCompare Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Comparison mode must be vbBinaryCompare or vbTextCompare."
    End If
End Sub

Private Function CreateMembershipCondition( _
    ByVal values As Variant, _
    ByVal negate As Boolean, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim expression As ROneCOne
    Dim operatorCode As Long
    Dim receiver As ROneCOne
    Dim sequenceValue As ROneCOne

    Set receiver = ConditionReceiver(False, "In")
    Set sequenceValue = NormalizeMembershipSequence(values, "In")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "IsIn")
    End If
    operatorCode = OP_IN_SEQUENCE
    If negate Then operatorCode = OP_NOT_IN_SEQUENCE
    Set expression = New ROneCOne
    expression.ConfigureMembershipExpression _
        operatorCode, receiver, sequenceValue, comparerValue
    Set CreateMembershipCondition = CompleteWhereOrExpression(expression)
End Function

Private Function NormalizeMembershipSequence( _
    ByVal values As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim sequenceValue As ROneCOne

    If IsObject(values) Then
        Set candidate = values
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set sequenceValue = candidate
                RequireSequence sequenceValue, memberName
                Set NormalizeMembershipSequence = sequenceValue
                Exit Function
            End If
        End If
    End If
    Set sequenceValue = ROneCOne.ListOf(vbVariant)
    sequenceValue.AddRange values
    Set NormalizeMembershipSequence = sequenceValue
End Function

Private Function IsSequenceContainer(ByVal value As Variant) As Boolean
    Dim candidate As Object

    If IsArray(value) Then
        IsSequenceContainer = True
        Exit Function
    End If
    If Not IsObject(value) Then Exit Function
    Set candidate = value
    If candidate Is Nothing Then Exit Function
    If TypeOf candidate Is ROneCOne Then
        IsSequenceContainer = _
            candidate.InternalRole = ROLE_LIST Or _
            candidate.InternalRole = ROLE_QUERY
    Else
        IsSequenceContainer = (TypeName(candidate) = "Collection")
    End If
End Function

Private Function IsExpressionVariant(ByVal candidate As Variant) As Boolean
    Dim candidateObject As Object
    Dim runtimeValue As ROneCOne

    If Not IsObject(candidate) Then Exit Function
    Set candidateObject = candidate
    If candidateObject Is Nothing Then Exit Function
    If Not TypeOf candidateObject Is ROneCOne Then Exit Function
    Set runtimeValue = candidateObject
    Select Case runtimeValue.InternalRole
        Case ROLE_VALUE, ROLE_PARAMETER, ROLE_EXPRESSION
            IsExpressionVariant = True
    End Select
End Function

Private Function CreateQuantifierCondition( _
    ByVal operatorCode As Long, _
    ByVal predicate As Variant, _
    ByVal hasPredicate As Boolean _
) As ROneCOne
    Dim expression As ROneCOne
    Dim predicateValue As ROneCOne
    Dim receiver As ROneCOne

    Set receiver = ConditionReceiver(True, "Nested quantifier")
    If hasPredicate Then
        Set predicateValue = NormalizeStandaloneUnary( _
            RequireROneCOne(predicate, "Nested quantifier"), _
            "Nested quantifier")
    End If
    Set expression = New ROneCOne
    expression.ConfigureQuantifierExpression _
        operatorCode, receiver, predicateValue
    Set CreateQuantifierCondition = CompleteWhereOrExpression(expression)
End Function

Private Function CompleteWhereOrExpression( _
    ByVal expression As ROneCOne _
) As ROneCOne
    If mRole = ROLE_WHERE_BUILDER Then
        Set CompleteWhereOrExpression = mSource.Where(expression)
    Else
        Set CompleteWhereOrExpression = expression
    End If
End Function

Private Sub ExecuteTryOperation()
    Dim arguments As Collection
    Dim catchAction As ROneCOne
    Dim catchError As ROneCOne
    Dim errorValue As ROneCOne
    Dim finallyError As ROneCOne

    RequireTry Me, "Execute"
    Set arguments = New Collection
    If Not InvokeActionCaptured(mTryBody, arguments, errorValue) Then
        Set catchAction = FindCatchAction(errorValue.ErrorNumber)
        If Not catchAction Is Nothing Then
            Set arguments = New Collection
            If catchAction.Arity <> 0 Then
                arguments.Add WrapInvocationArgument(errorValue)
            End If
            If InvokeActionCaptured(catchAction, arguments, catchError) Then
                Set errorValue = Nothing
            Else
                Set errorValue = catchError
            End If
        End If
    End If

    If Not mFinallyAction Is Nothing Then
        Set arguments = New Collection
        If Not InvokeActionCaptured(mFinallyAction, arguments, finallyError) Then
            Set errorValue = finallyError
        End If
    End If
    If Not errorValue Is Nothing Then RaiseCapturedError errorValue
End Sub

Private Function InvokeActionCaptured( _
    ByVal actionValue As ROneCOne, _
    ByVal arguments As Collection, _
    ByRef errorValue As ROneCOne _
) As Boolean
    Dim ignored As ROneCOne

    On Error GoTo InvocationFailed
    Set ignored = actionValue.InvokeValue(arguments)
    InvokeActionCaptured = True
    Exit Function

InvocationFailed:
    Set errorValue = New ROneCOne
    errorValue.ConfigureError Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
    Err.Clear
    On Error GoTo 0
End Function

Private Function FindCatchAction(ByVal errorNumber As Long) As ROneCOne
    Dim index As Long

    For index = 1 To mCatchActions.Count
        If CLng(mCatchNumbers.Item(index)) = 0 Or _
            CLng(mCatchNumbers.Item(index)) = errorNumber Then
            Set FindCatchAction = mCatchActions.Item(index)
            Exit Function
        End If
    Next index
End Function

Private Sub RaiseCapturedError(ByVal errorValue As ROneCOne)
    Err.Raise errorValue.ErrorNumber, errorValue.ErrorSource, _
        errorValue.Message, errorValue.HelpFile, errorValue.HelpContext
End Sub

Friend Function InvokeValue(ByVal arguments As Collection) As ROneCOne
    Dim delegateValue As ROneCOne
    Dim firstResult As ROneCOne
    Dim nextArguments As Collection
    Dim result As ROneCOne

    Select Case mRole
        Case ROLE_LAMBDA
            If arguments.Count <> mParameters.Count Then
                RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                    "Expected " & CStr(mParameters.Count) & " argument(s), received " & _
                    CStr(arguments.Count) & "."
            End If
            Set result = EvaluateNode(mLeft, mParameters, arguments)

        Case ROLE_METHOD
            ValidateDelegateArguments arguments
            Set result = InvokeMethod(arguments)

        Case ROLE_PROCEDURE
            ValidateDelegateArguments arguments
            Set result = InvokeProcedure(arguments)

        Case ROLE_NATIVE
            ValidateDelegateArguments arguments
            Set result = InvokeNative(arguments)

        Case ROLE_MULTICAST
            ValidateDelegateArguments arguments
            For Each delegateValue In mInvocationList
                Set result = delegateValue.InvokeValue(arguments)
            Next delegateValue

        Case ROLE_COMPOSITION
            Set firstResult = mLeft.InvokeValue(arguments)
            Set nextArguments = New Collection
            nextArguments.Add firstResult
            Set result = mNext.InvokeValue(nextArguments)

        Case Else
            RaiseContractError ERROR_INVALID_OPERATION, "Run", _
                "Only delegates can be invoked."
    End Select
    Set InvokeValue = FinalizeDelegateResult(result)
End Function

Private Function EvaluateNode( _
    ByVal node As ROneCOne, _
    ByVal parameters As Collection, _
    ByVal arguments As Collection _
) As ROneCOne
    Dim leftResult As ROneCOne
    Dim leftText As String
    Dim objectResult As Object
    Dim rightResult As ROneCOne
    Dim rightText As String
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

            If IsMemberOperator(operatorCode) Then
                Set EvaluateNode = ReadMember( _
                    leftResult, _
                    node.InternalMethodName, _
                    IsObjectMemberOperator(operatorCode), _
                    IsSafeMemberOperator(operatorCode))
                Exit Function
            End If

            If Not node.InternalRight Is Nothing Then
                Set rightResult = EvaluateNode(node.InternalRight, parameters, arguments)
            End If

            If operatorCode = OP_IN_SEQUENCE Or _
                operatorCode = OP_NOT_IN_SEQUENCE Then
                scalarResult = SequenceContainsWrapped( _
                    node.InternalQueryValue, leftResult, _
                    node.InternalQueryComparer)
                If operatorCode = OP_NOT_IN_SEQUENCE Then
                    scalarResult = Not CBool(scalarResult)
                End If
                Set EvaluateNode = CreateScalarResult(scalarResult)
                Exit Function
            End If

            If IsQuantifierOperator(operatorCode) Then
                scalarResult = EvaluateNestedQuantifier( _
                    leftResult, node.InternalQueryDelegate, operatorCode)
                Set EvaluateNode = CreateScalarResult(scalarResult)
                Exit Function
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
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = RequireScalar(leftResult, "LessThan") < _
                            RequireScalar(rightResult, "LessThan")
                    End If
                Case OP_LESS_THAN_OR_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = _
                            RequireScalar(leftResult, "LessThanOrEqual") <= _
                            RequireScalar(rightResult, "LessThanOrEqual")
                    End If
                Case OP_GREATER_THAN
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = RequireScalar(leftResult, "GreaterThan") > _
                            RequireScalar(rightResult, "GreaterThan")
                    End If
                Case OP_GREATER_THAN_OR_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = False
                    Else
                        scalarResult = _
                            RequireScalar(leftResult, "GreaterThanOrEqual") >= _
                            RequireScalar(rightResult, "GreaterThanOrEqual")
                    End If
                Case OP_AND_ALSO
                    scalarResult = CBool(RequireScalar(rightResult, "AndAlso"))
                Case OP_OR_ELSE
                    scalarResult = CBool(RequireScalar(rightResult, "OrElse"))
                Case OP_NOT
                    scalarResult = Not CBool(RequireScalar(leftResult, "NotExpression"))
                Case OP_STARTS_WITH
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "StartsWith"))
                    rightText = CStr(RequireScalar(rightResult, "StartsWith"))
                    scalarResult = (StrComp( _
                        Left$(leftText, Len(rightText)), rightText, _
                        node.InternalExpressionOption) = 0)
                Case OP_ENDS_WITH
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "EndsWith"))
                    rightText = CStr(RequireScalar(rightResult, "EndsWith"))
                    scalarResult = (StrComp( _
                        Right$(leftText, Len(rightText)), rightText, _
                        node.InternalExpressionOption) = 0)
                Case OP_CONTAINS_TEXT
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "ContainsText"))
                    rightText = CStr(RequireScalar(rightResult, "ContainsText"))
                    scalarResult = (InStr( _
                        1, leftText, rightText, _
                        node.InternalExpressionOption) > 0)
                Case OP_IS_NOTHING, OP_IS_NOT_NOTHING
                    If Not leftResult.InternalHasObjectValue Then
                        RaiseContractError ERROR_TYPE_MISMATCH, "IsNothing", _
                            "IsNothing requires an object-valued expression."
                    End If
                    Set objectResult = leftResult.InternalObjectValue
                    scalarResult = (objectResult Is Nothing)
                    If operatorCode = OP_IS_NOT_NOTHING Then
                        scalarResult = Not CBool(scalarResult)
                    End If
                Case OP_IS_NULL_OR_EMPTY
                    If leftResult.InternalHasObjectValue Then
                        RaiseContractError ERROR_TYPE_MISMATCH, "IsNullOrEmpty", _
                            "IsNullOrEmpty requires a scalar expression."
                    End If
                    scalarResult = IsNull(leftResult.InternalScalarValue)
                    If Not CBool(scalarResult) Then
                        scalarResult = (Len(CStr( _
                            leftResult.InternalScalarValue)) = 0)
                    End If
                Case OP_MATCHES_PATTERN
                    If HasNullScalar(leftResult) Then GoTo FalseStringResult
                    leftText = CStr(RequireScalar(leftResult, "MatchesPattern"))
                    rightText = CStr(RequireScalar(rightResult, "MatchesPattern"))
                    If node.InternalExpressionOption = vbTextCompare Then
                        leftText = LCase$(leftText)
                        rightText = LCase$(rightText)
                    End If
                    scalarResult = leftText Like rightText
                Case OP_STRING_EQUAL, OP_STRING_NOT_EQUAL
                    If HasNullScalar(leftResult) Or HasNullScalar(rightResult) Then
                        scalarResult = ValuesEqual(leftResult, rightResult)
                    Else
                        scalarResult = (StrComp( _
                            CStr(RequireScalar(leftResult, "EqualToIgnoreCase")), _
                            CStr(RequireScalar(rightResult, "EqualToIgnoreCase")), _
                            node.InternalExpressionOption) = 0)
                    End If
                    If operatorCode = OP_STRING_NOT_EQUAL Then
                        scalarResult = Not CBool(scalarResult)
                    End If
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "EvaluateNode", _
                        "Unknown expression operator."
            End Select

            Set EvaluateNode = CreateScalarResult(scalarResult)
            Exit Function

FalseStringResult:
            Set EvaluateNode = CreateScalarResult(False)

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
    Dim values() As Variant

    If arguments.Count > 8 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Object-method delegates currently support at most 8 arguments."
    End If
    PrepareInvocationValues arguments, values

    On Error GoTo InvocationFailed
    If mHasReturnSignature And mReturnType = vbObject Then
        Select Case arguments.Count
            Case 0
                Set returnedObject = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, values(0))
            Case 2
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1))
            Case 3
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2))
            Case 4
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3))
            Case 5
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4))
            Case 6
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), values(5))
            Case 7
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6))
            Case 8
                Set returnedObject = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6), values(7))
        End Select
        Set InvokeMethod = CreateObjectResult(returnedObject)
    Else
        Select Case arguments.Count
            Case 0
                returnedValue = CallByName(mTarget, mMethodName, VbMethod)
            Case 1
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, values(0))
            Case 2
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1))
            Case 3
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2))
            Case 4
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3))
            Case 5
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4))
            Case 6
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), values(5))
            Case 7
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6))
            Case 8
                returnedValue = CallByName( _
                    mTarget, mMethodName, VbMethod, _
                    values(0), values(1), values(2), values(3), values(4), _
                    values(5), values(6), values(7))
        End Select
        Set InvokeMethod = CreateScalarResult(returnedValue)
    End If
    Exit Function

InvocationFailed:
    RaiseContractError ERROR_METHOD_INVOCATION, "Run", _
        "Method '" & mMethodName & "' failed with " & CStr(Err.Number) & ": " & _
        Err.Description
End Function

Private Function InvokeProcedure(ByVal arguments As Collection) As ROneCOne
    Dim procedureName As String
    Dim returnedObject As Object
    Dim returnedValue As Variant
    Dim values() As Variant

    If arguments.Count > 30 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Workbook procedure delegates support at most 30 arguments."
    End If
    PrepareInvocationValues arguments, values
    procedureName = ResolveProcedureName(mMethodName)

    On Error GoTo InvocationFailed
    If mHasReturnSignature And mReturnType = vbObject Then
        Set returnedObject = RunWorkbookProcedureObject(procedureName, values, arguments.Count)
        Set InvokeProcedure = CreateObjectResult(returnedObject)
    Else
        returnedValue = RunWorkbookProcedure(procedureName, values, arguments.Count)
        Set InvokeProcedure = CreateScalarResult(returnedValue)
    End If
    Exit Function

InvocationFailed:
    RaiseContractError ERROR_METHOD_INVOCATION, "Run", _
        "Procedure '" & mMethodName & "' failed with " & CStr(Err.Number) & ": " & _
        Err.Description
End Function

Private Function InvokeNative(ByVal arguments As Collection) As ROneCOne
    Dim argumentPointers() As LongPtr
    Dim argumentTypes() As Integer
    Dim argumentValues() As Variant
    Dim descriptor As ROneCOne
    Dim hResult As Long
    Dim index As Long
    Dim returnedValue As Variant
    Dim wrapped As ROneCOne

    If Not mHasParameterSignature Or _
        (Not mIsActionDelegate And Not mHasReturnSignature) Then
        RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
            "Native invocation requires complete Takes and Returns metadata."
    End If
    If arguments.Count > 0 Then
        ReDim argumentPointers(0 To arguments.Count - 1)
        ReDim argumentTypes(0 To arguments.Count - 1)
        ReDim argumentValues(0 To arguments.Count - 1)
        For index = 0 To arguments.Count - 1
            Set descriptor = mParameters.Item(index + 1)
            Set wrapped = arguments.Item(index + 1)
            If descriptor.InternalIsByRef Then
                argumentValues(index) = CLngLng(wrapped.InternalReferenceAddress)
                argumentTypes(index) = vbLongLong
            ElseIf wrapped.InternalHasObjectValue Then
                Set argumentValues(index) = wrapped.InternalObjectValue
                argumentTypes(index) = descriptor.InternalExpectedType
            Else
                argumentValues(index) = wrapped.InternalScalarValue
                argumentTypes(index) = descriptor.InternalExpectedType
            End If
            argumentPointers(index) = VarPtr(argumentValues(index))
        Next index
        hResult = DispCallFunc( _
            0, mNativeAddress, CALLCONV_STDCALL, CInt(mReturnType), _
            arguments.Count, VarPtr(argumentTypes(0)), _
            VarPtr(argumentPointers(0)), VarPtr(returnedValue))
    Else
        hResult = DispCallFunc( _
            0, mNativeAddress, CALLCONV_STDCALL, CInt(mReturnType), _
            0, 0, 0, VarPtr(returnedValue))
    End If
    If hResult <> HRESULT_SUCCESS Then
        RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
            "DispCallFunc failed with HRESULT " & CStr(hResult) & "."
    End If

    If mHasReturnSignature And mReturnType = vbObject Then
        Set InvokeNative = CreateObjectResult(returnedValue)
    Else
        Set InvokeNative = CreateScalarResult(returnedValue)
    End If
End Function

Private Sub PrepareInvocationValues( _
    ByVal arguments As Collection, _
    ByRef values() As Variant _
)
    Dim index As Long
    Dim wrapped As ROneCOne

    If arguments.Count = 0 Then Exit Sub
    ReDim values(0 To arguments.Count - 1)
    For index = 0 To arguments.Count - 1
        Set wrapped = arguments.Item(index + 1)
        If wrapped.InternalHasObjectValue Then
            Set values(index) = wrapped.InternalObjectValue
        Else
            values(index) = wrapped.InternalScalarValue
        End If
    Next index
End Sub

Private Function ResolveProcedureName(ByVal procedureName As String) As String
    If InStr(1, procedureName, "!", vbBinaryCompare) > 0 Then
        ResolveProcedureName = procedureName
    Else
        ResolveProcedureName = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
            "'!" & procedureName
    End If
End Function

Private Sub ValidateDelegateArguments(ByVal arguments As Collection)
    Dim argumentValue As ROneCOne
    Dim descriptor As ROneCOne
    Dim index As Long

    If Not mHasParameterSignature Then
        If mRole = ROLE_NATIVE Then
            RaiseContractError ERROR_NATIVE_INVOCATION, "Native", _
                "Native invocation requires a complete Takes signature."
        End If
        Exit Sub
    End If
    If arguments.Count <> mParameters.Count Then
        RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
            "Expected " & CStr(mParameters.Count) & " argument(s), received " & _
            CStr(arguments.Count) & "."
    End If

    For index = 1 To mParameters.Count
        Set descriptor = mParameters.Item(index)
        Set argumentValue = arguments.Item(index)
        If descriptor.InternalIsByRef Then
            If mRole <> ROLE_NATIVE Then
                RaiseContractError ERROR_BYREF_INVOCATION, "Run", _
                    "True ByRef requires a signature-bound native delegate."
            End If
            If argumentValue.InternalRole <> ROLE_REFERENCE Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "Parameter " & CStr(index) & " requires a ROneCOne reference."
            End If
            If argumentValue.InternalExpectedType <> _
                descriptor.InternalExpectedType Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "The ByRef argument type does not match its signature."
            End If
        Else
            If argumentValue.InternalRole = ROLE_REFERENCE Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "A reference was supplied to a ByVal parameter."
            End If
            ValidateParameterValue descriptor, argumentValue
        End If
    Next index
End Sub

Private Function FinalizeDelegateResult(ByVal result As ROneCOne) As ROneCOne
    Dim actualType As VbVarType

    If result Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Run", _
            "The delegate did not produce an invocation result."
    End If
    If mIsActionDelegate Then
        Set FinalizeDelegateResult = CreateScalarResult(Empty)
        Exit Function
    End If
    If Not mHasReturnSignature Or mReturnType = vbVariant Then
        Set FinalizeDelegateResult = result
        Exit Function
    End If

    If mReturnType = vbObject Then
        If Not result.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "The delegate result is not an object."
        End If
        If Len(mReturnObjectType) > 0 Then
            If StrComp(TypeName(result.InternalObjectValue), mReturnObjectType, _
                vbTextCompare) <> 0 Then
                RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                    "The delegate returned the wrong object type."
            End If
        End If
    Else
        If result.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "The delegate result is an object, not a scalar value."
        End If
        actualType = VarType(result.InternalScalarValue)
        If actualType <> mReturnType Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Run", _
                "Expected " & TypeNameForVarType(mReturnType) & ", received " & _
                TypeNameForVarType(actualType) & "."
        End If
    End If
    Set FinalizeDelegateResult = result
End Function

Private Function RunWorkbookProcedure( _
    ByVal procedureName As String, _
    ByRef values() As Variant, _
    ByVal argumentCount As Long _
) As Variant
    Select Case argumentCount
        Case 0
            RunWorkbookProcedure = Application.Run(procedureName)
        Case 1
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0))
        Case 2
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1))
        Case 3
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2))
        Case 4
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3))
        Case 5
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4))
        Case 6
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5))
        Case 7
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6))
        Case 8
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7))
        Case 9
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8))
        Case 10
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9))
        Case 11
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10))
        Case 12
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11))
        Case 13
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12))
        Case 14
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13))
        Case 15
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14))
        Case 16
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15))
        Case 17
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16))
        Case 18
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17))
        Case 19
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18))
        Case 20
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19))
        Case 21
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20))
        Case 22
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21))
        Case 23
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22))
        Case 24
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23))
        Case 25
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24))
        Case 26
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25))
        Case 27
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26))
        Case 28
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27))
        Case 29
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28))
        Case 30
            RunWorkbookProcedure = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28), _
                values(29))
        Case Else
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "Workbook procedures support at most 30 arguments."
    End Select
End Function

Private Function NormalizeUnaryAction( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As ROneCOne
    RequireSequence Me, memberName
    ValidateAction candidate, memberName, False
    If candidate.Arity <> 1 And candidate.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The Action must accept exactly one value."
    End If
    Set NormalizeUnaryAction = candidate
End Function

Private Function RunWorkbookProcedureObject( _
    ByVal procedureName As String, _
    ByRef values() As Variant, _
    ByVal argumentCount As Long _
) As Object
    Select Case argumentCount
        Case 0
            Set RunWorkbookProcedureObject = Application.Run(procedureName)
        Case 1
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0))
        Case 2
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1))
        Case 3
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2))
        Case 4
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3))
        Case 5
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4))
        Case 6
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5))
        Case 7
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6))
        Case 8
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7))
        Case 9
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8))
        Case 10
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9))
        Case 11
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10))
        Case 12
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11))
        Case 13
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12))
        Case 14
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13))
        Case 15
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14))
        Case 16
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15))
        Case 17
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16))
        Case 18
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17))
        Case 19
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18))
        Case 20
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19))
        Case 21
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20))
        Case 22
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21))
        Case 23
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22))
        Case 24
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23))
        Case 25
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24))
        Case 26
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25))
        Case 27
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26))
        Case 28
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27))
        Case 29
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28))
        Case 30
            Set RunWorkbookProcedureObject = Application.Run( _
                procedureName, values(0), values(1), values(2), values(3), values(4), _
                values(5), values(6), values(7), values(8), values(9), values(10), _
                values(11), values(12), values(13), values(14), values(15), values(16), _
                values(17), values(18), values(19), values(20), values(21), values(22), _
                values(23), values(24), values(25), values(26), values(27), values(28), _
                values(29))
        Case Else
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "Workbook procedures support at most 30 arguments."
    End Select
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

Private Function IsMemberOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case OP_MEMBER_VALUE, OP_MEMBER_OBJECT
            IsMemberOperator = True
        Case OP_SAFE_MEMBER_VALUE, OP_SAFE_MEMBER_OBJECT
            IsMemberOperator = True
    End Select
End Function

Private Function IsObjectMemberOperator(ByVal operatorCode As Long) As Boolean
    IsObjectMemberOperator = _
        operatorCode = OP_MEMBER_OBJECT Or _
        operatorCode = OP_SAFE_MEMBER_OBJECT
End Function

Private Function IsSafeMemberOperator(ByVal operatorCode As Long) As Boolean
    IsSafeMemberOperator = _
        operatorCode = OP_SAFE_MEMBER_VALUE Or _
        operatorCode = OP_SAFE_MEMBER_OBJECT
End Function

Private Function IsQuantifierOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case OP_ANY_MATCH, OP_ALL_MATCH, OP_NONE_MATCH
            IsQuantifierOperator = True
    End Select
End Function

Private Function SequenceContainsWrapped( _
    ByVal sequenceValue As ROneCOne, _
    ByVal candidate As ROneCOne, _
    Optional ByVal comparer As ROneCOne _
) As Boolean
    Dim values As Collection
    Dim wrapped As ROneCOne

    RequireSequence sequenceValue, "In"
    Set values = sequenceValue.MaterializeWrapped
    For Each wrapped In values
        If ValuesEqualWithComparer( _
            wrapped, candidate, comparer, "IsIn") Then
            SequenceContainsWrapped = True
            Exit Function
        End If
    Next wrapped
End Function

Private Function EvaluateNestedQuantifier( _
    ByVal nestedValue As ROneCOne, _
    ByVal predicate As ROneCOne, _
    ByVal operatorCode As Long _
) As Boolean
    Dim nestedObject As Object
    Dim sequenceValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If Not nestedValue.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Nested quantifier", _
            "A nested quantifier requires an object-valued sequence member."
    End If
    Set nestedObject = nestedValue.InternalObjectValue
    If nestedObject Is Nothing Then
        EvaluateNestedQuantifier = (operatorCode <> OP_ANY_MATCH)
        Exit Function
    End If
    If Not TypeOf nestedObject Is ROneCOne Then
        RaiseContractError ERROR_TYPE_MISMATCH, "Nested quantifier", _
            "The nested member must return a ROneCOne sequence."
    End If
    Set sequenceValue = nestedObject
    RequireSequence sequenceValue, "Nested quantifier"
    Set values = sequenceValue.MaterializeWrapped

    If operatorCode = OP_ANY_MATCH Then
        If predicate Is Nothing Then
            EvaluateNestedQuantifier = (values.Count > 0)
            Exit Function
        End If
        For Each wrapped In values
            If InvokePredicate(predicate, wrapped, "AnyMatch") Then
                EvaluateNestedQuantifier = True
                Exit Function
            End If
        Next wrapped
        Exit Function
    End If

    If operatorCode = OP_ALL_MATCH Then
        EvaluateNestedQuantifier = True
        For Each wrapped In values
            If Not InvokePredicate(predicate, wrapped, "AllMatch") Then
                EvaluateNestedQuantifier = False
                Exit Function
            End If
        Next wrapped
        Exit Function
    End If

    EvaluateNestedQuantifier = True
    If predicate Is Nothing Then
        EvaluateNestedQuantifier = (values.Count = 0)
        Exit Function
    End If
    For Each wrapped In values
        If InvokePredicate(predicate, wrapped, "NoneMatch") Then
            EvaluateNestedQuantifier = False
            Exit Function
        End If
    Next wrapped
End Function

Private Function ReadMember( _
    ByVal targetValue As ROneCOne, _
    ByVal memberName As String, _
    ByVal returnsObject As Boolean, _
    ByVal nullSafe As Boolean _
) As ROneCOne
    Dim capturedDescription As String
    Dim returnedObject As Object
    Dim returnedValue As Variant
    Dim target As Object

    If Not targetValue.InternalHasObjectValue Then
        RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
            "Member access requires an object value."
    End If
    Set target = targetValue.InternalObjectValue
    If target Is Nothing Then
        If nullSafe Then
            If returnsObject Then
                Set ReadMember = CreateObjectResult(Nothing)
            Else
                Set ReadMember = CreateScalarResult(Null)
            End If
            Exit Function
        End If
        RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
            "Cannot access '" & memberName & "' on Nothing."
    End If

    On Error GoTo AccessFailed
    If returnsObject Then
        Set returnedObject = CallByName(target, memberName, VbGet)
        Set ReadMember = CreateObjectResult(returnedObject)
    Else
        returnedValue = CallByName(target, memberName, VbGet)
        Set ReadMember = CreateScalarResult(returnedValue)
    End If
    On Error GoTo 0
    Exit Function

AccessFailed:
    capturedDescription = Err.Description
    On Error GoTo 0
    RaiseContractError ERROR_MEMBER_ACCESS, "Member", _
        "Cannot read '" & memberName & "' from " & TypeName(target) & _
        ". " & capturedDescription
End Function

Private Function HasNullScalar(ByVal wrapped As ROneCOne) As Boolean
    If wrapped Is Nothing Then Exit Function
    If wrapped.InternalHasObjectValue Then Exit Function
    HasNullScalar = IsNull(wrapped.InternalScalarValue)
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
        If IsNull(leftValue.InternalScalarValue) Or _
            IsNull(rightValue.InternalScalarValue) Then
            ValuesEqual = IsNull(leftValue.InternalScalarValue) And _
                IsNull(rightValue.InternalScalarValue)
        Else
            ValuesEqual = _
                (leftValue.InternalScalarValue = rightValue.InternalScalarValue)
        End If
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
    If IsDelegateRole(candidate.InternalRole) Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires a delegate."
End Sub

Private Sub RequireEvent(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The event cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_EVENT
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires an event."
End Sub

Private Sub RequireTry(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The Try operation cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_TRY
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This operation requires a Try operation."
End Sub

Private Sub RequireError(ByVal candidate As ROneCOne, ByVal context As String)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The exception cannot be Nothing."
    End If
    Select Case candidate.InternalRole
        Case ROLE_ERROR
            Exit Sub
    End Select
    RaiseContractError ERROR_INVALID_OPERATION, context, _
        "This property requires a captured exception."
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
    Set mInvocationList = Nothing
    Set mTryBody = Nothing
    Set mCatchActions = Nothing
    Set mCatchNumbers = Nothing
    Set mFinallyAction = Nothing
    Set mItems = Nothing
    Set mEnumerationValues = Nothing
    Set mElementParameter = Nothing
    Set mSource = Nothing
    Set mQueryDelegate = Nothing
    Set mQueryComparer = Nothing
    Set mQueryValue = Nothing
End Sub

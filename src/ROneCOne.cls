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

Private Const ROLE_EMPTY As Long = 0
Private Const ROLE_VALUE As Long = 1
Private Const ROLE_PARAMETER As Long = 2
Private Const ROLE_EXPRESSION As Long = 3
Private Const ROLE_LAMBDA As Long = 4
Private Const ROLE_METHOD As Long = 5
Private Const ROLE_COMPOSITION As Long = 6

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
Attribute Run.VB_Description = "Runs a lambda, method delegate, or composition."
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
Attribute Add.VB_Description = "Creates an addition expression."
    Set Add = CreateBinaryExpression(OP_ADD, operand)
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
        Case vbBoolean
            TypeNameForVarType = "Boolean"
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
End Sub

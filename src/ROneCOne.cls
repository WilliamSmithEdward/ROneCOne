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

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal length As LongPtr _
)

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
Private Const ROLE_DICTIONARY As Long = 17
Private Const ROLE_HASH_SET As Long = 18
Private Const ROLE_COLLECTION As Long = 19
Private Const ROLE_NODE As Long = 20
Private Const ROLE_IMMUTABLE As Long = 21
Private Const ROLE_CONCURRENT As Long = 22
Private Const ROLE_ENTRY As Long = 23
Private Const ROLE_TASK_FACTORY As Long = 24
Private Const ROLE_TASK As Long = 25
Private Const ROLE_CANCELLATION_SOURCE As Long = 26
Private Const ROLE_CANCELLATION_TOKEN As Long = 27
Private Const ROLE_DATA_TABLE As Long = 28
Private Const ROLE_DATA_COLUMN As Long = 29
Private Const ROLE_DATA_ROW As Long = 30
Private Const ROLE_DATA_SET As Long = 31
Private Const ROLE_DATA_RELATION As Long = 32
Private Const ROLE_DATA_VIEW As Long = 33
Private Const ROLE_DB_CONNECTION As Long = 34
Private Const ROLE_DB_COMMAND As Long = 35
Private Const ROLE_DB_PARAMETER As Long = 36
Private Const ROLE_DB_READER As Long = 37
Private Const ROLE_DB_ADAPTER As Long = 38
Private Const ROLE_DB_TRANSACTION As Long = 39
Private Const ROLE_TASK_COMPLETION_SOURCE As Long = 40
Private Const ROLE_PROGRESS As Long = 41

Private Const TASK_WORK As Long = 1
Private Const TASK_COMPLETED As Long = 2
Private Const TASK_DELAY As Long = 3
Private Const TASK_WHEN_ALL As Long = 4
Private Const TASK_WHEN_ANY As Long = 5
Private Const TASK_CONTINUATION As Long = 6
Private Const TASK_PROMISE As Long = 7
Private Const TASK_PROVIDER_CALL As Long = 8

Private Const TASK_STATUS_CREATED As Long = 0
Private Const TASK_STATUS_RUNNING As Long = 1
Private Const TASK_STATUS_COMPLETED As Long = 2
Private Const TASK_STATUS_CANCELED As Long = 3
Private Const TASK_STATUS_FAULTED As Long = 4

Private Const ROW_STATE_DETACHED As Long = 0
Private Const ROW_STATE_UNCHANGED As Long = 1
Private Const ROW_STATE_ADDED As Long = 2
Private Const ROW_STATE_DELETED As Long = 3
Private Const ROW_STATE_MODIFIED As Long = 4

Private Const COLLECTION_DICTIONARY As Long = 1
Private Const COLLECTION_HASH_SET As Long = 2
Private Const COLLECTION_QUEUE As Long = 3
Private Const COLLECTION_STACK As Long = 4
Private Const COLLECTION_LINKED_LIST As Long = 5
Private Const COLLECTION_SORTED_LIST As Long = 6
Private Const COLLECTION_SORTED_DICTIONARY As Long = 7
Private Const COLLECTION_SORTED_SET As Long = 8
Private Const COLLECTION_ORDERED_DICTIONARY As Long = 9
Private Const COLLECTION_PRIORITY_QUEUE As Long = 10
Private Const COLLECTION_OBSERVABLE As Long = 11
Private Const COLLECTION_READ_ONLY As Long = 12
Private Const COLLECTION_KEYED As Long = 13
Private Const COLLECTION_CONCURRENT_DICTIONARY As Long = 14
Private Const COLLECTION_CONCURRENT_QUEUE As Long = 15
Private Const COLLECTION_CONCURRENT_STACK As Long = 16
Private Const COLLECTION_CONCURRENT_BAG As Long = 17
Private Const COLLECTION_BLOCKING As Long = 18
Private Const COLLECTION_IMMUTABLE_ARRAY As Long = 19
Private Const COLLECTION_IMMUTABLE_LIST As Long = 20
Private Const COLLECTION_IMMUTABLE_DICTIONARY As Long = 21
Private Const COLLECTION_IMMUTABLE_HASH_SET As Long = 22
Private Const COLLECTION_IMMUTABLE_QUEUE As Long = 23
Private Const COLLECTION_IMMUTABLE_STACK As Long = 24
Private Const COLLECTION_IMMUTABLE_SORTED_DICTIONARY As Long = 25
Private Const COLLECTION_IMMUTABLE_SORTED_SET As Long = 26
Private Const COLLECTION_LOOKUP As Long = 27
Private Const COLLECTION_GROUPING As Long = 28

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
Private Const QUERY_ORDER As Long = 6
Private Const QUERY_ORDER_DESCENDING As Long = 7
Private Const QUERY_ORDER_BY As Long = 8
Private Const QUERY_ORDER_BY_DESCENDING As Long = 9
Private Const QUERY_THEN_BY As Long = 10
Private Const QUERY_THEN_BY_DESCENDING As Long = 11
Private Const QUERY_APPEND As Long = 12
Private Const QUERY_PREPEND As Long = 13
Private Const QUERY_REVERSE As Long = 14
Private Const QUERY_DISTINCT_BY As Long = 15
Private Const QUERY_WHERE_NOT As Long = 16
Private Const QUERY_TAKE_WHILE As Long = 17
Private Const QUERY_SKIP_WHILE As Long = 18
Private Const QUERY_TAKE_LAST As Long = 19
Private Const QUERY_SKIP_LAST As Long = 20
Private Const QUERY_CONCAT As Long = 21
Private Const QUERY_UNION As Long = 22
Private Const QUERY_INTERSECT As Long = 23
Private Const QUERY_EXCEPT As Long = 24
Private Const QUERY_DEFAULT_IF_EMPTY As Long = 25
Private Const QUERY_CHUNK As Long = 26
Private Const QUERY_SELECT_MANY As Long = 27
Private Const QUERY_CAST As Long = 28
Private Const QUERY_OF_TYPE As Long = 29
Private Const QUERY_UNION_BY As Long = 30
Private Const QUERY_INTERSECT_BY As Long = 31
Private Const QUERY_EXCEPT_BY As Long = 32

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
Private mCollectionKind As Long
Private mKeys As Collection
Private mPriorities As Collection
Private mKeyType As VbVarType
Private mKeyObjectType As String
Private mKeySelector As ROneCOne
Private mCollectionChanged As ROneCOne
Private mNodePrevious As ROneCOne
Private mNodeNext As ROneCOne
Private mIsBuilder As Boolean
Private mAddingCompleted As Boolean
Private mTaskBody As ROneCOne
Private mTaskResult As ROneCOne
Private mTaskException As ROneCOne
Private mTaskChildren As Collection
Private mTaskStatus As Long
Private mCancellationSource As ROneCOne
Private mCancellationRequested As Boolean
Private mName As String
Private mOriginalItems As Collection
Private mPrimaryKey As Collection
Private mDataTables As Collection
Private mDataRelations As Collection
Private mAllowDBNull As Boolean
Private mUnique As Boolean
Private mAutoIncrement As Boolean
Private mAutoIncrementSeed As LongLong
Private mAutoIncrementStep As LongLong
Private mAutoIncrementCurrent As LongLong
Private mDefaultValue As ROneCOne
Private mExpressionText As String
Private mRowState As Long
Private mProviderObject As Object
Private mConnectionString As String
Private mCommandText As String
Private mProviderName As String
Private mProviderType As Long
Private mProviderSize As Long
Private mProviderPrecision As Byte
Private mProviderScale As Byte
Private mSourceColumn As String
Private mSourceVersion As String
Private mCommandTimeout As Long
Private mCommandType As Long
Private mConnectionTimeout As Long
Private mTransactionActive As Boolean

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

Public Function Value(Optional ByVal capturedValue As Variant) As Variant
Attribute Value.VB_Description = "Captures a scalar or object as an expression value."
    Dim node As ROneCOne
    Dim capturedObject As Object

    If mRole = ROLE_NODE Or mRole = ROLE_ENTRY Then
        If Not IsMissing(capturedValue) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
                "This value does not accept an argument."
        End If
        AssignPublicResult Value, mQueryValue
        Exit Function
    End If
    If mRole <> ROLE_EMPTY Then
        RaiseContractError ERROR_INVALID_OPERATION, "Value", _
            "Value creation must start from the ROneCOne factory."
    End If
    If IsMissing(capturedValue) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Value", _
            "A captured value is required."
    End If

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

Public Function DictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute DictionaryOf.VB_Description = "Creates a Dictionary<TKey, TValue>."
    Set DictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_DICTIONARY, keyType, valueType, comparer)
End Function

Public Function HashSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute HashSetOf.VB_Description = "Creates a HashSet<T>."
    Set HashSetOf = CreateSequenceCollection( _
        ROLE_HASH_SET, COLLECTION_HASH_SET, typeToken, comparer)
End Function

Public Function QueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute QueueOf.VB_Description = "Creates a Queue<T>."
    Set QueueOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_QUEUE, typeToken)
End Function

Public Function StackOf(ByVal typeToken As Variant) As ROneCOne
Attribute StackOf.VB_Description = "Creates a Stack<T>."
    Set StackOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_STACK, typeToken)
End Function

Public Function LinkedListOf(ByVal typeToken As Variant) As ROneCOne
Attribute LinkedListOf.VB_Description = "Creates a LinkedList<T>."
    Set LinkedListOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_LINKED_LIST, typeToken)
End Function

Public Function SortedListOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedListOf.VB_Description = "Creates a SortedList<TKey, TValue>."
    Set SortedListOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_SORTED_LIST, keyType, valueType, comparer)
End Function

Public Function SortedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedDictionaryOf.VB_Description = _
    "Creates a SortedDictionary<TKey, TValue>."
    Set SortedDictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_SORTED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function SortedSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute SortedSetOf.VB_Description = "Creates a SortedSet<T>."
    Set SortedSetOf = CreateSequenceCollection( _
        ROLE_HASH_SET, COLLECTION_SORTED_SET, typeToken, comparer)
End Function

Public Function OrderedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute OrderedDictionaryOf.VB_Description = _
    "Creates an OrderedDictionary<TKey, TValue>."
    Set OrderedDictionaryOf = CreateMapCollection( _
        ROLE_DICTIONARY, COLLECTION_ORDERED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function PriorityQueueOf( _
    ByVal elementType As Variant, _
    ByVal priorityType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute PriorityQueueOf.VB_Description = _
    "Creates a PriorityQueue<TElement, TPriority>."
    Set PriorityQueueOf = CreateMapCollection( _
        ROLE_COLLECTION, COLLECTION_PRIORITY_QUEUE, _
        priorityType, elementType, comparer)
End Function

Public Function ObservableCollectionOf(ByVal typeToken As Variant) As ROneCOne
Attribute ObservableCollectionOf.VB_Description = _
    "Creates an ObservableCollection<T>."
    Set ObservableCollectionOf = CreateSequenceCollection( _
        ROLE_COLLECTION, COLLECTION_OBSERVABLE, typeToken)
    ObservableCollectionOf.ConfigureObservableCollection
End Function

Public Function ReadOnlyCollectionOf(ByVal source As ROneCOne) As ROneCOne
Attribute ReadOnlyCollectionOf.VB_Description = _
    "Creates a live ReadOnlyCollection<T> view."
    Dim result As ROneCOne

    RequireSequence source, "ReadOnlyCollectionOf"
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_READ_ONLY, _
        source.InternalElementType, source.InternalElementObjectType
    result.ConfigureReadOnlySource source
    Set ReadOnlyCollectionOf = result
End Function

Public Function KeyedCollectionOf( _
    ByVal keyType As Variant, _
    ByVal itemType As Variant, _
    ByVal keySelector As ROneCOne, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute KeyedCollectionOf.VB_Description = _
    "Creates a KeyedCollection<TKey, TItem>."
    Dim result As ROneCOne

    Set result = CreateMapCollection( _
        ROLE_COLLECTION, COLLECTION_KEYED, keyType, itemType, comparer)
    result.ConfigureCollectionKeySelector NormalizeStandaloneUnary( _
        keySelector, "KeyedCollectionOf")
    Set KeyedCollectionOf = result
End Function

Public Function ConcurrentDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ConcurrentDictionaryOf.VB_Description = _
    "Creates a ConcurrentDictionary<TKey, TValue>."
    Set ConcurrentDictionaryOf = CreateMapCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ConcurrentQueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentQueueOf.VB_Description = "Creates a ConcurrentQueue<T>."
    Set ConcurrentQueueOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_QUEUE, typeToken)
End Function

Public Function ConcurrentStackOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentStackOf.VB_Description = "Creates a ConcurrentStack<T>."
    Set ConcurrentStackOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_STACK, typeToken)
End Function

Public Function ConcurrentBagOf(ByVal typeToken As Variant) As ROneCOne
Attribute ConcurrentBagOf.VB_Description = "Creates a ConcurrentBag<T>."
    Set ConcurrentBagOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_CONCURRENT_BAG, typeToken)
End Function

Public Function BlockingCollectionOf( _
    ByVal typeToken As Variant, _
    Optional ByVal boundedCapacity As Long = -1 _
) As ROneCOne
Attribute BlockingCollectionOf.VB_Description = _
    "Creates a cooperative BlockingCollection<T>."
    Set BlockingCollectionOf = CreateSequenceCollection( _
        ROLE_CONCURRENT, COLLECTION_BLOCKING, typeToken)
    BlockingCollectionOf.ConfigureCollectionCapacity boundedCapacity
End Function

Public Function ImmutableArrayOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableArrayOf.VB_Description = "Creates an empty ImmutableArray<T>."
    Set ImmutableArrayOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_ARRAY, typeToken)
End Function

Public Function ImmutableListOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableListOf.VB_Description = "Creates an empty ImmutableList<T>."
    Set ImmutableListOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_LIST, typeToken)
End Function

Public Function ImmutableDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableDictionaryOf.VB_Description = _
    "Creates an ImmutableDictionary<TKey, TValue>."
    Set ImmutableDictionaryOf = CreateMapCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ImmutableHashSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableHashSetOf.VB_Description = _
    "Creates an ImmutableHashSet<T>."
    Set ImmutableHashSetOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_HASH_SET, typeToken, comparer)
End Function

Public Function ImmutableQueueOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableQueueOf.VB_Description = "Creates an empty ImmutableQueue<T>."
    Set ImmutableQueueOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_QUEUE, typeToken)
End Function

Public Function ImmutableStackOf(ByVal typeToken As Variant) As ROneCOne
Attribute ImmutableStackOf.VB_Description = "Creates an empty ImmutableStack<T>."
    Set ImmutableStackOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_STACK, typeToken)
End Function

Public Function ImmutableSortedDictionaryOf( _
    ByVal keyType As Variant, _
    ByVal valueType As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableSortedDictionaryOf.VB_Description = _
    "Creates an ImmutableSortedDictionary<TKey, TValue>."
    Set ImmutableSortedDictionaryOf = CreateMapCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_SORTED_DICTIONARY, _
        keyType, valueType, comparer)
End Function

Public Function ImmutableSortedSetOf( _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ImmutableSortedSetOf.VB_Description = _
    "Creates an ImmutableSortedSet<T>."
    Set ImmutableSortedSetOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_SORTED_SET, typeToken, comparer)
End Function

Public Function [Empty](ByVal typeToken As Variant) As ROneCOne
Attribute Empty.VB_Description = "Creates an empty typed enumerable."
    Set [Empty] = EmptyOf(typeToken)
End Function

Public Function EmptyOf(ByVal typeToken As Variant) As ROneCOne
Attribute EmptyOf.VB_Description = "Creates an empty typed enumerable in legal VBA syntax."
    Set EmptyOf = CreateSequenceCollection( _
        ROLE_IMMUTABLE, COLLECTION_IMMUTABLE_ARRAY, typeToken)
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

    If mRole = ROLE_TASK Then
        EnsureTaskCompleted
        Set Exception = mTaskException
        Exit Property
    End If

    Set errorValue = New ROneCOne
    errorValue.ConfigureError 0, vbNullString, vbNullString, vbNullString, 0
    Set Exception = errorValue
End Property

Public Property Get Task() As ROneCOne
Attribute Task.VB_Description = "Gets the C#-style Task factory surface."
    Dim factory As ROneCOne

    If mRole = ROLE_TASK_COMPLETION_SOURCE Then
        Set Task = mSource
        Exit Property
    End If
    Set factory = New ROneCOne
    factory.ConfigureTaskFactory
    Set Task = factory
End Property

Public Property Get CompletedTask() As ROneCOne
Attribute CompletedTask.VB_Description = "Gets a successfully completed Task."
    Set CompletedTask = TaskFromResult(Empty)
End Property

Public Function FromResult(ByVal value As Variant) As ROneCOne
Attribute FromResult.VB_Description = "Creates a completed Task<T> from a value."
    Set FromResult = TaskFromResult(value)
End Function

Public Function TaskRun( _
    ByVal work As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute TaskRun.VB_Description = "Creates a cooperatively scheduled Task."
    Dim taskValue As ROneCOne

    ValidateTaskWork work, "TaskRun"
    ValidateCancellationToken cancellationToken, "TaskRun"
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_WORK, work, cancellationToken
    Set TaskRun = taskValue
End Function

Public Function TaskFromResult(ByVal value As Variant) As ROneCOne
Attribute TaskFromResult.VB_Description = "Creates a completed Task<T> from a value."
    Dim taskValue As ROneCOne

    Set taskValue = New ROneCOne
    taskValue.ConfigureCompletedTask WrapInvocationArgument(value)
    Set TaskFromResult = taskValue
End Function

Public Function Delay( _
    ByVal milliseconds As Long, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute Delay.VB_Description = "Creates a cooperative delay Task."
    Dim taskValue As ROneCOne

    If milliseconds < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Delay", _
            "The delay cannot be negative."
    End If
    ValidateCancellationToken cancellationToken, "Delay"
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_DELAY, Nothing, cancellationToken
    taskValue.ConfigureCollectionCapacity milliseconds
    Set Delay = taskValue
End Function

Public Function WhenAll(ByVal tasks As Variant) As ROneCOne
Attribute WhenAll.VB_Description = "Creates a Task that completes after every input Task."
    Set WhenAll = CreateTaskCombinator(tasks, TASK_WHEN_ALL, "WhenAll")
End Function

Public Function WhenAny(ByVal tasks As Variant) As ROneCOne
Attribute WhenAny.VB_Description = "Creates a Task that completes after one input Task."
    Set WhenAny = CreateTaskCombinator(tasks, TASK_WHEN_ANY, "WhenAny")
End Function

Public Function Await() As Variant
Attribute Await.VB_Description = "Cooperatively awaits completion and returns the Task result."
    RequireTask Me, "Await"
    EnsureTaskCompleted
    RaiseTaskTerminalState "Await"
    AssignPublicResult Await, mTaskResult
End Function

Public Function Wait(Optional ByVal timeoutMilliseconds As Long = -1) As Boolean
Attribute Wait.VB_Description = "Waits cooperatively for Task completion."
    RequireTask Me, "Wait"
    If timeoutMilliseconds < -1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Wait", _
            "The timeout must be -1 or nonnegative."
    End If
    If Not WaitForTaskCompletion(timeoutMilliseconds) Then Exit Function
    RaiseTaskTerminalState "Wait"
    Wait = True
End Function

Public Function ContinueWith(ByVal continuation As ROneCOne) As ROneCOne
Attribute ContinueWith.VB_Description = "Creates a continuation Task."
    Dim taskValue As ROneCOne

    RequireTask Me, "ContinueWith"
    If continuation Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ContinueWith", _
            "The continuation cannot be Nothing."
    End If
    RequireDelegate continuation, "ContinueWith"
    If continuation.Arity <> 1 And continuation.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "ContinueWith", _
            "The continuation must accept the antecedent Task."
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureTask TASK_CONTINUATION, continuation, Nothing
    taskValue.ConfigureReadOnlySource Me
    Set ContinueWith = taskValue
End Function

Public Property Get Status() As String
Attribute Status.VB_Description = "Gets the .NET-aligned Task status name."
    RequireTask Me, "Status"
    Select Case mTaskStatus
        Case TASK_STATUS_CREATED
            Status = "Created"
        Case TASK_STATUS_RUNNING
            Status = "Running"
        Case TASK_STATUS_COMPLETED
            Status = "RanToCompletion"
        Case TASK_STATUS_CANCELED
            Status = "Canceled"
        Case TASK_STATUS_FAULTED
            Status = "Faulted"
    End Select
End Property

Public Property Get IsCompleted() As Boolean
Attribute IsCompleted.VB_Description = "Reports terminal Task or blocking-collection state."
    If mRole = ROLE_TASK Then
        IsCompleted = (mTaskStatus >= TASK_STATUS_COMPLETED)
    ElseIf mCollectionKind = COLLECTION_BLOCKING Then
        IsCompleted = (mAddingCompleted And mItems.Count = 0)
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "IsCompleted", _
            "IsCompleted requires a Task or BlockingCollection."
    End If
End Property

Public Property Get IsCanceled() As Boolean
Attribute IsCanceled.VB_Description = "Reports whether a Task was canceled."
    RequireTask Me, "IsCanceled"
    IsCanceled = (mTaskStatus = TASK_STATUS_CANCELED)
End Property

Public Property Get IsFaulted() As Boolean
Attribute IsFaulted.VB_Description = "Reports whether a Task faulted."
    RequireTask Me, "IsFaulted"
    IsFaulted = (mTaskStatus = TASK_STATUS_FAULTED)
End Property

Public Property Get Result() As Variant
Attribute Result.VB_Description = "Waits for and returns a Task<T> result."
    RequireTask Me, "Result"
    EnsureTaskCompleted
    RaiseTaskTerminalState "Result"
    AssignPublicResult Result, mTaskResult
End Property

Public Function CancellationTokenSource() As ROneCOne
Attribute CancellationTokenSource.VB_Description = "Creates a CancellationTokenSource."
    Dim source As ROneCOne

    Set source = New ROneCOne
    source.ConfigureCancellationSource
    Set CancellationTokenSource = source
End Function

Public Function TaskCompletionSourceOf(ByVal typeToken As Variant) As ROneCOne
Attribute TaskCompletionSourceOf.VB_Description = _
    "Creates a typed cooperative TaskCompletionSource<T>."
    Dim objectType As String
    Dim source As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, _
        "TaskCompletionSourceOf"
    Set source = New ROneCOne
    source.ConfigureTaskCompletionSource valueType, objectType
    Set TaskCompletionSourceOf = source
End Function

Public Function ProgressOf( _
    ByVal typeToken As Variant, _
    ByVal handler As ROneCOne _
) As ROneCOne
Attribute ProgressOf.VB_Description = "Creates a typed synchronous progress reporter."
    Dim objectType As String
    Dim progressValue As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "ProgressOf"
    ValidateProgressHandler handler, valueType, objectType
    Set progressValue = New ROneCOne
    progressValue.ConfigureProgress valueType, objectType, handler
    Set ProgressOf = progressValue
End Function

Public Property Get Token() As ROneCOne
Attribute Token.VB_Description = "Gets the token owned by a CancellationTokenSource."
    Dim tokenValue As ROneCOne

    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "Token", _
            "Token requires a CancellationTokenSource."
    End If
    Set tokenValue = New ROneCOne
    tokenValue.ConfigureCancellationToken Me
    Set Token = tokenValue
End Property

Public Sub Cancel()
Attribute Cancel.VB_Description = "Requests cancellation and runs registered callbacks."
    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "Cancel", _
            "Cancel requires a CancellationTokenSource."
    End If
    CancelCancellationSource
End Sub

Public Sub CancelAfter(ByVal milliseconds As Long)
Attribute CancelAfter.VB_Description = "Schedules cooperative cancellation after a delay."
    If mRole <> ROLE_CANCELLATION_SOURCE Then
        RaiseContractError ERROR_INVALID_OPERATION, "CancelAfter", _
            "CancelAfter requires a CancellationTokenSource."
    End If
    If milliseconds < -1 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "CancelAfter", _
            "The delay must be -1 or nonnegative."
    End If
    mQueryCount = milliseconds
    If milliseconds >= 0 Then mScalarValue = Timer
End Sub

Public Property Get IsCancellationRequested() As Boolean
Attribute IsCancellationRequested.VB_Description = "Reports whether cancellation was requested."
    If mRole = ROLE_CANCELLATION_SOURCE Then
        EvaluateScheduledCancellation
        IsCancellationRequested = mCancellationRequested
    ElseIf mRole = ROLE_CANCELLATION_TOKEN Then
        IsCancellationRequested = mCancellationSource.IsCancellationRequested
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "IsCancellationRequested", _
            "This property requires a cancellation source or token."
    End If
End Property

Public Function Register(ByVal callback As ROneCOne) As ROneCOne
Attribute Register.VB_Description = "Registers a callback for cooperative cancellation."
    Dim source As ROneCOne

    If mRole = ROLE_CANCELLATION_SOURCE Then
        Set source = Me
    ElseIf mRole = ROLE_CANCELLATION_TOKEN Then
        Set source = mCancellationSource
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Register", _
            "Register requires a cancellation source or token."
    End If
    ValidateAction callback, "Register", True
    source.RegisterCancellationCallback callback
    Set Register = Me
End Function

Public Sub ThrowIfCancellationRequested()
Attribute ThrowIfCancellationRequested.VB_Description = _
    "Raises when cooperative cancellation was requested."
    If Not IsCancellationRequested Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, _
        "ThrowIfCancellationRequested", "Cancellation was requested."
End Sub

Public Function TrySetResult(ByVal value As Variant) As Boolean
Attribute TrySetResult.VB_Description = "Attempts to complete a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetResult"
    TrySetResult = mSource.TryCompletePromise( _
        WrapTypedCompletionValue(value), TASK_STATUS_COMPLETED, Nothing)
End Function

Public Sub SetResult(ByVal value As Variant)
Attribute SetResult.VB_Description = "Completes a TaskCompletionSource<T> successfully."
    If Not TrySetResult(value) Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetResult", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Function TrySetCanceled() As Boolean
Attribute TrySetCanceled.VB_Description = "Attempts to cancel a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetCanceled"
    TrySetCanceled = mSource.TryCompletePromise( _
        Nothing, TASK_STATUS_CANCELED, Nothing)
End Function

Public Sub SetCanceled()
Attribute SetCanceled.VB_Description = "Cancels a TaskCompletionSource<T>."
    If Not TrySetCanceled Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetCanceled", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Function TrySetException(ByVal errorValue As ROneCOne) As Boolean
Attribute TrySetException.VB_Description = _
    "Attempts to fault a TaskCompletionSource<T>."
    RequireTaskCompletionSource "TrySetException"
    If errorValue Is Nothing Or errorValue.InternalRole <> ROLE_ERROR Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "TrySetException", _
            "TrySetException requires a ROneCOne exception value."
    End If
    TrySetException = mSource.TryCompletePromise( _
        Nothing, TASK_STATUS_FAULTED, errorValue)
End Function

Public Sub SetException(ByVal errorValue As ROneCOne)
Attribute SetException.VB_Description = "Faults a TaskCompletionSource<T>."
    If Not TrySetException(errorValue) Then
        RaiseContractError ERROR_INVALID_OPERATION, "SetException", _
            "The TaskCompletionSource is already complete."
    End If
End Sub

Public Sub Report(ByVal value As Variant)
Attribute Report.VB_Description = "Reports one typed progress value."
    Dim arguments As Collection
    Dim ignored As ROneCOne
    Dim wrapped As ROneCOne

    If mRole <> ROLE_PROGRESS Then
        RaiseContractError ERROR_INVALID_OPERATION, "Report", _
            "Report requires a Progress<T> value."
    End If
    Set wrapped = WrapInvocationArgument(value)
    ValidateWrappedType wrapped, mExpectedType, mExpectedObjectType, _
        "Report", "progress value"
    Set arguments = New Collection
    arguments.Add wrapped
    Set ignored = mTaskBody.InvokeValue(arguments)
End Sub

Public Function DataTable(Optional ByVal tableName As String = vbNullString) As ROneCOne
Attribute DataTable.VB_Description = "Creates an in-memory DataTable."
    Dim tableValue As ROneCOne

    Set tableValue = New ROneCOne
    tableValue.ConfigureDataTable tableName
    Set DataTable = tableValue
End Function

Public Function DataColumn( _
    ByVal columnName As String, _
    ByVal typeToken As Variant _
) As ROneCOne
Attribute DataColumn.VB_Description = "Creates a typed DataColumn."
    Dim columnValue As ROneCOne
    Dim objectType As String
    Dim valueType As VbVarType

    columnName = Trim$(columnName)
    If Len(columnName) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataColumn", _
            "The column name cannot be empty."
    End If
    ResolveTypeDescriptor typeToken, valueType, objectType, "DataColumn"
    Set columnValue = New ROneCOne
    columnValue.ConfigureDataColumn columnName, valueType, objectType
    Set DataColumn = columnValue
End Function

Public Function DataSet(Optional ByVal dataSetName As String = vbNullString) As ROneCOne
Attribute DataSet.VB_Description = "Creates an in-memory relational DataSet."
    Dim dataValue As ROneCOne

    Set dataValue = New ROneCOne
    dataValue.ConfigureDataSet dataSetName
    Set DataSet = dataValue
End Function

Public Function DataRelation( _
    ByVal relationName As String, _
    ByVal parentColumn As ROneCOne, _
    ByVal childColumn As ROneCOne, _
    Optional ByVal createConstraints As Boolean = True _
) As ROneCOne
Attribute DataRelation.VB_Description = "Creates a parent-child DataRelation."
    Dim relationValue As ROneCOne

    RequireDataColumn parentColumn, "DataRelation parent"
    RequireDataColumn childColumn, "DataRelation child"
    If parentColumn.InternalExpectedType <> childColumn.InternalExpectedType Or _
        parentColumn.InternalExpectedObjectType <> _
            childColumn.InternalExpectedObjectType Then
        RaiseContractError ERROR_TYPE_MISMATCH, "DataRelation", _
            "Parent and child columns must have the same data type."
    End If
    Set relationValue = New ROneCOne
    relationValue.ConfigureDataRelation relationName, parentColumn, _
        childColumn, createConstraints
    Set DataRelation = relationValue
End Function

Public Function DataView(ByVal table As ROneCOne) As ROneCOne
Attribute DataView.VB_Description = "Creates a live DataView over a DataTable."
    Dim viewValue As ROneCOne

    RequireDataTable table, "DataView"
    Set viewValue = New ROneCOne
    viewValue.ConfigureDataView table
    Set DataView = viewValue
End Function

Public Function AsUnique( _
    Optional ByVal allowNull As Boolean = False _
) As ROneCOne
Attribute AsUnique.VB_Description = "Marks a DataColumn as unique and fluent."
    RequireDataColumn Me, "AsUnique"
    mUnique = True
    If Not allowNull Then mAllowDBNull = False
    Set AsUnique = Me
End Function

Public Function WithDefault(ByVal value As Variant) As ROneCOne
Attribute WithDefault.VB_Description = "Sets a DataColumn default value fluently."
    Dim wrapped As ROneCOne

    RequireDataColumn Me, "WithDefault"
    Set wrapped = WrapDataValue(value)
    ValidateDataColumnValue Me, wrapped, "WithDefault"
    Set mDefaultValue = wrapped
    Set WithDefault = Me
End Function

Public Function AutoNumber( _
    Optional ByVal seed As LongLong = 0, _
    Optional ByVal stepValue As LongLong = 1 _
) As ROneCOne
Attribute AutoNumber.VB_Description = "Configures an auto-increment integer DataColumn."
    RequireDataColumn Me, "AutoNumber"
    Select Case mExpectedType
        Case vbByte, vbInteger, vbLong, vbLongLong
        Case Else
            RaiseContractError ERROR_TYPE_MISMATCH, "AutoNumber", _
                "AutoNumber requires an integer DataColumn."
    End Select
    If stepValue = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AutoNumber", _
            "The auto-increment step cannot be zero."
    End If
    mAutoIncrement = True
    mAutoIncrementSeed = seed
    mAutoIncrementStep = stepValue
    mAutoIncrementCurrent = seed
    mAllowDBNull = False
    Set AutoNumber = Me
End Function

Public Property Get AutoIncrement() As Boolean
Attribute AutoIncrement.VB_Description = "Reports whether a DataColumn auto-increments."
    RequireDataColumn Me, "AutoIncrement"
    AutoIncrement = mAutoIncrement
End Property

Public Property Get DefaultValue() As Variant
Attribute DefaultValue.VB_Description = "Gets a DataColumn default value."
    RequireDataColumn Me, "DefaultValue"
    If mDefaultValue Is Nothing Then
        DefaultValue = Empty
    Else
        AssignPublicResult DefaultValue, mDefaultValue
    End If
End Property

Public Property Get ColumnName() As String
Attribute ColumnName.VB_Description = "Gets a DataColumn name."
    RequireDataColumn Me, "ColumnName"
    ColumnName = mName
End Property

Public Property Let ColumnName(ByVal value As String)
    RequireDataColumn Me, "ColumnName"
    value = Trim$(value)
    If Len(value) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "ColumnName", _
            "The column name cannot be empty."
    End If
    mName = value
End Property

Public Property Get TableName() As String
Attribute TableName.VB_Description = "Gets a DataTable name."
    RequireDataTable Me, "TableName"
    TableName = mName
End Property

Public Property Get DataSetName() As String
Attribute DataSetName.VB_Description = "Gets a DataSet name."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "DataSetName", _
            "DataSetName requires a DataSet."
    End If
    DataSetName = mName
End Property

Public Property Get EnforceConstraints() As Boolean
Attribute EnforceConstraints.VB_Description = "Gets or sets DataSet constraint enforcement."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "EnforceConstraints", _
            "EnforceConstraints requires a DataSet."
    End If
    EnforceConstraints = mUnique
End Property

Public Property Let EnforceConstraints(ByVal value As Boolean)
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "EnforceConstraints", _
            "EnforceConstraints requires a DataSet."
    End If
    mUnique = value
End Property

Public Property Get DataTypeName() As String
Attribute DataTypeName.VB_Description = "Gets a DataColumn generic type name."
    RequireDataColumn Me, "DataTypeName"
    If Len(mExpectedObjectType) > 0 Then
        DataTypeName = mExpectedObjectType
    Else
        DataTypeName = TypeNameForVarType(mExpectedType)
    End If
End Property

Public Property Get AllowDBNull() As Boolean
Attribute AllowDBNull.VB_Description = "Gets or sets whether a DataColumn accepts Null."
    RequireDataColumn Me, "AllowDBNull"
    AllowDBNull = mAllowDBNull
End Property

Public Property Let AllowDBNull(ByVal value As Boolean)
    RequireDataColumn Me, "AllowDBNull"
    mAllowDBNull = value
End Property

Public Property Get Unique() As Boolean
Attribute Unique.VB_Description = "Gets or sets DataColumn uniqueness."
    RequireDataColumn Me, "Unique"
    Unique = mUnique
End Property

Public Property Let Unique(ByVal value As Boolean)
    RequireDataColumn Me, "Unique"
    mUnique = value
End Property

Public Sub AddColumn(ByVal column As ROneCOne)
Attribute AddColumn.VB_Description = "Adds a DataColumn to a DataTable schema."
    RequireDataTable Me, "AddColumn"
    RequireDataColumn column, "AddColumn"
    If FindNamedObjectIndex(mKeys, column.ColumnName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddColumn", _
            "A column with the same name already exists."
    End If
    If Not column.InternalSource Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddColumn", _
            "The DataColumn already belongs to a table."
    End If
    column.ConfigureReadOnlySource Me
    mKeys.Add column
End Sub

Public Function Column( _
    ByVal columnName As String, _
    ByVal typeToken As Variant _
) As ROneCOne
Attribute Column.VB_Description = "Defines and returns a typed DataColumn."
    Dim columnValue As ROneCOne

    RequireDataTable Me, "Column"
    Set columnValue = ROneCOne.DataColumn(columnName, typeToken)
    AddColumn columnValue
    Set Column = columnValue
End Function

Public Function NewRow() As ROneCOne
Attribute NewRow.VB_Description = "Creates a detached DataRow for a DataTable."
    Dim rowValue As ROneCOne

    RequireDataTable Me, "NewRow"
    Set rowValue = New ROneCOne
    rowValue.ConfigureDataRow Me
    Set NewRow = rowValue
End Function

Public Sub AddRow(ByVal row As ROneCOne)
Attribute AddRow.VB_Description = "Adds a detached DataRow after validating constraints."
    RequireDataTable Me, "AddRow"
    RequireDataRow row, "AddRow"
    If Not row.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
            "The DataRow was created by a different table."
    End If
    If row.RowState <> "Detached" Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
            "Only a detached DataRow can be added."
    End If
    ValidateDataRowConstraints row, Nothing
    ValidateForeignKeyConstraints row
    mItems.Add row
    row.ConfigureRowState ROW_STATE_ADDED
    mVersion = mVersion + 1
End Sub

Public Function LoadRow(ByVal values As Variant) As ROneCOne
Attribute LoadRow.VB_Description = "Creates and adds a DataRow from an array or sequence."
    Dim index As Long
    Dim row As ROneCOne
    Dim wrappedValues As Collection

    RequireDataTable Me, "LoadRow"
    Set wrappedValues = BuildDynamicArguments(values)
    If wrappedValues.Count > mKeys.Count Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "LoadRow", _
            "The row contains more values than the table has columns."
    End If
    Set row = NewRow
    For index = 1 To wrappedValues.Count
        row.Item(index - 1) = Unwrap(wrappedValues.Item(index))
    Next index
    AddRow row
    Set LoadRow = row
End Function

Public Function SelectRows(ByVal predicate As Variant) As ROneCOne
Attribute SelectRows.VB_Description = "Returns DataRows matching a typed predicate."
    RequireDataTable Me, "SelectRows"
    Set SelectRows = Rows.Where(predicate).ToList
End Function

Public Function CloneTable() As ROneCOne
Attribute CloneTable.VB_Description = "Copies a DataTable schema without rows."
    RequireDataTable Me, "CloneTable"
    Set CloneTable = CloneDataTableSchema(Me)
End Function

Public Function Copy() As ROneCOne
Attribute Copy.VB_Description = "Copies a DataTable schema and current rows."
    Dim result As ROneCOne
    Dim row As ROneCOne

    RequireDataTable Me, "Copy"
    Set result = CloneDataTableSchema(Me)
    For Each row In mItems
        If row.RowState <> "Deleted" Then CopyDataRowInto row, result
    Next row
    result.AcceptChanges
    Set Copy = result
End Function

Public Sub Merge(ByVal sourceTable As ROneCOne)
Attribute Merge.VB_Description = "Merges compatible rows into this DataTable."
    Dim row As ROneCOne

    RequireDataTable Me, "Merge"
    RequireDataTable sourceTable, "Merge"
    ValidateCompatibleDataTable sourceTable, "Merge"
    For Each row In sourceTable.InternalDataRows
        If row.RowState <> "Deleted" Then CopyDataRowInto row, Me
    Next row
End Sub

Public Function GetChanges() As ROneCOne
Attribute GetChanges.VB_Description = "Copies rows that have pending changes."
    Dim result As ROneCOne
    Dim row As ROneCOne
    Dim table As ROneCOne

    If mRole = ROLE_DATA_SET Then
        Set result = ROneCOne.DataSet(mName)
        For Each table In mDataTables
            Set row = table.GetChanges
            If row.Rows.Count > 0 Then result.AddTable row
        Next table
        Set GetChanges = result
        Exit Function
    End If
    RequireDataTable Me, "GetChanges"
    Set result = CloneDataTableSchema(Me)
    For Each row In mItems
        If row.RowState <> "Unchanged" Then CopyDataRowInto row, result
    Next row
    Set GetChanges = result
End Function

Public Sub Delete()
Attribute Delete.VB_Description = "Marks a DataRow as deleted."
    RequireDataRow Me, "Delete"
    If mRowState = ROW_STATE_DETACHED Then
        RaiseContractError ERROR_INVALID_OPERATION, "Delete", _
            "A detached DataRow cannot be deleted."
    End If
    mRowState = ROW_STATE_DELETED
    mSource.TouchDataVersion
End Sub

Public Sub AddTable(ByVal table As ROneCOne)
Attribute AddTable.VB_Description = "Adds a DataTable to a DataSet."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddTable", _
            "AddTable requires a DataSet."
    End If
    RequireDataTable table, "AddTable"
    If FindNamedObjectIndex(mDataTables, table.TableName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddTable", _
            "A table with the same name already exists."
    End If
    If Not table.InternalSource Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddTable", _
            "The DataTable already belongs to a DataSet."
    End If
    table.ConfigureReadOnlySource Me
    mDataTables.Add table
End Sub

Public Sub AddRelation(ByVal relation As ROneCOne)
Attribute AddRelation.VB_Description = "Adds a DataRelation to a DataSet."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "AddRelation requires a DataSet."
    End If
    RequireDataRelation relation, "AddRelation"
    If FindNamedObjectIndex(mDataRelations, relation.InternalName) >= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "AddRelation", _
            "A relation with the same name already exists."
    End If
    ValidateRelationOwnership relation
    relation.ConfigureReadOnlySource Me
    mDataRelations.Add relation
End Sub

Public Property Get Columns() As ROneCOne
Attribute Columns.VB_Description = "Gets the named DataColumn collection."
    RequireDataTable Me, "Columns"
    Set Columns = CreateNamedDataCollection(mKeys)
End Property

Public Property Get Rows() As ROneCOne
Attribute Rows.VB_Description = "Gets the DataRow collection."
    RequireDataTable Me, "Rows"
    Set Rows = CreateObjectListSnapshot(mItems)
End Property

Public Property Get Tables() As ROneCOne
Attribute Tables.VB_Description = "Gets the named DataTable collection."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "Tables", _
            "Tables requires a DataSet."
    End If
    Set Tables = CreateNamedDataCollection(mDataTables)
End Property

Public Property Get Relations() As ROneCOne
Attribute Relations.VB_Description = "Gets the named DataRelation collection."
    If mRole <> ROLE_DATA_SET Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relations", _
            "Relations requires a DataSet."
    End If
    Set Relations = CreateNamedDataCollection(mDataRelations)
End Property

Public Property Get PrimaryKey() As ROneCOne
Attribute PrimaryKey.VB_Description = "Gets or sets the DataTable primary-key columns."
    RequireDataTable Me, "PrimaryKey"
    Set PrimaryKey = CreateObjectListSnapshot(mPrimaryKey)
End Property

Public Property Set PrimaryKey(ByVal columns As ROneCOne)
    Dim column As ROneCOne

    RequireDataTable Me, "PrimaryKey"
    RequireSequence columns, "PrimaryKey"
    Set mPrimaryKey = New Collection
    For Each column In columns
        RequireDataColumn column, "PrimaryKey"
        If Not column.InternalSource Is Me Then
            RaiseContractError ERROR_INVALID_OPERATION, "PrimaryKey", _
                "Every primary-key column must belong to this table."
        End If
        column.ConfigureUniqueColumn
        mPrimaryKey.Add column
    Next column
    If mPrimaryKey.Count = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "PrimaryKey", _
            "A primary key requires at least one column."
    End If
End Property

Public Function Find(ByVal key As Variant) As ROneCOne
Attribute Find.VB_Description = "Finds a DataRow by a single-column primary key."
    Dim row As ROneCOne

    RequireDataTable Me, "Find"
    If mPrimaryKey.Count <> 1 Then
        RaiseContractError ERROR_INVALID_OPERATION, "Find", _
            "This Find overload requires a single-column primary key."
    End If
    For Each row In mItems
        If DataValuesEqual(row.DataFieldWrapped(mPrimaryKey.Item(1).ColumnName), _
            WrapDataValue(key), "Find") Then
            Set Find = row
            Exit Function
        End If
    Next row
End Function

Public Sub AcceptChanges()
Attribute AcceptChanges.VB_Description = "Commits current DataRow values as original values."
    Dim index As Long
    Dim row As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        AcceptRowChanges
        Exit Sub
    End If
    If mRole = ROLE_DATA_TABLE Then
        For index = mItems.Count To 1 Step -1
            Set row = mItems.Item(index)
            If row.RowState = "Deleted" Then
                mItems.Remove index
            Else
                row.AcceptChanges
            End If
        Next index
        Exit Sub
    End If
    If mRole = ROLE_DATA_SET Then
        For Each row In mDataTables
            row.AcceptChanges
        Next row
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "AcceptChanges", _
        "AcceptChanges requires a DataRow, DataTable, or DataSet."
End Sub

Public Sub RejectChanges()
Attribute RejectChanges.VB_Description = "Rolls back DataRow values to their original state."
    Dim index As Long
    Dim row As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        RejectRowChanges
        Exit Sub
    End If
    If mRole = ROLE_DATA_TABLE Then
        For index = mItems.Count To 1 Step -1
            Set row = mItems.Item(index)
            If row.RowState = "Added" Then
                mItems.Remove index
            Else
                row.RejectChanges
            End If
        Next index
        Exit Sub
    End If
    If mRole = ROLE_DATA_SET Then
        For Each row In mDataTables
            row.RejectChanges
        Next row
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "RejectChanges", _
        "RejectChanges requires a DataRow, DataTable, or DataSet."
End Sub

Public Property Get RowState() As String
Attribute RowState.VB_Description = "Gets the .NET-aligned DataRow state name."
    RequireDataRow Me, "RowState"
    Select Case mRowState
        Case ROW_STATE_DETACHED
            RowState = "Detached"
        Case ROW_STATE_UNCHANGED
            RowState = "Unchanged"
        Case ROW_STATE_ADDED
            RowState = "Added"
        Case ROW_STATE_DELETED
            RowState = "Deleted"
        Case ROW_STATE_MODIFIED
            RowState = "Modified"
    End Select
End Property

Public Function GetChildRows(ByVal relationName As String) As ROneCOne
Attribute GetChildRows.VB_Description = "Gets child rows through a DataRelation."
    RequireDataRow Me, "GetChildRows"
    Set GetChildRows = NavigateRelation(relationName, True)
End Function

Public Function GetParentRow(ByVal relationName As String) As ROneCOne
Attribute GetParentRow.VB_Description = "Gets the parent row through a DataRelation."
    Dim parents As ROneCOne

    RequireDataRow Me, "GetParentRow"
    Set parents = NavigateRelation(relationName, False)
    If parents.Count > 0 Then Set GetParentRow = parents.Item(0)
End Function

Public Function WithFilter(ByVal predicate As Variant) As ROneCOne
Attribute WithFilter.VB_Description = "Applies a typed filter to a DataView."
    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithFilter", _
            "WithFilter requires a DataView."
    End If
    Set mQueryDelegate = NormalizeContextUnary(predicate, "WithFilter", False)
    Set WithFilter = Me
End Function

Public Function WithSort( _
    ByVal keySelector As Variant, _
    Optional ByVal descending As Boolean = False _
) As ROneCOne
Attribute WithSort.VB_Description = "Applies one typed sort key to a DataView."
    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithSort", _
            "WithSort requires a DataView."
    End If
    Set mKeySelector = NormalizeContextUnary(keySelector, "WithSort", False)
    mExpressionOption = Abs(CLng(descending))
    Set WithSort = Me
End Function

Public Function ToTable( _
    Optional ByVal tableName As String = vbNullString _
) As ROneCOne
Attribute ToTable.VB_Description = "Materializes a DataView as an independent DataTable."
    Dim result As ROneCOne
    Dim row As Variant

    If mRole <> ROLE_DATA_VIEW Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToTable", _
            "ToTable requires a DataView."
    End If
    Set result = CloneDataTableSchema(mSource)
    If Len(tableName) > 0 Then result.ConfigureDataName tableName
    For Each row In Me
        CopyDataRowInto row, result
    Next row
    result.AcceptChanges
    Set ToTable = result
End Function

Public Function DbConnection( _
    ByVal connectionString As String, _
    Optional ByVal providerName As String = "OLEDB" _
) As ROneCOne
Attribute DbConnection.VB_Description = "Creates a lazy OLE DB or ODBC connection."
    Dim connection As ROneCOne

    Set connection = New ROneCOne
    connection.ConfigureDbConnection connectionString, providerName
    Set DbConnection = connection
End Function

Public Function DbCommand( _
    ByVal commandText As String, _
    ByVal connection As ROneCOne _
) As ROneCOne
Attribute DbCommand.VB_Description = "Creates a parameterized database command."
    Dim command As ROneCOne

    RequireDbConnection connection, "DbCommand"
    Set command = New ROneCOne
    command.ConfigureDbCommand commandText, connection
    Set DbCommand = command
End Function

Public Function DbParameter( _
    ByVal parameterName As String, _
    ByVal value As Variant, _
    Optional ByVal direction As Long = 1 _
) As ROneCOne
Attribute DbParameter.VB_Description = "Creates an inferred database parameter."
    Dim parameter As ROneCOne

    Set parameter = New ROneCOne
    parameter.ConfigureDbParameter parameterName, WrapDataValue(value), direction
    Set DbParameter = parameter
End Function

Public Function DbDataAdapter(ByVal selectCommand As ROneCOne) As ROneCOne
Attribute DbDataAdapter.VB_Description = "Creates a DataAdapter over a select command."
    Dim adapter As ROneCOne

    RequireDbCommand selectCommand, "DbDataAdapter"
    Set adapter = New ROneCOne
    adapter.ConfigureDbAdapter selectCommand
    Set DbDataAdapter = adapter
End Function

Public Sub AddParameter(ByVal parameter As ROneCOne)
Attribute AddParameter.VB_Description = "Adds a parameter to a database command."
    RequireDbCommand Me, "AddParameter"
    If parameter Is Nothing Or parameter.InternalRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddParameter", _
            "AddParameter requires a DbParameter."
    End If
    mItems.Add parameter
End Sub

Public Function WithParameter( _
    ByVal parameterName As String, _
    ByVal value As Variant, _
    Optional ByVal direction As Long = 1 _
) As ROneCOne
Attribute WithParameter.VB_Description = "Adds an inferred parameter and returns the command."
    AddParameter ROneCOne.DbParameter(parameterName, value, direction)
    Set WithParameter = Me
End Function

Public Function WithTimeout(ByVal seconds As Long) As ROneCOne
Attribute WithTimeout.VB_Description = "Sets the command timeout and returns the command."
    RequireDbCommand Me, "WithTimeout"
    If seconds < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WithTimeout", _
            "Command timeout cannot be negative."
    End If
    mCommandTimeout = seconds
    Set WithTimeout = Me
End Function

Public Function WithCommandType(ByVal commandType As Long) As ROneCOne
Attribute WithCommandType.VB_Description = "Sets the ADO command type and returns the command."
    RequireDbCommand Me, "WithCommandType"
    mCommandType = commandType
    Set WithCommandType = Me
End Function

Public Function FromColumn( _
    ByVal columnName As String, _
    Optional ByVal sourceVersion As String = "Current" _
) As ROneCOne
Attribute FromColumn.VB_Description = "Binds an adapter parameter to a DataColumn value."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "FromColumn", _
            "FromColumn requires a DbParameter."
    End If
    mSourceColumn = Trim$(columnName)
    If Len(mSourceColumn) = 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "FromColumn", _
            "Source column cannot be empty."
    End If
    mSourceVersion = NormalizeDataRowVersion(sourceVersion, "FromColumn")
    Set FromColumn = Me
End Function

Public Function WithDbType( _
    ByVal providerType As Long, _
    Optional ByVal size As Long = 0 _
) As ROneCOne
Attribute WithDbType.VB_Description = "Sets the ADO data type and optional parameter size."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithDbType", _
            "WithDbType requires a DbParameter."
    End If
    If size < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "WithDbType", _
            "Parameter size cannot be negative."
    End If
    mProviderType = providerType
    mProviderSize = size
    Set WithDbType = Me
End Function

Public Function WithPrecision( _
    ByVal precision As Byte, _
    ByVal numericScale As Byte _
) As ROneCOne
Attribute WithPrecision.VB_Description = "Sets numeric parameter precision and scale."
    If mRole <> ROLE_DB_PARAMETER Then
        RaiseContractError ERROR_INVALID_OPERATION, "WithPrecision", _
            "WithPrecision requires a DbParameter."
    End If
    mProviderPrecision = precision
    mProviderScale = numericScale
    Set WithPrecision = Me
End Function

Public Property Get Parameters() As ROneCOne
Attribute Parameters.VB_Description = "Gets a read-only view of command parameters."
    RequireDbCommand Me, "Parameters"
    Set Parameters = CreateReadOnlyCollectionFromItems(mItems)
End Property

Public Property Get State() As String
Attribute State.VB_Description = "Gets the .NET-aligned provider resource state."
    Dim providerState As Long

    If mRole = ROLE_DB_TRANSACTION Then
        If mTransactionActive Then State = "Open" Else State = "Closed"
        Exit Property
    End If
    If mRole <> ROLE_DB_CONNECTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "State", _
            "State requires a DbConnection or transaction."
    End If
    If mProviderObject Is Nothing Then
        State = "Closed"
        Exit Property
    End If
    providerState = CLng(CallByName(mProviderObject, "State", VbGet))
    Select Case providerState
        Case 0
            State = "Closed"
        Case 1
            State = "Open"
        Case 2
            State = "Connecting"
        Case 4
            State = "Executing"
        Case 8
            State = "Fetching"
        Case Else
            State = "Open"
    End Select
End Property

Public Sub Connect()
Attribute Connect.VB_Description = "Opens a database connection explicitly."
    RequireDbConnection Me, "Connect"
    If Not mProviderObject Is Nothing Then
        If CLng(CallByName(mProviderObject, "State", VbGet)) <> 0 Then Exit Sub
    End If
    Set mProviderObject = CreateObject("ADODB.Connection")
    If mConnectionTimeout > 0 Then
        CallByName mProviderObject, "ConnectionTimeout", VbLet, mConnectionTimeout
    End If
    CallByName mProviderObject, "Open", VbMethod, mConnectionString
End Sub

Public Function OpenAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute OpenAsync.VB_Description = "Returns a cooperatively scheduled connection-open Task."
    Set OpenAsync = CreateProviderTask( _
        "Connect", cancellationToken:=cancellationToken)
End Function

Public Sub Disconnect()
Attribute Disconnect.VB_Description = "Closes a provider resource deterministically."
    If mRole = ROLE_DB_CONNECTION Or mRole = ROLE_DB_READER Then
        If Not mProviderObject Is Nothing Then
            If CLng(CallByName(mProviderObject, "State", VbGet)) <> 0 Then
                CallByName mProviderObject, "Close", VbMethod
            End If
        End If
        Set mProviderObject = Nothing
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Disconnect", _
        "Disconnect requires a connection or data reader."
End Sub

Public Function ExecuteReader() As ROneCOne
Attribute ExecuteReader.VB_Description = "Executes a command and returns a forward-only reader."
    Dim commandObject As Object
    Dim reader As ROneCOne
    Dim recordset As Object

    RequireDbCommand Me, "ExecuteReader"
    Set commandObject = BuildAdoCommand
    Set recordset = commandObject.Execute
    Set reader = New ROneCOne
    reader.ConfigureDbReader recordset
    Set ExecuteReader = reader
End Function

Public Function ExecuteReaderAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteReaderAsync.VB_Description = "Returns a Task that executes a data reader."
    Set ExecuteReaderAsync = CreateProviderTask( _
        "ExecuteReader", cancellationToken:=cancellationToken)
End Function

Public Function ExecuteNonQuery() As Long
Attribute ExecuteNonQuery.VB_Description = "Executes a command and returns affected rows."
    Dim commandObject As Object
    Dim recordsAffected As Variant
    Dim recordset As Object

    RequireDbCommand Me, "ExecuteNonQuery"
    Set commandObject = BuildAdoCommand
    Set recordset = commandObject.Execute(recordsAffected)
    ExecuteNonQuery = CLng(recordsAffected)
    If Not recordset Is Nothing Then
        If CLng(CallByName(recordset, "State", VbGet)) <> 0 Then
            CallByName recordset, "Close", VbMethod
        End If
    End If
End Function

Public Function ExecuteNonQueryAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteNonQueryAsync.VB_Description = "Returns a Task that executes a non-query."
    Set ExecuteNonQueryAsync = CreateProviderTask( _
        "ExecuteNonQuery", cancellationToken:=cancellationToken)
End Function

Public Function ExecuteScalar() As Variant
Attribute ExecuteScalar.VB_Description = "Executes a command and returns its first field."
    Dim reader As ROneCOne

    Set reader = ExecuteReader
    If reader.Read Then ExecuteScalar = reader.Item(0)
    reader.Disconnect
End Function

Public Function ExecuteScalarAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ExecuteScalarAsync.VB_Description = "Returns a Task that executes a scalar command."
    Set ExecuteScalarAsync = CreateProviderTask( _
        "ExecuteScalar", cancellationToken:=cancellationToken)
End Function

Public Function BeginTransaction() As ROneCOne
Attribute BeginTransaction.VB_Description = "Begins a provider transaction."
    Dim transaction As ROneCOne

    RequireDbConnection Me, "BeginTransaction"
    Connect
    CallByName mProviderObject, "BeginTrans", VbMethod
    Set transaction = New ROneCOne
    transaction.ConfigureDbTransaction Me
    Set BeginTransaction = transaction
End Function

Public Sub Commit()
Attribute Commit.VB_Description = "Commits an active provider transaction."
    If mRole <> ROLE_DB_TRANSACTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "Commit", _
            "Commit requires a transaction."
    End If
    If Not mTransactionActive Then
        RaiseContractError ERROR_INVALID_OPERATION, "Commit", _
            "The transaction has already completed."
    End If
    CallByName mSource.InternalProviderObject, "CommitTrans", VbMethod
    mTransactionActive = False
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Public Sub Rollback()
Attribute Rollback.VB_Description = "Rolls back an active provider transaction."
    If mRole <> ROLE_DB_TRANSACTION Then
        RaiseContractError ERROR_INVALID_OPERATION, "Rollback", _
            "Rollback requires a transaction."
    End If
    If Not mTransactionActive Then
        RaiseContractError ERROR_INVALID_OPERATION, "Rollback", _
            "The transaction has already completed."
    End If
    CallByName mSource.InternalProviderObject, "RollbackTrans", VbMethod
    mTransactionActive = False
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Public Function Read() As Boolean
Attribute Read.VB_Description = "Advances a database reader to the next record."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Read", _
            "Read requires a data reader."
    End If
    If CBool(CallByName(mProviderObject, "EOF", VbGet)) Then Exit Function
    If mQueryCount > 0 Then CallByName mProviderObject, "MoveNext", VbMethod
    If CBool(CallByName(mProviderObject, "EOF", VbGet)) Then Exit Function
    mQueryCount = mQueryCount + 1
    Read = True
End Function

Public Property Get FieldCount() As Long
Attribute FieldCount.VB_Description = "Gets the number of fields in a data reader."
    Dim fields As Object

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "FieldCount", _
            "FieldCount requires a data reader."
    End If
    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    FieldCount = CLng(CallByName(fields, "Count", VbGet))
End Property

Public Function GetName(ByVal ordinal As Long) As String
Attribute GetName.VB_Description = "Gets a data-reader field name by ordinal."
    Dim field As Object
    Dim fields As Object

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetName", _
            "GetName requires a data reader."
    End If
    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    Set field = CallByName(fields, "Item", VbGet, ordinal)
    GetName = CStr(CallByName(field, "Name", VbGet))
End Function

Public Function GetOrdinal(ByVal fieldName As String) As Long
Attribute GetOrdinal.VB_Description = "Gets a data-reader field ordinal by name."
    Dim ordinal As Long

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetOrdinal", _
            "GetOrdinal requires a data reader."
    End If
    For ordinal = 0 To FieldCount - 1
        If StrComp(GetName(ordinal), fieldName, vbTextCompare) = 0 Then
            GetOrdinal = ordinal
            Exit Function
        End If
    Next ordinal
    RaiseContractError ERROR_INVALID_ARGUMENT, "GetOrdinal", _
        "The requested field was not found."
End Function

Public Function GetValues() As ROneCOne
Attribute GetValues.VB_Description = "Copies the current reader row into a typed sequence."
    Dim ordinal As Long
    Dim result As ROneCOne

    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "GetValues", _
            "GetValues requires a data reader."
    End If
    Set result = ROneCOne.ListOf(vbVariant)
    For ordinal = 0 To FieldCount - 1
        result.Add ReadProviderField(ordinal)
    Next ordinal
    Set GetValues = result
End Function

Public Function IsDBNull(ByVal ordinal As Variant) As Boolean
Attribute IsDBNull.VB_Description = "Tests whether the current provider field contains Null."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "IsDBNull", _
            "IsDBNull requires a data reader."
    End If
    IsDBNull = IsNull(ReadProviderField(ordinal))
End Function

Public Property Get HasRows() As Boolean
Attribute HasRows.VB_Description = "Reports whether a data reader contains at least one row."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "HasRows", _
            "HasRows requires a data reader."
    End If
    If mProviderObject Is Nothing Then Exit Property
    HasRows = Not CBool(CallByName(mProviderObject, "EOF", VbGet))
End Property

Public Property Get IsClosed() As Boolean
Attribute IsClosed.VB_Description = "Reports whether a provider reader is closed."
    If mRole <> ROLE_DB_READER Then
        RaiseContractError ERROR_INVALID_OPERATION, "IsClosed", _
            "IsClosed requires a data reader."
    End If
    If mProviderObject Is Nothing Then
        IsClosed = True
    Else
        IsClosed = (CLng(CallByName(mProviderObject, "State", VbGet)) = 0)
    End If
End Property

Public Function ReadAsync( _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute ReadAsync.VB_Description = "Returns a Task that advances a provider reader."
    Set ReadAsync = CreateProviderTask( _
        "Read", cancellationToken:=cancellationToken)
End Function

Public Property Get SelectCommand() As ROneCOne
Attribute SelectCommand.VB_Description = "Gets the adapter select command."
    RequireDbAdapter Me, "SelectCommand"
    Set SelectCommand = mSource
End Property

Public Property Set SelectCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "SelectCommand"
    RequireDbCommand value, "SelectCommand"
    Set mSource = value
End Property

Public Property Get InsertCommand() As ROneCOne
Attribute InsertCommand.VB_Description = "Gets the adapter insert command."
    RequireDbAdapter Me, "InsertCommand"
    Set InsertCommand = mLeft
End Property

Public Property Set InsertCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "InsertCommand"
    RequireDbCommand value, "InsertCommand"
    Set mLeft = value
End Property

Public Property Get UpdateCommand() As ROneCOne
Attribute UpdateCommand.VB_Description = "Gets the adapter update command."
    RequireDbAdapter Me, "UpdateCommand"
    Set UpdateCommand = mRight
End Property

Public Property Set UpdateCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "UpdateCommand"
    RequireDbCommand value, "UpdateCommand"
    Set mRight = value
End Property

Public Property Get DeleteCommand() As ROneCOne
Attribute DeleteCommand.VB_Description = "Gets the adapter delete command."
    RequireDbAdapter Me, "DeleteCommand"
    Set DeleteCommand = mNext
End Property

Public Property Set DeleteCommand(ByVal value As ROneCOne)
    RequireDbAdapter Me, "DeleteCommand"
    RequireDbCommand value, "DeleteCommand"
    Set mNext = value
End Property

Public Function Fill(ByVal dataTarget As ROneCOne) As Long
Attribute Fill.VB_Description = "Fills a DataTable or DataSet from a DataAdapter."
    Dim reader As ROneCOne
    Dim table As ROneCOne

    If mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Fill", _
            "Fill requires a DataAdapter."
    End If
    If dataTarget.InternalRole = ROLE_DATA_SET Then
        Set table = ROneCOne.DataTable("Table")
        dataTarget.AddTable table
    Else
        RequireDataTable dataTarget, "Fill"
        Set table = dataTarget
    End If
    Set reader = mSource.ExecuteReader
    Fill = FillTableFromReader(table, reader)
    reader.Disconnect
    table.AcceptChanges
End Function

Public Function FillAsync( _
    ByVal dataTarget As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute FillAsync.VB_Description = "Returns a Task that fills a DataTable or DataSet."
    Set FillAsync = CreateProviderTask( _
        "Fill", dataTarget, cancellationToken)
End Function

Public Function Update(ByVal dataTarget As ROneCOne) As Long
Attribute Update.VB_Description = "Persists DataSet changes through configured commands."
    Dim table As ROneCOne

    If mRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, "Update", _
            "Update requires a DataAdapter."
    End If
    If dataTarget.InternalRole = ROLE_DATA_SET Then
        For Each table In dataTarget.InternalDataTables
            Update = Update + UpdateDataTable(table)
        Next table
        Exit Function
    End If
    RequireDataTable dataTarget, "Update"
    Update = UpdateDataTable(dataTarget)
End Function

Public Function UpdateAsync( _
    ByVal dataTarget As ROneCOne, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
Attribute UpdateAsync.VB_Description = "Returns a Task that persists DataTable changes."
    Set UpdateAsync = CreateProviderTask( _
        "Update", dataTarget, cancellationToken)
End Function

Public Sub Dispose()
Attribute Dispose.VB_Description = "Releases a provider resource deterministically."
    If mRole = ROLE_DB_CONNECTION Or mRole = ROLE_DB_READER Then
        Disconnect
        Exit Sub
    End If
    If mRole = ROLE_DB_TRANSACTION Then
        If mTransactionActive Then Rollback
        Exit Sub
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "Dispose", _
        "Dispose requires a provider connection, reader, or transaction."
End Sub

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

Public Function OutOf(ByVal typeToken As Variant) As ROneCOne
Attribute OutOf.VB_Description = "Creates a typed managed out-value container."
    Dim objectType As String
    Dim referenceValue As ROneCOne
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "OutOf"
    Set referenceValue = New ROneCOne
    referenceValue.ConfigureManagedReference valueType, objectType
    Set OutOf = referenceValue
End Function

Public Property Get Current() As Variant
Attribute Current.VB_Description = "Gets the value held by a managed out-value container."
    RequireReference Me, "Current"
    If mQueryValue Is Nothing Then
        AssignDefaultReferenceResult Current
    Else
        AssignPublicResult Current, mQueryValue
    End If
End Property

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
    Dim itemIndex As Long
    Dim result As ROneCOne

    If mRole = ROLE_TASK_FACTORY Then
        If argumentValues.Count < 1 Or argumentValues.Count > 2 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Task.Run", _
                "Task.Run requires work and an optional cancellation token."
        End If
        Set argumentValue = argumentValues.Item(1)
        If argumentValues.Count = 1 Then
            Set result = TaskRun( _
                RequireROneCOne(Unwrap(argumentValue), "Task.Run"))
        Else
            Set result = TaskRun( _
                RequireROneCOne(Unwrap(argumentValue), "Task.Run"), _
                RequireROneCOne(Unwrap(argumentValues.Item(2)), "Task.Run"))
        End If
        Set destination = result
        Exit Sub
    End If

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

    If IsGenericCollectionRole(mRole) And _
        (IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED) Then
        If argumentValues.Count <> 1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, "Run", _
                "A keyed indexer requires one argument."
        End If
        Set argumentValue = argumentValues.Item(1)
        itemIndex = FindKeyIndex(WrapKey(Unwrap(argumentValue), "Run"), "Run")
        If itemIndex < 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Run", _
                "The requested key was not found."
        End If
        AssignPublicResult destination, mItems.Item(itemIndex + 1)
        Exit Sub
    End If

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
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
        If IsGenericCollectionRole(mRole) Then
            Set result = MaterializeWrapped.Item( _
                CLng(argumentValue.InternalScalarValue) + 1)
        Else
            Set result = WrappedItem(CLng(argumentValue.InternalScalarValue))
        End If
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

Public Function Add( _
    ByVal operand As Variant, _
    Optional ByVal value As Variant _
) As ROneCOne
Attribute Add.VB_Description = "Adds a list item or creates an addition expression."
    Dim clone As ROneCOne
    Dim wrapped As ROneCOne

    If mRole = ROLE_LIST Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "List<T>.Add accepts one item."
        End If
        Set wrapped = WrapElement(operand, "Add")
        mItems.Add wrapped
        AddUnwrappedValue mEnumerationValues, wrapped
        mVersion = mVersion + 1
        Set Add = Me
    ElseIf mRole = ROLE_IMMUTABLE Then
        Set clone = CloneGenericCollection(Me)
        clone.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
        clone.AddMutable operand, value, "Add"
        clone.ConfigureCollectionRole ROLE_IMMUTABLE
        Set Add = clone
    ElseIf IsGenericCollectionRole(mRole) Then
        AddMutable operand, value, "Add"
        Set Add = Me
    Else
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Add", _
                "An addition expression accepts one operand."
        End If
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
Attribute Concat.VB_Description = "Concatenates strings or compatible sequences."
    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
        Set Concat = ConcatSequence(operand)
    Else
        Set Concat = CreateBinaryExpression(OP_CONCAT, operand)
    End If
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

Public Property Get Item(ByVal index As Variant) As Variant
Attribute Item.VB_Description = "Gets or sets a zero-based typed list item."
    Dim itemIndex As Long
    Dim wrapped As ROneCOne

    If mRole = ROLE_DB_READER Then
        Set wrapped = WrapDataValue(ReadProviderField(index))
    ElseIf mRole = ROLE_DATA_ROW Then
        Set wrapped = DataRowItem(index)
    ElseIf IsDictionaryCollection Then
        itemIndex = FindKeyIndex(WrapKey(index, "Item"), "Item")
        If itemIndex < 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Item", _
                "The requested key was not found."
        End If
        Set wrapped = mItems.Item(itemIndex + 1)
    ElseIf mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP Then
        itemIndex = FindKeyIndex(WrapKey(index, "Item"), "Item")
        If itemIndex < 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Item", _
                "The requested key was not found."
        End If
        Set wrapped = mItems.Item(itemIndex + 1)
    ElseIf mCollectionKind = COLLECTION_READ_ONLY Then
        Set wrapped = mSource.MaterializeWrapped.Item(CLng(index) + 1)
    ElseIf IsGenericCollectionRole(mRole) Then
        ValidateIndex CLng(index), mItems.Count, "Item"
        Set wrapped = mItems.Item(CLng(index) + 1)
    Else
        Set wrapped = WrappedItem(CLng(index))
    End If
    AssignPublicResult Item, wrapped
End Property

Public Property Let Item(ByVal index As Variant, ByVal value As Variant)
    SetItemValue index, value
End Property

Public Property Set Item(ByVal index As Variant, ByVal value As Object)
    SetItemValue index, value
End Property

Public Property Get Count(Optional ByVal predicate As Variant) As Long
Attribute Count.VB_Description = "Counts every element or only predicate matches."
    Dim predicateValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If IsGenericCollectionRole(mRole) And IsMissing(predicate) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            Count = mSource.Count
        Else
            Count = mItems.Count
        End If
        Exit Property
    End If
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

Public Property Get LongCount() As LongLong
Attribute LongCount.VB_Description = "Counts sequence elements as a 64-bit integer."
    Dim values As Collection

    Set values = MaterializeWrapped
    LongCount = CLngLng(values.Count)
End Property

Public Function TryGetNonEnumeratedCount( _
    ByVal countReference As ROneCOne _
) As Boolean
Attribute TryGetNonEnumeratedCount.VB_Description = _
    "Gets a known count without executing a deferred query."
    RequireReference countReference, "TryGetNonEnumeratedCount"
    If mRole = ROLE_QUERY Then Exit Function
    If mRole <> ROLE_LIST And Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, _
            "TryGetNonEnumeratedCount", "This operation requires a sequence."
    End If
    AssignReferenceWrapped countReference, CreateScalarResult(CLng(Count)), _
        "TryGetNonEnumeratedCount"
    TryGetNonEnumeratedCount = True
End Function

Public Property Get GenericTypeName() As String
Attribute GenericTypeName.VB_Description = "Gets the C#-style generic sequence type name."
    RequireSequence Me, "GenericTypeName"
    If IsGenericCollectionRole(mRole) Then
        GenericTypeName = BuildCollectionTypeName
        Exit Property
    End If
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
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "Clear"
        Set mItems = New Collection
        Set mKeys = New Collection
        Set mPriorities = New Collection
        mVersion = mVersion + 1
        NotifyCollectionChanged "Reset"
        Exit Sub
    End If
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

    If mRole = ROLE_LIST Or mRole = ROLE_QUERY Or _
        IsGenericCollectionRole(mRole) Then
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
    Dim immutableResult As ROneCOne
    Dim result As ROneCOne

    If mRole = ROLE_LIST Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "List removal accepts exactly one value."
        End If
        Remove = RemoveListItem(value)
        Exit Function
    End If

    If mRole = ROLE_IMMUTABLE Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Immutable collection removal accepts one value or key."
        End If
        Set immutableResult = CloneGenericCollection(Me)
        immutableResult.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
        RemoveCollectionValue immutableResult, value, "Remove"
        immutableResult.ConfigureCollectionRole ROLE_IMMUTABLE
        Set Remove = immutableResult
        Exit Function
    End If
    If IsGenericCollectionRole(mRole) Then
        If Not IsMissing(removal) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Remove", _
                "Collection removal accepts one value or key."
        End If
        Remove = RemoveCollectionValue(Me, value, "Remove")
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
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "RemoveAt"
        ValidateIndex index, mItems.Count, "RemoveAt"
        RemoveCollectionAt index
        Exit Sub
    End If
    RequireList Me, "RemoveAt"
    ValidateIndex index, mItems.Count, "RemoveAt"
    mItems.Remove index + 1
    mEnumerationValues.Remove index + 1
    mVersion = mVersion + 1
End Sub

Public Function ContainsKey(ByVal key As Variant) As Boolean
Attribute ContainsKey.VB_Description = "Tests whether a keyed collection contains a key."
    RequireKeyedCollection "ContainsKey"
    ContainsKey = (FindKeyIndex(WrapKey(key, "ContainsKey"), "ContainsKey") >= 0)
End Function

Public Function TryAdd( _
    ByVal keyOrValue As Variant, _
    Optional ByVal value As Variant _
) As Boolean
Attribute TryAdd.VB_Description = "Adds only when a key or set value is absent."
    Dim keyWrapped As ROneCOne

    RequireMutableCollection "TryAdd"
    If IsDictionaryCollection Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "TryAdd", _
                "A dictionary value is required."
        End If
        Set keyWrapped = WrapKey(keyOrValue, "TryAdd")
        If FindKeyIndex(keyWrapped, "TryAdd") >= 0 Then Exit Function
        AddDictionaryPair keyWrapped, WrapElement(value, "TryAdd"), "TryAdd"
        TryAdd = True
        Exit Function
    End If
    If IsSetCollection Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "TryAdd", _
                "A set accepts one value."
        End If
        Set keyWrapped = WrapElement(keyOrValue, "TryAdd")
        If FindElementIndex(keyWrapped, "TryAdd") >= 0 Then Exit Function
        AddSetValue keyWrapped, "TryAdd"
        TryAdd = True
        Exit Function
    End If
    RaiseContractError ERROR_INVALID_OPERATION, "TryAdd", _
        "TryAdd requires a dictionary or set."
End Function

Public Function TryGetValue( _
    ByVal key As Variant, _
    ByVal valueReference As ROneCOne _
) As Boolean
Attribute TryGetValue.VB_Description = "Copies a keyed value into a typed reference."
    Dim itemIndex As Long

    RequireKeyedCollection "TryGetValue"
    RequireReference valueReference, "TryGetValue"
    itemIndex = FindKeyIndex(WrapKey(key, "TryGetValue"), "TryGetValue")
    If itemIndex < 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(itemIndex + 1), _
        "TryGetValue"
    TryGetValue = True
End Function

Public Property Get Keys() As ROneCOne
Attribute Keys.VB_Description = "Gets a typed snapshot of the collection keys."
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireKeyedCollection "Keys"
    Set result = New ROneCOne
    result.ConfigureList mKeyType, mKeyObjectType
    For Each wrapped In mKeys
        result.InternalAddWrapped wrapped
    Next wrapped
    Set Keys = result
End Property

Public Property Get Values() As ROneCOne
Attribute Values.VB_Description = "Gets a typed snapshot of the collection values."
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireKeyedCollection "Values"
    Set result = New ROneCOne
    result.ConfigureList mElementType, mElementObjectType
    For Each wrapped In mItems
        result.InternalAddWrapped wrapped
    Next wrapped
    Set Values = result
End Property

Public Property Get Key() As Variant
Attribute Key.VB_Description = "Gets the key of a KeyValuePair value."
    If mRole <> ROLE_ENTRY And mCollectionKind <> COLLECTION_GROUPING Then
        RaiseContractError ERROR_INVALID_OPERATION, "Key", _
            "Key requires a KeyValuePair value."
    End If
    AssignPublicResult Key, mLeft
End Property

Public Function Enqueue( _
    ByVal element As Variant, _
    Optional ByVal priority As Variant _
) As ROneCOne
Attribute Enqueue.VB_Description = "Adds an element to a queue or priority queue."
    Dim target As ROneCOne

    If mRole = ROLE_IMMUTABLE Then
        Set target = CloneGenericCollection(Me)
        target.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Else
        RequireMutableCollection "Enqueue"
        Set target = Me
    End If
    If IsPriorityQueueKind(target.InternalCollectionKind) Then
        If IsMissing(priority) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Enqueue", _
                "A priority value is required."
        End If
        target.AddPriorityValue target.WrapElement(element, "Enqueue"), _
            target.WrapKey(priority, "Enqueue")
    ElseIf IsQueueKind(target.InternalCollectionKind) Then
        If Not IsMissing(priority) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "Enqueue", _
                "This queue does not accept a priority."
        End If
        target.InternalQueueValue target.WrapElement(element, "Enqueue")
    Else
        RaiseContractError ERROR_INVALID_OPERATION, "Enqueue", _
            "Enqueue requires a queue."
    End If
    If mRole = ROLE_IMMUTABLE Then
        target.ConfigureCollectionRole ROLE_IMMUTABLE
    End If
    Set Enqueue = target
End Function

Public Function Dequeue() As Variant
Attribute Dequeue.VB_Description = "Removes and returns the next queue element."
    Dim wrapped As ROneCOne

    RequireMutableCollection "Dequeue"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Dequeue", _
            "Dequeue requires a queue."
    End If
    If mItems.Count = 0 Then RaiseEmptySequence "Dequeue"
    Set wrapped = mItems.Item(1)
    RemoveCollectionAt 0
    AssignPublicResult Dequeue, wrapped
End Function

Public Function TryDequeue(ByVal valueReference As ROneCOne) As Boolean
Attribute TryDequeue.VB_Description = "Attempts to remove the next queue element."
    RequireMutableCollection "TryDequeue"
    RequireReference valueReference, "TryDequeue"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryDequeue", _
            "TryDequeue requires a queue."
    End If
    If mItems.Count = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(1), "TryDequeue"
    RemoveCollectionAt 0
    TryDequeue = True
End Function

Public Function Push(ByVal element As Variant) As ROneCOne
Attribute Push.VB_Description = "Pushes an element onto a stack."
    Dim target As ROneCOne

    If mRole = ROLE_IMMUTABLE Then
        Set target = CloneGenericCollection(Me)
        target.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Else
        RequireMutableCollection "Push"
        Set target = Me
    End If
    If Not IsStackKind(target.InternalCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Push", _
            "Push requires a stack."
    End If
    target.InternalPushValue target.WrapElement(element, "Push")
    If mRole = ROLE_IMMUTABLE Then
        target.ConfigureCollectionRole ROLE_IMMUTABLE
    End If
    Set Push = target
End Function

Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes and returns the top stack element."
    Dim wrapped As ROneCOne

    RequireMutableCollection "Pop"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Pop", _
            "Pop requires a stack."
    End If
    If mItems.Count = 0 Then RaiseEmptySequence "Pop"
    Set wrapped = mItems.Item(1)
    RemoveCollectionAt 0
    AssignPublicResult Pop, wrapped
End Function

Public Function TryPop(ByVal valueReference As ROneCOne) As Boolean
Attribute TryPop.VB_Description = "Attempts to remove the top stack element."
    RequireMutableCollection "TryPop"
    RequireReference valueReference, "TryPop"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryPop", _
            "TryPop requires a stack."
    End If
    If mItems.Count = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(1), "TryPop"
    RemoveCollectionAt 0
    TryPop = True
End Function

Public Function Peek() As Variant
Attribute Peek.VB_Description = "Returns the next queue or stack element without removal."
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) And _
        Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "Peek", _
            "Peek requires a queue or stack."
    End If
    If mItems.Count = 0 Then RaiseEmptySequence "Peek"
    AssignPublicResult Peek, mItems.Item(1)
End Function

Public Function TryPeek(ByVal valueReference As ROneCOne) As Boolean
Attribute TryPeek.VB_Description = "Attempts to read the next queue or stack element."
    RequireReference valueReference, "TryPeek"
    If Not IsQueueKind(mCollectionKind) And _
        Not IsPriorityQueueKind(mCollectionKind) And _
        Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "TryPeek", _
            "TryPeek requires a queue or stack."
    End If
    If mItems.Count = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(1), "TryPeek"
    TryPeek = True
End Function

Public Function AddFirst(ByVal value As Variant) As ROneCOne
Attribute AddFirst.VB_Description = "Adds an item at the beginning of a linked list."
    RequireLinkedList "AddFirst"
    RequireMutableCollection "AddFirst"
    If mItems.Count = 0 Then
        mItems.Add WrapElement(value, "AddFirst")
    Else
        mItems.Add WrapElement(value, "AddFirst"), Before:=1
    End If
    mVersion = mVersion + 1
    Set AddFirst = FirstNode
End Function

Public Function AddLast(ByVal value As Variant) As ROneCOne
Attribute AddLast.VB_Description = "Adds an item at the end of a linked list."
    RequireLinkedList "AddLast"
    RequireMutableCollection "AddLast"
    mItems.Add WrapElement(value, "AddLast")
    mVersion = mVersion + 1
    Set AddLast = LastNode
End Function

Public Sub RemoveFirst()
Attribute RemoveFirst.VB_Description = "Removes the first linked-list item."
    RequireLinkedList "RemoveFirst"
    RequireMutableCollection "RemoveFirst"
    If mItems.Count = 0 Then RaiseEmptySequence "RemoveFirst"
    RemoveCollectionAt 0
End Sub

Public Sub RemoveLast()
Attribute RemoveLast.VB_Description = "Removes the last linked-list item."
    RequireLinkedList "RemoveLast"
    RequireMutableCollection "RemoveLast"
    If mItems.Count = 0 Then RaiseEmptySequence "RemoveLast"
    RemoveCollectionAt mItems.Count - 1
End Sub

Public Property Get FirstNode() As ROneCOne
Attribute FirstNode.VB_Description = "Gets the first linked-list node."
    RequireLinkedList "FirstNode"
    If mItems.Count = 0 Then Exit Property
    Set FirstNode = CreateCollectionNode(0)
End Property

Public Property Get LastNode() As ROneCOne
Attribute LastNode.VB_Description = "Gets the last linked-list node."
    RequireLinkedList "LastNode"
    If mItems.Count = 0 Then Exit Property
    Set LastNode = CreateCollectionNode(mItems.Count - 1)
End Property

Public Property Get MinValue() As Variant
Attribute MinValue.VB_Description = "Gets the minimum element of a sorted set."
    RequireSortedSet "MinValue"
    If mItems.Count = 0 Then RaiseEmptySequence "MinValue"
    AssignPublicResult MinValue, mItems.Item(1)
End Property

Public Property Get MaxValue() As Variant
Attribute MaxValue.VB_Description = "Gets the maximum element of a sorted set."
    RequireSortedSet "MaxValue"
    If mItems.Count = 0 Then RaiseEmptySequence "MaxValue"
    AssignPublicResult MaxValue, mItems.Item(mItems.Count)
End Property

Public Function GetViewBetween( _
    ByVal lowerValue As Variant, _
    ByVal upperValue As Variant _
) As ROneCOne
Attribute GetViewBetween.VB_Description = "Gets a sorted-set range view."
    Dim lowerWrapped As ROneCOne
    Dim result As ROneCOne
    Dim upperWrapped As ROneCOne
    Dim wrapped As ROneCOne

    RequireSortedSet "GetViewBetween"
    Set lowerWrapped = WrapElement(lowerValue, "GetViewBetween")
    Set upperWrapped = WrapElement(upperValue, "GetViewBetween")
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_HASH_SET, COLLECTION_SORTED_SET, _
        mElementType, mElementObjectType, vbEmpty, vbNullString, mQueryComparer
    For Each wrapped In mItems
        If CompareWithComparer(wrapped, lowerWrapped, mQueryComparer, _
            "GetViewBetween") >= 0 And _
            CompareWithComparer(wrapped, upperWrapped, mQueryComparer, _
            "GetViewBetween") <= 0 Then
            result.AddSetValue wrapped, "GetViewBetween"
        End If
    Next wrapped
    Set GetViewBetween = result
End Function

Public Property Get CollectionChanged() As ROneCOne
Attribute CollectionChanged.VB_Description = "Gets the observable collection change event."
    If mCollectionKind <> COLLECTION_OBSERVABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "CollectionChanged", _
            "CollectionChanged requires an ObservableCollection<T>."
    End If
    Set CollectionChanged = mCollectionChanged
End Property

Public Property Get IsReadOnly() As Boolean
Attribute IsReadOnly.VB_Description = "Reports whether mutating operations are rejected."
    IsReadOnly = (mRole = ROLE_IMMUTABLE Or _
        mCollectionKind = COLLECTION_READ_ONLY Or _
        mCollectionKind = COLLECTION_LOOKUP Or _
        mCollectionKind = COLLECTION_GROUPING)
End Property

Public Function ToBuilder() As ROneCOne
Attribute ToBuilder.VB_Description = "Creates a mutable builder for an immutable collection."
    Dim result As ROneCOne

    If mRole <> ROLE_IMMUTABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToBuilder", _
            "ToBuilder requires an immutable collection."
    End If
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    result.ConfigureCollectionBuilder True
    Set ToBuilder = result
End Function

Public Function ToImmutable() As ROneCOne
Attribute ToImmutable.VB_Description = "Creates an immutable collection snapshot."
    Dim result As ROneCOne

    If Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, "ToImmutable", _
            "ToImmutable requires a generic collection."
    End If
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    result.ConfigureCollectionBuilder False
    Set ToImmutable = result
End Function

Public Sub UnionWith(ByVal other As Variant)
Attribute UnionWith.VB_Description = "Modifies a set to contain values from both sets."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "UnionWith"
    RequireMutableCollection "UnionWith"
    Set incoming = BuildIncomingItems(other, "UnionWith")
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "UnionWith") < 0 Then
            AddSetValue wrapped, "UnionWith"
        End If
    Next wrapped
End Sub

Public Sub IntersectWith(ByVal other As Variant)
Attribute IntersectWith.VB_Description = "Keeps only values found in both sets."
    Dim incoming As Collection
    Dim index As Long

    RequireSetCollection "IntersectWith"
    RequireMutableCollection "IntersectWith"
    Set incoming = BuildIncomingItems(other, "IntersectWith")
    For index = mItems.Count To 1 Step -1
        If FindWrappedInCollection(incoming, mItems.Item(index), _
            "IntersectWith") < 0 Then RemoveCollectionAt index - 1
    Next index
End Sub

Public Function Overlaps(ByVal other As Variant) As Boolean
Attribute Overlaps.VB_Description = "Tests whether two sets share at least one value."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "Overlaps"
    Set incoming = BuildIncomingItems(other, "Overlaps")
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "Overlaps") >= 0 Then
            Overlaps = True
            Exit Function
        End If
    Next wrapped
End Function

Public Function SetEquals(ByVal other As Variant) As Boolean
Attribute SetEquals.VB_Description = "Tests whether two sets contain the same values."
    Dim incoming As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "SetEquals"
    Set incoming = BuildIncomingItems(other, "SetEquals")
    If incoming.Count <> mItems.Count Then Exit Function
    For Each wrapped In incoming
        If FindElementIndex(wrapped, "SetEquals") < 0 Then Exit Function
    Next wrapped
    SetEquals = True
End Function

Public Function GetOrAdd( _
    ByVal key As Variant, _
    ByVal value As Variant _
) As Variant
Attribute GetOrAdd.VB_Description = "Gets an existing value or atomically adds one."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne

    RequireMutableCollection "GetOrAdd"
    RequireKeyedCollection "GetOrAdd"
    Set keyWrapped = WrapKey(key, "GetOrAdd")
    itemIndex = FindKeyIndex(keyWrapped, "GetOrAdd")
    If itemIndex < 0 Then
        AddDictionaryPair keyWrapped, WrapElement(value, "GetOrAdd"), "GetOrAdd"
        itemIndex = FindKeyIndex(keyWrapped, "GetOrAdd")
    End If
    AssignPublicResult GetOrAdd, mItems.Item(itemIndex + 1)
End Function

Public Function TryRemove( _
    ByVal key As Variant, _
    ByVal valueReference As ROneCOne _
) As Boolean
Attribute TryRemove.VB_Description = "Atomically removes a keyed value when present."
    Dim itemIndex As Long

    RequireMutableCollection "TryRemove"
    RequireKeyedCollection "TryRemove"
    RequireReference valueReference, "TryRemove"
    itemIndex = FindKeyIndex(WrapKey(key, "TryRemove"), "TryRemove")
    If itemIndex < 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(itemIndex + 1), "TryRemove"
    RemoveCollectionAt itemIndex
    TryRemove = True
End Function

Public Function TryUpdate( _
    ByVal key As Variant, _
    ByVal newValue As Variant, _
    ByVal comparisonValue As Variant _
) As Boolean
Attribute TryUpdate.VB_Description = "Updates a concurrent value after comparison succeeds."
    Dim itemIndex As Long
    Dim expected As ROneCOne

    RequireConcurrentDictionary "TryUpdate"
    itemIndex = FindKeyIndex(WrapKey(key, "TryUpdate"), "TryUpdate")
    If itemIndex < 0 Then Exit Function
    Set expected = WrapElement(comparisonValue, "TryUpdate")
    If Not ValuesEqual(mItems.Item(itemIndex + 1), expected) Then Exit Function
    ReplaceCollectionItem itemIndex, WrapElement(newValue, "TryUpdate")
    TryUpdate = True
End Function

Public Function AddOrUpdate( _
    ByVal key As Variant, _
    ByVal addValue As Variant, _
    ByVal updateValue As Variant _
) As Variant
Attribute AddOrUpdate.VB_Description = "Adds a concurrent value or replaces the existing value."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim selected As ROneCOne

    RequireConcurrentDictionary "AddOrUpdate"
    Set keyWrapped = WrapKey(key, "AddOrUpdate")
    itemIndex = FindKeyIndex(keyWrapped, "AddOrUpdate")
    If itemIndex < 0 Then
        Set selected = WrapElement(addValue, "AddOrUpdate")
        AddDictionaryPair keyWrapped, selected, "AddOrUpdate"
    Else
        Set selected = WrapElement(updateValue, "AddOrUpdate")
        ReplaceCollectionItem itemIndex, selected
    End If
    AssignPublicResult AddOrUpdate, selected
End Function

Public Sub ExceptWith(ByVal other As Variant)
Attribute ExceptWith.VB_Description = "Removes every value found in another sequence."
    Dim incoming As Collection
    Dim index As Long

    RequireSetCollection "ExceptWith"
    RequireMutableCollection "ExceptWith"
    Set incoming = BuildIncomingItems(other, "ExceptWith")
    For index = mItems.Count To 1 Step -1
        If FindWrappedInCollection(incoming, mItems.Item(index), _
            "ExceptWith") >= 0 Then RemoveCollectionAt index - 1
    Next index
End Sub

Public Sub SymmetricExceptWith(ByVal other As Variant)
Attribute SymmetricExceptWith.VB_Description = _
    "Keeps values present in exactly one of two sets."
    Dim incoming As Collection
    Dim itemIndex As Long
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection "SymmetricExceptWith"
    RequireMutableCollection "SymmetricExceptWith"
    Set incoming = BuildIncomingItems(other, "SymmetricExceptWith")
    Set uniqueValues = DistinctWrappedValues(incoming, "SymmetricExceptWith")
    For Each wrapped In uniqueValues
        itemIndex = FindElementIndex(wrapped, "SymmetricExceptWith")
        If itemIndex < 0 Then
            AddSetValue wrapped, "SymmetricExceptWith"
        Else
            RemoveCollectionAt itemIndex
        End If
    Next wrapped
End Sub

Public Function IsSubsetOf(ByVal other As Variant) As Boolean
Attribute IsSubsetOf.VB_Description = "Tests whether every set value occurs in another sequence."
    IsSubsetOf = IsSetSubset(other, False, "IsSubsetOf")
End Function

Public Function IsProperSubsetOf(ByVal other As Variant) As Boolean
Attribute IsProperSubsetOf.VB_Description = "Tests whether this set is a strict subset."
    IsProperSubsetOf = IsSetSubset(other, True, "IsProperSubsetOf")
End Function

Public Function IsSupersetOf(ByVal other As Variant) As Boolean
Attribute IsSupersetOf.VB_Description = "Tests whether this set contains another sequence."
    IsSupersetOf = IsSetSuperset(other, False, "IsSupersetOf")
End Function

Public Function IsProperSupersetOf(ByVal other As Variant) As Boolean
Attribute IsProperSupersetOf.VB_Description = "Tests whether this set is a strict superset."
    IsProperSupersetOf = IsSetSuperset(other, True, "IsProperSupersetOf")
End Function

Public Function RemoveWhere(ByVal predicate As Variant) As Long
Attribute RemoveWhere.VB_Description = "Removes set values matching a predicate."
    Dim index As Long
    Dim predicateValue As ROneCOne

    RequireSetCollection "RemoveWhere"
    RequireMutableCollection "RemoveWhere"
    Set predicateValue = NormalizeContextUnary(predicate, "RemoveWhere", False)
    For index = mItems.Count To 1 Step -1
        If InvokePredicate(predicateValue, mItems.Item(index), "RemoveWhere") Then
            RemoveCollectionAt index - 1
            RemoveWhere = RemoveWhere + 1
        End If
    Next index
End Function

Public Function TryAddItem(ByVal value As Variant) As Boolean
Attribute TryAddItem.VB_Description = "Attempts to add to a blocking collection."
    RequireBlockingCollection "TryAddItem"
    If mAddingCompleted Then Exit Function
    If mQueryCount >= 0 And mItems.Count >= mQueryCount Then Exit Function
    AddMutable value, , "TryAddItem"
    TryAddItem = True
End Function

Public Sub CompleteAdding()
Attribute CompleteAdding.VB_Description = "Prevents further blocking-collection additions."
    RequireBlockingCollection "CompleteAdding"
    mAddingCompleted = True
End Sub

Public Property Get IsAddingCompleted() As Boolean
Attribute IsAddingCompleted.VB_Description = "Reports whether adding has completed."
    RequireBlockingCollection "IsAddingCompleted"
    IsAddingCompleted = mAddingCompleted
End Property

Public Function TakeItem() As Variant
Attribute TakeItem.VB_Description = "Removes the next blocking-collection value."
    RequireBlockingCollection "TakeItem"
    If mItems.Count = 0 Then RaiseEmptySequence "TakeItem"
    AssignPublicResult TakeItem, mItems.Item(1)
    RemoveCollectionAt 0
End Function

Public Function TryTake(ByVal valueReference As ROneCOne) As Boolean
Attribute TryTake.VB_Description = "Attempts to remove a blocking-collection value."
    RequireBlockingCollection "TryTake"
    RequireReference valueReference, "TryTake"
    If mItems.Count = 0 Then Exit Function
    AssignReferenceWrapped valueReference, mItems.Item(1), "TryTake"
    RemoveCollectionAt 0
    TryTake = True
End Function

Public Function TryPeekPair( _
    ByVal elementReference As ROneCOne, _
    ByVal priorityReference As ROneCOne _
) As Boolean
Attribute TryPeekPair.VB_Description = "Reads the next priority-queue element and priority."
    RequirePriorityQueue "TryPeekPair"
    RequireReference elementReference, "TryPeekPair"
    RequireReference priorityReference, "TryPeekPair"
    If mItems.Count = 0 Then Exit Function
    AssignReferenceWrapped elementReference, mItems.Item(1), "TryPeekPair"
    AssignReferenceWrapped priorityReference, mPriorities.Item(1), "TryPeekPair"
    TryPeekPair = True
End Function

Public Function TryDequeuePair( _
    ByVal elementReference As ROneCOne, _
    ByVal priorityReference As ROneCOne _
) As Boolean
Attribute TryDequeuePair.VB_Description = _
    "Removes the next priority-queue element and priority."
    If Not TryPeekPair(elementReference, priorityReference) Then Exit Function
    RequireMutableCollection "TryDequeuePair"
    RemoveCollectionAt 0
    TryDequeuePair = True
End Function

Public Function DequeueEnqueue( _
    ByVal element As Variant, _
    ByVal priority As Variant _
) As Variant
Attribute DequeueEnqueue.VB_Description = _
    "Dequeues the minimum and then enqueues a replacement."
    Dim removed As ROneCOne

    RequirePriorityQueue "DequeueEnqueue"
    RequireMutableCollection "DequeueEnqueue"
    If mItems.Count = 0 Then RaiseEmptySequence "DequeueEnqueue"
    Set removed = mItems.Item(1)
    RemoveCollectionAt 0
    AddPriorityValue WrapElement(element, "DequeueEnqueue"), _
        WrapKey(priority, "DequeueEnqueue")
    AssignPublicResult DequeueEnqueue, removed
End Function

Public Function EnqueueDequeue( _
    ByVal element As Variant, _
    ByVal priority As Variant _
) As Variant
Attribute EnqueueDequeue.VB_Description = _
    "Enqueues a value and immediately dequeues the minimum."
    Dim removed As ROneCOne

    RequirePriorityQueue "EnqueueDequeue"
    RequireMutableCollection "EnqueueDequeue"
    AddPriorityValue WrapElement(element, "EnqueueDequeue"), _
        WrapKey(priority, "EnqueueDequeue")
    Set removed = mItems.Item(1)
    RemoveCollectionAt 0
    AssignPublicResult EnqueueDequeue, removed
End Function

Public Property Get UnorderedItems() As ROneCOne
Attribute UnorderedItems.VB_Description = "Gets priority-queue element-priority pairs."
    Dim index As Long
    Dim result As ROneCOne

    RequirePriorityQueue "UnorderedItems"
    Set result = ROneCOne.ListOf(vbObject)
    For index = 1 To mItems.Count
        result.Add CreateDictionaryEntry( _
            mItems.Item(index), mPriorities.Item(index))
    Next index
    Set UnorderedItems = result
End Property

Public Function AddBefore( _
    ByVal node As ROneCOne, _
    ByVal value As Variant _
) As ROneCOne
Attribute AddBefore.VB_Description = "Adds a value before a linked-list node."
    Set AddBefore = InsertRelativeNode(node, value, False, "AddBefore")
End Function

Public Function AddAfter( _
    ByVal node As ROneCOne, _
    ByVal value As Variant _
) As ROneCOne
Attribute AddAfter.VB_Description = "Adds a value after a linked-list node."
    Set AddAfter = InsertRelativeNode(node, value, True, "AddAfter")
End Function

Public Property Get Previous() As ROneCOne
Attribute Previous.VB_Description = "Gets the previous linked-list node."
    Dim index As Long

    RequireCollectionNode "Previous"
    index = mSource.FindNodeWrappedIndex(mQueryValue)
    If index > 0 Then Set Previous = mSource.CreateNodeAt(index - 1)
End Property

Public Property Get NextNode() As ROneCOne
Attribute NextNode.VB_Description = "Gets the next linked-list node."
    Dim index As Long

    RequireCollectionNode "NextNode"
    index = mSource.FindNodeWrappedIndex(mQueryValue)
    If index >= 0 And index + 1 < mSource.Count Then
        Set NextNode = mSource.CreateNodeAt(index + 1)
    End If
End Property

Public Property Get List() As ROneCOne
Attribute List.VB_Description = "Gets the linked list that owns this node."
    RequireCollectionNode "List"
    Set List = mSource
End Property

Public Function FindNode(ByVal value As Variant) As ROneCOne
Attribute FindNode.VB_Description = "Finds the first linked-list node with a value."
    Dim index As Long

    RequireLinkedList "FindNode"
    index = FindElementIndex(WrapElement(value, "FindNode"), "FindNode")
    If index >= 0 Then Set FindNode = CreateCollectionNode(index)
End Function

Public Function FindLastNode(ByVal value As Variant) As ROneCOne
Attribute FindLastNode.VB_Description = "Finds the last linked-list node with a value."
    Dim candidate As ROneCOne
    Dim index As Long

    RequireLinkedList "FindLastNode"
    Set candidate = WrapElement(value, "FindLastNode")
    For index = mItems.Count To 1 Step -1
        If ValuesEqualWithComparer(mItems.Item(index), candidate, _
            mQueryComparer, "FindLastNode") Then
            Set FindLastNode = CreateCollectionNode(index - 1)
            Exit Function
        End If
    Next index
End Function

Public Sub RemoveNode(ByVal node As ROneCOne)
Attribute RemoveNode.VB_Description = "Removes an owned linked-list node."
    Dim index As Long

    RequireLinkedList "RemoveNode"
    RequireMutableCollection "RemoveNode"
    RequireOwnedNode node, "RemoveNode"
    index = FindNodeWrappedIndex(node.InternalQueryValue)
    If index < 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, "RemoveNode", _
            "The node no longer belongs to this list."
    End If
    RemoveCollectionAt index
End Sub

Public Function AddRangeImmutable(ByVal values As Variant) As ROneCOne
Attribute AddRangeImmutable.VB_Description = "Returns an immutable collection with added values."
    Dim incoming As Collection
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireImmutableCollection "AddRangeImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Set incoming = BuildIncomingItems(values, "AddRangeImmutable")
    For Each wrapped In incoming
        result.AddMutable Unwrap(wrapped), memberName:="AddRangeImmutable"
    Next wrapped
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set AddRangeImmutable = result
End Function

Public Function SetItem( _
    ByVal keyOrIndex As Variant, _
    ByVal value As Variant _
) As ROneCOne
Attribute SetItem.VB_Description = "Returns an immutable collection with one value replaced."
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim result As ROneCOne

    RequireImmutableCollection "SetItem"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    If result.IsDictionaryCollection Then
        Set keyWrapped = result.WrapKey(keyOrIndex, "SetItem")
        itemIndex = result.FindKeyIndex(keyWrapped, "SetItem")
        If itemIndex < 0 Then
            result.AddDictionaryPair keyWrapped, _
                result.WrapElement(value, "SetItem"), "SetItem"
        Else
            result.ReplaceCollectionItem itemIndex, _
                result.WrapElement(value, "SetItem")
        End If
    Else
        itemIndex = CLng(keyOrIndex)
        ValidateIndex itemIndex, result.Count, "SetItem"
        result.ReplaceCollectionItem itemIndex, _
            result.WrapElement(value, "SetItem")
    End If
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set SetItem = result
End Function

Public Function RemoveRange(ByVal values As Variant) As ROneCOne
Attribute RemoveRange.VB_Description = "Returns an immutable collection without selected values."
    Dim incoming As Collection
    Dim result As ROneCOne
    Dim wrapped As ROneCOne

    RequireImmutableCollection "RemoveRange"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    Set incoming = BuildIncomingItems(values, "RemoveRange")
    For Each wrapped In incoming
        RemoveCollectionValue result, Unwrap(wrapped), "RemoveRange"
    Next wrapped
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set RemoveRange = result
End Function

Public Function DequeueImmutable() As ROneCOne
Attribute DequeueImmutable.VB_Description = "Returns an immutable queue without its first value."
    Dim result As ROneCOne

    RequireImmutableCollection "DequeueImmutable"
    If Not IsQueueKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "DequeueImmutable", _
            "DequeueImmutable requires an immutable queue."
    End If
    If mItems.Count = 0 Then RaiseEmptySequence "DequeueImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    result.RemoveCollectionAt 0
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set DequeueImmutable = result
End Function

Public Function PopImmutable() As ROneCOne
Attribute PopImmutable.VB_Description = "Returns an immutable stack without its top value."
    Dim result As ROneCOne

    RequireImmutableCollection "PopImmutable"
    If Not IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, "PopImmutable", _
            "PopImmutable requires an immutable stack."
    End If
    If mItems.Count = 0 Then RaiseEmptySequence "PopImmutable"
    Set result = CloneGenericCollection(Me)
    result.ConfigureCollectionRole MutableRoleForKind(mCollectionKind)
    result.RemoveCollectionAt 0
    result.ConfigureCollectionRole ROLE_IMMUTABLE
    Set PopImmutable = result
End Function

Public Property Get Element() As ROneCOne
Attribute Element.VB_Description = "Gets a typed parameter representing each sequence element."
    RequireSequence Me, "Element"
    If mElementParameter Is Nothing Then
        Set mElementParameter = New ROneCOne
        mElementParameter.ConfigureParameter mElementType, mElementObjectType
    End If
    Set Element = mElementParameter
End Property

Public Function AsEnumerable() As ROneCOne
Attribute AsEnumerable.VB_Description = "Returns this sequence through its enumerable surface."
    RequireSequence Me, "AsEnumerable"
    Set AsEnumerable = Me
End Function

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

Public Function TakeWhile(ByVal predicate As Variant) As ROneCOne
Attribute TakeWhile.VB_Description = "Takes elements while a predicate remains true."
    Set TakeWhile = CreateQuery( _
        QUERY_TAKE_WHILE, NormalizeContextUnary(predicate, "TakeWhile", False), _
        Nothing, 0, InternalElementType, InternalElementObjectType)
End Function

Public Function SkipWhile(ByVal predicate As Variant) As ROneCOne
Attribute SkipWhile.VB_Description = "Skips elements while a predicate remains true."
    Set SkipWhile = CreateQuery( _
        QUERY_SKIP_WHILE, NormalizeContextUnary(predicate, "SkipWhile", False), _
        Nothing, 0, InternalElementType, InternalElementObjectType)
End Function

Public Function TakeLast(ByVal count As Long) As ROneCOne
Attribute TakeLast.VB_Description = "Takes the final count elements."
    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "TakeLast", _
            "Count cannot be negative."
    End If
    Set TakeLast = CreateQuery(QUERY_TAKE_LAST, Nothing, Nothing, count, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function SkipLast(ByVal count As Long) As ROneCOne
Attribute SkipLast.VB_Description = "Skips the final count elements."
    If count < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "SkipLast", _
            "Count cannot be negative."
    End If
    Set SkipLast = CreateQuery(QUERY_SKIP_LAST, Nothing, Nothing, count, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function ConcatSequence(ByVal other As Variant) As ROneCOne
Attribute ConcatSequence.VB_Description = "Concatenates two compatible sequences."
    Dim otherSequence As ROneCOne

    Set otherSequence = NormalizeMembershipSequence(other, "Concat")
    RequireCompatibleSequence otherSequence, "Concat"
    Set ConcatSequence = CreateQuery(QUERY_CONCAT, Nothing, otherSequence, 0, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function Union( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Union.VB_Description = "Produces the set union of two sequences."
    Set Union = CreateSetQuery(QUERY_UNION, other, comparer, "Union")
End Function

Public Function Intersect( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Intersect.VB_Description = "Produces the set intersection of two sequences."
    Set Intersect = CreateSetQuery(QUERY_INTERSECT, other, comparer, "Intersect")
End Function

Public Function Except( _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute Except.VB_Description = "Produces values absent from the second sequence."
    Set Except = CreateSetQuery(QUERY_EXCEPT, other, comparer, "Except")
End Function

Public Function DefaultIfEmpty(Optional ByVal defaultValue As Variant) As ROneCOne
Attribute DefaultIfEmpty.VB_Description = "Supplies one default when a sequence is empty."
    Dim wrappedDefault As ROneCOne

    If Not IsMissing(defaultValue) Then
        Set wrappedDefault = WrapElement(defaultValue, "DefaultIfEmpty")
    End If
    Set DefaultIfEmpty = CreateQuery( _
        QUERY_DEFAULT_IF_EMPTY, Nothing, wrappedDefault, 0, _
        InternalElementType, InternalElementObjectType)
End Function

Public Function Chunk(ByVal size As Long) As ROneCOne
Attribute Chunk.VB_Description = "Splits a sequence into typed List<T> chunks."
    If size <= 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Chunk", _
            "Chunk size must be positive."
    End If
    Set Chunk = CreateQuery( _
        QUERY_CHUNK, Nothing, Nothing, size, vbObject, "ROneCOne")
End Function

Public Function SelectMany( _
    ByVal selector As Variant, _
    Optional ByVal resultType As Variant _
) As ROneCOne
Attribute SelectMany.VB_Description = "Flattens nested sequences through a selector."
    Dim objectType As String
    Dim valueType As VbVarType

    If IsMissing(resultType) Then
        valueType = vbVariant
    Else
        ResolveTypeDescriptor resultType, valueType, objectType, "SelectMany"
    End If
    Set SelectMany = CreateQuery( _
        QUERY_SELECT_MANY, NormalizeContextUnary(selector, "SelectMany", False), _
        Nothing, 0, valueType, objectType)
End Function

Public Function Cast(ByVal typeToken As Variant) As ROneCOne
Attribute Cast.VB_Description = "Casts every sequence value to the requested runtime type."
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "Cast"
    Set Cast = CreateQuery( _
        QUERY_CAST, Nothing, Nothing, 0, valueType, objectType)
End Function

Public Function OfType(ByVal typeToken As Variant) As ROneCOne
Attribute OfType.VB_Description = "Filters values assignable to the requested runtime type."
    Dim objectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor typeToken, valueType, objectType, "OfType"
    Set OfType = CreateQuery( _
        QUERY_OF_TYPE, Nothing, Nothing, 0, valueType, objectType)
End Function

Public Function UnionBy( _
    ByVal other As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute UnionBy.VB_Description = "Produces a union using selected distinct keys."
    Set UnionBy = CreateSetByQuery( _
        QUERY_UNION_BY, other, keySelector, comparer, "UnionBy", True)
End Function

Public Function IntersectBy( _
    ByVal keys As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute IntersectBy.VB_Description = "Keeps source values whose selected key is present."
    Set IntersectBy = CreateSetByQuery( _
        QUERY_INTERSECT_BY, keys, keySelector, comparer, _
        "IntersectBy", False)
End Function

Public Function ExceptBy( _
    ByVal keys As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ExceptBy.VB_Description = "Keeps source values whose selected key is absent."
    Set ExceptBy = CreateSetByQuery( _
        QUERY_EXCEPT_BY, keys, keySelector, comparer, "ExceptBy", False)
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
    Set OrderBy = CreatePrimaryOrderQuery( _
        QUERY_ORDER_BY, selectorDelegate, _
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
    Set OrderByDescending = CreatePrimaryOrderQuery( _
        QUERY_ORDER_BY_DESCENDING, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function Order(Optional ByVal comparer As Variant) As ROneCOne
Attribute Order.VB_Description = "Defers stable ascending ordering by each element."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary(Element, "Order", False)
    If IsMissing(comparer) Then
        Set comparerValue = Nothing
    Else
        Set comparerValue = NormalizeComparer(comparer, False, "Order")
    End If
    Set Order = CreatePrimaryOrderQuery( _
        QUERY_ORDER, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function OrderDescending(Optional ByVal comparer As Variant) As ROneCOne
Attribute OrderDescending.VB_Description = "Defers stable descending ordering by each element."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    Set selectorDelegate = NormalizeContextUnary( _
        Element, "OrderDescending", False)
    If IsMissing(comparer) Then
        Set comparerValue = Nothing
    Else
        Set comparerValue = NormalizeComparer( _
            comparer, False, "OrderDescending")
    End If
    Set OrderDescending = CreatePrimaryOrderQuery( _
        QUERY_ORDER_DESCENDING, selectorDelegate, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function ThenBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ThenBy.VB_Description = "Adds a stable ascending key to an active ordered query."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    RequireOrderedQuery Me, "ThenBy"
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "ThenBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, False, "ThenBy")
    End If
    Set ThenBy = CreateQuery(QUERY_THEN_BY, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
End Function

Public Function ThenByDescending( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ThenByDescending.VB_Description = "Adds a stable descending key to an ordered query."
    Dim comparerValue As ROneCOne
    Dim selectorDelegate As ROneCOne

    RequireOrderedQuery Me, "ThenByDescending"
    Set selectorDelegate = NormalizeContextUnary( _
        keySelector, "ThenByDescending", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, False, "ThenByDescending")
    End If
    Set ThenByDescending = CreateQuery( _
        QUERY_THEN_BY_DESCENDING, selectorDelegate, Nothing, 0, _
        mElementType, mElementObjectType, comparerValue)
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

Public Function ElementAt(ByVal index As Long) As Variant
Attribute ElementAt.VB_Description = "Gets an element by zero-based index."
    Dim values As Collection

    Set values = MaterializeWrapped
    ValidateIndex index, values.Count, "ElementAt"
    AssignPublicResult ElementAt, values.Item(index + 1)
End Function

Public Function ElementAtOrDefault(ByVal index As Long) As Variant
Attribute ElementAtOrDefault.VB_Description = "Gets an element or the type default."
    Dim values As Collection

    Set values = MaterializeWrapped
    If index < 0 Or index >= values.Count Then
        AssignDefaultPublicResult ElementAtOrDefault
    Else
        AssignPublicResult ElementAtOrDefault, values.Item(index + 1)
    End If
End Function

Public Function Index() As ROneCOne
Attribute Index.VB_Description = "Pairs every element with its zero-based index."
    Dim indexValue As Long
    Dim result As ROneCOne
    Dim values As Collection

    Set result = ROneCOne.ListOf(vbObject)
    Set values = MaterializeWrapped
    For indexValue = 1 To values.Count
        result.Add CreateDictionaryEntry( _
            CreateScalarResult(CLng(indexValue - 1)), values.Item(indexValue))
    Next indexValue
    Set Index = result
End Function

Public Function Aggregate( _
    ByVal seedOrAccumulator As Variant, _
    Optional ByVal accumulator As Variant, _
    Optional ByVal resultSelector As Variant _
) As Variant
Attribute Aggregate.VB_Description = "Applies an accumulator across a sequence."
    Dim accumulatorValue As ROneCOne
    Dim current As ROneCOne
    Dim index As Long
    Dim resultSelectorValue As ROneCOne
    Dim values As Collection

    Set values = MaterializeWrapped
    If IsMissing(accumulator) Then
        If values.Count = 0 Then RaiseEmptySequence "Aggregate"
        Set accumulatorValue = NormalizeBinaryDelegate( _
            seedOrAccumulator, "Aggregate")
        Set current = values.Item(1)
        index = 2
    Else
        Set current = WrapInvocationArgument(seedOrAccumulator)
        Set accumulatorValue = NormalizeBinaryDelegate(accumulator, "Aggregate")
        index = 1
    End If
    For index = index To values.Count
        Set current = InvokeBinary(accumulatorValue, current, values.Item(index))
    Next index
    If Not IsMissing(resultSelector) Then
        Set resultSelectorValue = NormalizeStandaloneUnary( _
            RequireROneCOne(resultSelector, "Aggregate"), "Aggregate")
        Set current = InvokeUnary(resultSelectorValue, current)
    End If
    AssignPublicResult Aggregate, current
End Function

Public Function CountBy( _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute CountBy.VB_Description = "Counts elements grouped by a selected key."
    Dim comparerValue As ROneCOne
    Dim itemIndex As Long
    Dim key As ROneCOne
    Dim result As ROneCOne
    Dim selectorValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "CountBy", False)
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "CountBy")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        vbLong, vbNullString, vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(selectorValue, wrapped)
        itemIndex = result.FindKeyIndex(key, "CountBy")
        If itemIndex < 0 Then
            result.AddDictionaryPair key, CreateScalarResult(CLng(1)), "CountBy"
        Else
            result.ReplaceCollectionItem itemIndex, CreateScalarResult( _
                CLng(result.InternalCollectionItem(itemIndex).InternalScalarValue) + 1)
        End If
    Next wrapped
    Set CountBy = result
End Function

Public Function AggregateBy( _
    ByVal keySelector As Variant, _
    ByVal seed As Variant, _
    ByVal accumulator As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute AggregateBy.VB_Description = "Aggregates elements independently for each key."
    Dim accumulatorValue As ROneCOne
    Dim comparerValue As ROneCOne
    Dim current As ROneCOne
    Dim itemIndex As Long
    Dim key As ROneCOne
    Dim result As ROneCOne
    Dim selectorValue As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, "AggregateBy", False)
    Set accumulatorValue = NormalizeBinaryDelegate(accumulator, "AggregateBy")
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "AggregateBy")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        vbVariant, vbNullString, vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(selectorValue, wrapped)
        itemIndex = result.FindKeyIndex(key, "AggregateBy")
        If itemIndex < 0 Then
            Set current = WrapInvocationArgument(seed)
        Else
            Set current = result.InternalCollectionItem(itemIndex)
        End If
        Set current = InvokeBinary(accumulatorValue, current, wrapped)
        If itemIndex < 0 Then
            result.AddDictionaryPair key, current, "AggregateBy"
        Else
            result.ReplaceCollectionItem itemIndex, current
        End If
    Next wrapped
    Set AggregateBy = result
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

Public Function ToDictionary( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ToDictionary.VB_Description = "Materializes a Dictionary from selected keys."
    Dim comparerValue As ROneCOne
    Dim elementSelectorValue As ROneCOne
    Dim elementValue As ROneCOne
    Dim key As ROneCOne
    Dim keySelectorValue As ROneCOne
    Dim result As ROneCOne
    Dim resultObjectType As String
    Dim resultType As VbVarType
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set keySelectorValue = NormalizeContextUnary( _
        keySelector, "ToDictionary", False)
    resultType = InternalElementType
    resultObjectType = InternalElementObjectType
    If Not IsMissing(elementSelector) Then
        Set elementSelectorValue = NormalizeContextUnary( _
            elementSelector, "ToDictionary", False)
        resultType = vbVariant
        resultObjectType = vbNullString
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToDictionary")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_DICTIONARY, COLLECTION_DICTIONARY, _
        resultType, resultObjectType, vbVariant, _
        vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(keySelectorValue, wrapped)
        If result.FindKeyIndex(key, "ToDictionary") >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, "ToDictionary", _
                "The selected key is duplicated."
        End If
        If elementSelectorValue Is Nothing Then
            Set elementValue = wrapped
        Else
            Set elementValue = InvokeUnary(elementSelectorValue, wrapped)
        End If
        result.AddDictionaryPair key, elementValue, "ToDictionary"
    Next wrapped
    Set ToDictionary = result
End Function

Public Function ToHashSet(Optional ByVal comparer As Variant) As ROneCOne
Attribute ToHashSet.VB_Description = "Materializes a HashSet<T>."
    Dim comparerValue As ROneCOne
    Dim result As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToHashSet")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_HASH_SET, COLLECTION_HASH_SET, _
        InternalElementType, InternalElementObjectType, vbEmpty, _
        vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        If result.FindElementIndex(wrapped, "ToHashSet") < 0 Then
            result.AddSetValue wrapped, "ToHashSet"
        End If
    Next wrapped
    Set ToHashSet = result
End Function

Public Function ToLookup( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute ToLookup.VB_Description = "Materializes a one-to-many Lookup<TKey, TElement>."
    Dim comparerValue As ROneCOne
    Dim elementSelectorValue As ROneCOne
    Dim elementValue As ROneCOne
    Dim group As ROneCOne
    Dim groupIndex As Long
    Dim key As ROneCOne
    Dim keySelectorValue As ROneCOne
    Dim lookup As ROneCOne
    Dim values As Collection
    Dim wrapped As ROneCOne

    Set keySelectorValue = NormalizeContextUnary( _
        keySelector, "ToLookup", False)
    If Not IsMissing(elementSelector) Then
        Set elementSelectorValue = NormalizeContextUnary( _
            elementSelector, "ToLookup", False)
    End If
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, "ToLookup")
    End If
    Set lookup = New ROneCOne
    lookup.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_LOOKUP, _
        vbObject, "ROneCOne", vbVariant, vbNullString, comparerValue
    Set values = MaterializeWrapped
    For Each wrapped In values
        Set key = InvokeUnary(keySelectorValue, wrapped)
        If elementSelectorValue Is Nothing Then
            Set elementValue = wrapped
        Else
            Set elementValue = InvokeUnary(elementSelectorValue, wrapped)
        End If
        groupIndex = lookup.FindKeyIndex(key, "ToLookup")
        If groupIndex < 0 Then
            Set group = New ROneCOne
            group.ConfigureGrouping key, elementValue
            lookup.InternalAddLookupGroup key, group
        Else
            Set group = lookup.InternalCollectionItem(groupIndex)
            group.InternalAppendCollectionItem elementValue
        End If
    Next wrapped
    Set ToLookup = lookup
End Function

Public Function GroupBy( _
    ByVal keySelector As Variant, _
    Optional ByVal elementSelector As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
Attribute GroupBy.VB_Description = "Groups elements by a selected key."
    Set GroupBy = ToLookup(keySelector, elementSelector, comparer).ToList
End Function

Public Function Join( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant _
) As ROneCOne
Attribute Join.VB_Description = "Correlates two sequences by matching keys."
    Set Join = ExecuteJoin( _
        inner, outerKeySelector, innerKeySelector, resultSelector, False)
End Function

Public Function GroupJoin( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant _
) As ROneCOne
Attribute GroupJoin.VB_Description = "Correlates outer elements with grouped inner matches."
    Set GroupJoin = ExecuteJoin( _
        inner, outerKeySelector, innerKeySelector, resultSelector, True)
End Function

Public Function Zip( _
    ByVal second As Variant, _
    Optional ByVal resultSelector As Variant _
) As ROneCOne
Attribute Zip.VB_Description = "Combines elements at corresponding positions."
    Dim firstValues As Collection
    Dim index As Long
    Dim itemCount As Long
    Dim result As ROneCOne
    Dim resultSelectorValue As ROneCOne
    Dim secondSequence As ROneCOne
    Dim secondValues As Collection
    Dim selected As ROneCOne

    Set secondSequence = NormalizeMembershipSequence(second, "Zip")
    Set firstValues = MaterializeWrapped
    Set secondValues = secondSequence.MaterializeWrapped
    itemCount = firstValues.Count
    If secondValues.Count < itemCount Then itemCount = secondValues.Count
    Set result = ROneCOne.ListOf(vbVariant)
    If Not IsMissing(resultSelector) Then
        Set resultSelectorValue = NormalizeBinaryDelegate(resultSelector, "Zip")
    End If
    For index = 1 To itemCount
        If resultSelectorValue Is Nothing Then
            Set selected = CreateDictionaryEntry( _
                firstValues.Item(index), secondValues.Item(index))
            result.Add selected
        Else
            Set selected = InvokeBinary(resultSelectorValue, _
                firstValues.Item(index), secondValues.Item(index))
            result.InternalAddWrapped selected
        End If
    Next index
    Set Zip = result
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets the VBA For Each enumerator for this sequence."
Attribute NewEnum.VB_UserMemId = -4
    If mRole = ROLE_QUERY Or mRole = ROLE_DATA_VIEW Or _
        IsGenericCollectionRole(mRole) Then
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

Friend Property Get InternalSource() As ROneCOne
Attribute InternalSource.VB_MemberFlags = "40"
    Set InternalSource = mSource
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

Friend Property Get InternalCollectionKind() As Long
Attribute InternalCollectionKind.VB_MemberFlags = "40"
    InternalCollectionKind = mCollectionKind
End Property

Friend Property Get InternalCollectionItems() As Collection
Attribute InternalCollectionItems.VB_MemberFlags = "40"
    Set InternalCollectionItems = mItems
End Property

Friend Property Get InternalCollectionKeys() As Collection
Attribute InternalCollectionKeys.VB_MemberFlags = "40"
    Set InternalCollectionKeys = mKeys
End Property

Friend Property Get InternalCollectionPriorities() As Collection
Attribute InternalCollectionPriorities.VB_MemberFlags = "40"
    Set InternalCollectionPriorities = mPriorities
End Property

Friend Property Get InternalKeyType() As VbVarType
Attribute InternalKeyType.VB_MemberFlags = "40"
    InternalKeyType = mKeyType
End Property

Friend Property Get InternalKeyObjectType() As String
Attribute InternalKeyObjectType.VB_MemberFlags = "40"
    InternalKeyObjectType = mKeyObjectType
End Property

Friend Property Get InternalKeySelector() As ROneCOne
Attribute InternalKeySelector.VB_MemberFlags = "40"
    Set InternalKeySelector = mKeySelector
End Property

Friend Property Get InternalCollectionChanged() As ROneCOne
Attribute InternalCollectionChanged.VB_MemberFlags = "40"
    Set InternalCollectionChanged = mCollectionChanged
End Property

Friend Property Get InternalCollectionCapacity() As Long
Attribute InternalCollectionCapacity.VB_MemberFlags = "40"
    InternalCollectionCapacity = mQueryCount
End Property

Friend Property Get InternalCollectionVersion() As Long
Attribute InternalCollectionVersion.VB_MemberFlags = "40"
    InternalCollectionVersion = mVersion
End Property

Friend Property Get InternalIsBuilder() As Boolean
Attribute InternalIsBuilder.VB_MemberFlags = "40"
    InternalIsBuilder = mIsBuilder
End Property

Friend Property Get InternalName() As String
Attribute InternalName.VB_MemberFlags = "40"
    InternalName = mName
End Property

Friend Property Get InternalDataColumns() As Collection
Attribute InternalDataColumns.VB_MemberFlags = "40"
    Set InternalDataColumns = mKeys
End Property

Friend Property Get InternalDataRows() As Collection
Attribute InternalDataRows.VB_MemberFlags = "40"
    Set InternalDataRows = mItems
End Property

Friend Property Get InternalDataTables() As Collection
Attribute InternalDataTables.VB_MemberFlags = "40"
    Set InternalDataTables = mDataTables
End Property

Friend Property Get InternalDataRelations() As Collection
Attribute InternalDataRelations.VB_MemberFlags = "40"
    Set InternalDataRelations = mDataRelations
End Property

Friend Property Get InternalEnforceConstraints() As Boolean
Attribute InternalEnforceConstraints.VB_MemberFlags = "40"
    InternalEnforceConstraints = mUnique
End Property

Friend Property Get InternalUnique() As Boolean
Attribute InternalUnique.VB_MemberFlags = "40"
    InternalUnique = mUnique
End Property

Friend Property Get InternalAllowDBNull() As Boolean
Attribute InternalAllowDBNull.VB_MemberFlags = "40"
    InternalAllowDBNull = mAllowDBNull
End Property

Friend Property Get InternalRelationParentColumn() As ROneCOne
Attribute InternalRelationParentColumn.VB_MemberFlags = "40"
    Set InternalRelationParentColumn = mLeft
End Property

Friend Property Get InternalRelationChildColumn() As ROneCOne
Attribute InternalRelationChildColumn.VB_MemberFlags = "40"
    Set InternalRelationChildColumn = mRight
End Property

Friend Property Get InternalProviderObject() As Object
Attribute InternalProviderObject.VB_MemberFlags = "40"
    Set InternalProviderObject = mProviderObject
End Property

Friend Property Get InternalProviderType() As Long
Attribute InternalProviderType.VB_MemberFlags = "40"
    InternalProviderType = mProviderType
End Property

Friend Property Get InternalProviderSize() As Long
Attribute InternalProviderSize.VB_MemberFlags = "40"
    InternalProviderSize = mProviderSize
End Property

Friend Property Get InternalProviderPrecision() As Byte
Attribute InternalProviderPrecision.VB_MemberFlags = "40"
    InternalProviderPrecision = mProviderPrecision
End Property

Friend Property Get InternalProviderScale() As Byte
Attribute InternalProviderScale.VB_MemberFlags = "40"
    InternalProviderScale = mProviderScale
End Property

Friend Property Get InternalSourceColumn() As String
Attribute InternalSourceColumn.VB_MemberFlags = "40"
    InternalSourceColumn = mSourceColumn
End Property

Friend Property Get InternalSourceVersion() As String
Attribute InternalSourceVersion.VB_MemberFlags = "40"
    InternalSourceVersion = mSourceVersion
End Property

Friend Property Get InternalElementType() As VbVarType
Attribute InternalElementType.VB_MemberFlags = "40"
    If IsDictionaryCollection Then
        InternalElementType = vbObject
    Else
        InternalElementType = mElementType
    End If
End Property

Friend Property Get InternalElementObjectType() As String
Attribute InternalElementObjectType.VB_MemberFlags = "40"
    If IsDictionaryCollection Then
        InternalElementObjectType = "ROneCOne"
    Else
        InternalElementObjectType = mElementObjectType
    End If
End Property

Friend Property Get InternalValueType() As VbVarType
Attribute InternalValueType.VB_MemberFlags = "40"
    InternalValueType = mElementType
End Property

Friend Property Get InternalValueObjectType() As String
Attribute InternalValueObjectType.VB_MemberFlags = "40"
    InternalValueObjectType = mElementObjectType
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

Friend Sub ConfigureManagedReference( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mRole = ROLE_REFERENCE
    mExpectedType = valueType
    mExpectedObjectType = objectType
    mReferenceAddress = 0
End Sub

Friend Sub ConfigureManagedReferenceValue(ByVal wrapped As ROneCOne)
    Set mQueryValue = wrapped
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

Friend Sub ConfigureGenericCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal keyType As VbVarType = vbEmpty, _
    Optional ByVal keyObjectType As String = vbNullString, _
    Optional ByVal comparer As ROneCOne _
)
    mRole = roleValue
    mCollectionKind = collectionKind
    mElementType = elementType
    mElementObjectType = elementObjectType
    mKeyType = keyType
    mKeyObjectType = keyObjectType
    Set mQueryComparer = comparer
    Set mItems = New Collection
    Set mKeys = New Collection
    Set mPriorities = New Collection
    Set mEnumerationValues = New Collection
End Sub

Friend Sub ConfigureReadOnlySource(ByVal source As ROneCOne)
    Set mSource = source
End Sub

Friend Sub ConfigureCollectionKeySelector(ByVal selector As ROneCOne)
    Set mKeySelector = selector
End Sub

Friend Sub ConfigureObservableCollection()
    Set mCollectionChanged = ROneCOne.EventOf(vbString)
End Sub

Friend Sub ConfigureCollectionCapacity(ByVal boundedCapacity As Long)
    mQueryCount = boundedCapacity
End Sub

Friend Sub ConfigureCollectionRole(ByVal roleValue As Long)
    mRole = roleValue
End Sub

Friend Sub ConfigureCollectionBuilder(ByVal isBuilder As Boolean)
    mIsBuilder = isBuilder
End Sub

Friend Sub ConfigureCollectionVersion(ByVal version As Long)
    mVersion = version
End Sub

Friend Sub ConfigureCollectionNode( _
    ByVal wrapped As ROneCOne, _
    ByVal owner As ROneCOne _
)
    mRole = ROLE_NODE
    Set mQueryValue = wrapped
    Set mSource = owner
End Sub

Friend Sub ConfigureDictionaryEntry( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne _
)
    mRole = ROLE_ENTRY
    Set mLeft = keyWrapped
    Set mQueryValue = valueWrapped
End Sub

Friend Sub ConfigureTaskFactory()
    mRole = ROLE_TASK_FACTORY
End Sub

Friend Sub ConfigureTask( _
    ByVal taskKind As Long, _
    ByVal taskBody As ROneCOne, _
    ByVal cancellationToken As ROneCOne _
)
    mRole = ROLE_TASK
    mOperator = taskKind
    Set mTaskBody = taskBody
    Set mCancellationSource = cancellationToken
    Set mTaskChildren = New Collection
    mTaskStatus = TASK_STATUS_CREATED
End Sub

Friend Sub ConfigureTaskArguments(ByVal arguments As Collection)
    Set mTaskChildren = arguments
End Sub

Friend Sub ConfigureProviderTask( _
    ByVal providerTarget As ROneCOne, _
    ByVal methodName As String, _
    ByVal arguments As Collection, _
    ByVal cancellationToken As ROneCOne _
)
    ConfigureTask TASK_PROVIDER_CALL, Nothing, cancellationToken
    Set mSource = providerTarget
    mMethodName = methodName
    Set mTaskChildren = arguments
End Sub

Friend Sub ConfigureCompletedTask(ByVal result As ROneCOne)
    mRole = ROLE_TASK
    mOperator = TASK_COMPLETED
    Set mTaskResult = result
    Set mTaskChildren = New Collection
    mTaskStatus = TASK_STATUS_COMPLETED
End Sub

Friend Sub ConfigureTaskCompletionSource( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    Dim promise As ROneCOne

    mRole = ROLE_TASK_COMPLETION_SOURCE
    mExpectedType = valueType
    mExpectedObjectType = objectType
    Set promise = New ROneCOne
    promise.ConfigureTask TASK_PROMISE, Nothing, Nothing
    Set mSource = promise
End Sub

Friend Sub ConfigureProgress( _
    ByVal valueType As VbVarType, _
    ByVal objectType As String, _
    ByVal handler As ROneCOne _
)
    mRole = ROLE_PROGRESS
    mExpectedType = valueType
    mExpectedObjectType = objectType
    Set mTaskBody = handler
End Sub

Friend Sub ConfigureCancellationSource()
    mRole = ROLE_CANCELLATION_SOURCE
    mCancellationRequested = False
    mQueryCount = -1
    Set mInvocationList = New Collection
End Sub

Friend Sub ConfigureCancellationToken(ByVal source As ROneCOne)
    mRole = ROLE_CANCELLATION_TOKEN
    Set mCancellationSource = source
End Sub

Friend Sub RegisterCancellationCallback(ByVal callback As ROneCOne)
    mInvocationList.Add callback
End Sub

Friend Function TryCompletePromise( _
    ByVal result As ROneCOne, _
    ByVal terminalStatus As Long, _
    ByVal errorValue As ROneCOne _
) As Boolean
    If mRole <> ROLE_TASK Or mOperator <> TASK_PROMISE Then Exit Function
    If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Function
    Set mTaskResult = result
    Set mTaskException = errorValue
    mTaskStatus = terminalStatus
    TryCompletePromise = True
End Function

Friend Sub ConfigureDataTable(ByVal tableName As String)
    mRole = ROLE_DATA_TABLE
    mName = tableName
    Set mItems = New Collection
    Set mKeys = New Collection
    Set mPrimaryKey = New Collection
End Sub

Friend Sub ConfigureDataColumn( _
    ByVal columnName As String, _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    mRole = ROLE_DATA_COLUMN
    mName = columnName
    mExpectedType = valueType
    mExpectedObjectType = objectType
    mAllowDBNull = True
End Sub

Friend Sub ConfigureDataSet(ByVal dataSetName As String)
    mRole = ROLE_DATA_SET
    mName = dataSetName
    Set mDataTables = New Collection
    Set mDataRelations = New Collection
    mUnique = True
End Sub

Friend Sub ConfigureDataRelation( _
    ByVal relationName As String, _
    ByVal parentColumn As ROneCOne, _
    ByVal childColumn As ROneCOne, _
    ByVal createConstraints As Boolean _
)
    mRole = ROLE_DATA_RELATION
    mName = relationName
    Set mLeft = parentColumn
    Set mRight = childColumn
    mExpressionOption = Abs(CLng(createConstraints))
End Sub

Friend Sub ConfigureDataView(ByVal table As ROneCOne)
    mRole = ROLE_DATA_VIEW
    Set mSource = table
    mElementType = vbObject
    mElementObjectType = "ROneCOne"
End Sub

Friend Sub ConfigureDataName(ByVal value As String)
    mName = value
End Sub

Friend Sub ConfigureDataRow(ByVal table As ROneCOne)
    Dim column As ROneCOne

    mRole = ROLE_DATA_ROW
    Set mSource = table
    Set mItems = New Collection
    Set mOriginalItems = New Collection
    For Each column In table.InternalDataColumns
        mItems.Add DefaultDataColumnValue(column)
    Next column
    mRowState = ROW_STATE_DETACHED
End Sub

Friend Sub ConfigureRowState(ByVal rowState As Long)
    mRowState = rowState
End Sub

Friend Sub TouchDataVersion()
    mVersion = mVersion + 1
End Sub

Friend Sub ConfigureUniqueColumn()
    mUnique = True
    mAllowDBNull = False
End Sub

Friend Property Get InternalAutoIncrement() As Boolean
Attribute InternalAutoIncrement.VB_MemberFlags = "40"
    InternalAutoIncrement = mAutoIncrement
End Property

Friend Property Get InternalDefaultValue() As ROneCOne
Attribute InternalDefaultValue.VB_MemberFlags = "40"
    Set InternalDefaultValue = mDefaultValue
End Property

Friend Property Get InternalAutoIncrementSeed() As LongLong
Attribute InternalAutoIncrementSeed.VB_MemberFlags = "40"
    InternalAutoIncrementSeed = mAutoIncrementSeed
End Property

Friend Property Get InternalAutoIncrementStep() As LongLong
Attribute InternalAutoIncrementStep.VB_MemberFlags = "40"
    InternalAutoIncrementStep = mAutoIncrementStep
End Property

Friend Property Get InternalAutoIncrementCurrent() As LongLong
Attribute InternalAutoIncrementCurrent.VB_MemberFlags = "40"
    InternalAutoIncrementCurrent = mAutoIncrementCurrent
End Property

Friend Property Get InternalPrimaryKey() As Collection
Attribute InternalPrimaryKey.VB_MemberFlags = "40"
    Set InternalPrimaryKey = mPrimaryKey
End Property

Friend Sub CopyDataColumnConfiguration(ByVal source As ROneCOne)
    mAllowDBNull = source.InternalAllowDBNull
    mUnique = source.InternalUnique
    mAutoIncrement = source.InternalAutoIncrement
    mAutoIncrementSeed = source.InternalAutoIncrementSeed
    mAutoIncrementStep = source.InternalAutoIncrementStep
    mAutoIncrementCurrent = source.InternalAutoIncrementCurrent
    Set mDefaultValue = source.InternalDefaultValue
End Sub

Friend Function NextAutoIncrementValue() As ROneCOne
    Dim currentValue As LongLong

    currentValue = mAutoIncrementCurrent
    mAutoIncrementCurrent = mAutoIncrementCurrent + mAutoIncrementStep
    Select Case mExpectedType
        Case vbByte
            Set NextAutoIncrementValue = CreateScalarResult(CByte(currentValue))
        Case vbInteger
            Set NextAutoIncrementValue = CreateScalarResult(CInt(currentValue))
        Case vbLong
            Set NextAutoIncrementValue = CreateScalarResult(CLng(currentValue))
        Case Else
            Set NextAutoIncrementValue = CreateScalarResult(CLngLng(currentValue))
    End Select
End Function

Friend Sub ConfigureDbConnection( _
    ByVal connectionString As String, _
    ByVal providerName As String _
)
    mRole = ROLE_DB_CONNECTION
    mConnectionString = connectionString
    mProviderName = UCase$(Trim$(providerName))
    If mProviderName <> "OLEDB" And mProviderName <> "ODBC" Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DbConnection", _
            "Provider name must be OLEDB or ODBC."
    End If
End Sub

Friend Sub ConfigureDbCommand( _
    ByVal commandText As String, _
    ByVal connection As ROneCOne _
)
    mRole = ROLE_DB_COMMAND
    mCommandText = commandText
    Set mSource = connection
    Set mItems = New Collection
    mCommandType = 1
End Sub

Friend Sub ConfigureDbParameter( _
    ByVal parameterName As String, _
    ByVal valueWrapped As ROneCOne, _
    ByVal direction As Long _
)
    mRole = ROLE_DB_PARAMETER
    mName = parameterName
    Set mQueryValue = valueWrapped
    mExpressionOption = direction
    mSourceVersion = "Current"
End Sub

Friend Sub ConfigureDbAdapter(ByVal selectCommand As ROneCOne)
    mRole = ROLE_DB_ADAPTER
    Set mSource = selectCommand
End Sub

Friend Sub ConfigureDbReader(ByVal providerReader As Object)
    mRole = ROLE_DB_READER
    Set mProviderObject = providerReader
    mQueryCount = 0
End Sub

Friend Sub ConfigureDbTransaction(ByVal connection As ROneCOne)
    mRole = ROLE_DB_TRANSACTION
    Set mSource = connection
    mTaskStatus = TASK_STATUS_RUNNING
    mTransactionActive = True
End Sub

Friend Sub ConfigureGrouping( _
    ByVal keyWrapped As ROneCOne, _
    ByVal firstElement As ROneCOne _
)
    ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_GROUPING, _
        vbVariant, vbNullString, vbVariant
    Set mLeft = keyWrapped
    mItems.Add firstElement
End Sub

Friend Sub ConfigureEmptyGrouping(ByVal keyWrapped As ROneCOne)
    ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_GROUPING, _
        vbVariant, vbNullString, vbVariant
    Set mLeft = keyWrapped
End Sub

Friend Sub InternalAddLookupGroup( _
    ByVal keyWrapped As ROneCOne, _
    ByVal group As ROneCOne _
)
    mKeys.Add keyWrapped
    mItems.Add CreateObjectResult(group)
End Sub

Friend Function InternalCollectionItem(ByVal index As Long) As ROneCOne
    Dim wrapped As ROneCOne

    Set wrapped = mItems.Item(index + 1)
    If wrapped.InternalHasObjectValue Then
        Set InternalCollectionItem = wrapped.InternalObjectValue
    Else
        Set InternalCollectionItem = wrapped
    End If
End Function

Friend Sub InternalAddTaskChild(ByVal taskValue As ROneCOne)
    mTaskChildren.Add taskValue
End Sub

Friend Sub InternalAppendCollectionItem(ByVal wrapped As ROneCOne)
    mItems.Add wrapped
End Sub

Friend Sub InternalAppendCollectionKey(ByVal wrapped As ROneCOne)
    mKeys.Add wrapped
End Sub

Friend Sub InternalAppendCollectionPriority(ByVal wrapped As ROneCOne)
    mPriorities.Add wrapped
End Sub

Friend Sub InternalQueueValue(ByVal wrapped As ROneCOne)
    mItems.Add wrapped
    mVersion = mVersion + 1
End Sub

Friend Sub InternalPushValue(ByVal wrapped As ROneCOne)
    If mItems.Count = 0 Then
        mItems.Add wrapped
    Else
        mItems.Add wrapped, Before:=1
    End If
    mVersion = mVersion + 1
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
    ElseIf IsGenericCollectionRole(mRole) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            InternalSequenceVersion = mSource.InternalSequenceVersion
        Else
            InternalSequenceVersion = mVersion
        End If
    ElseIf mRole = ROLE_DATA_VIEW Then
        InternalSequenceVersion = mSource.InternalSequenceVersion
    ElseIf mRole = ROLE_DATA_TABLE Then
        InternalSequenceVersion = mVersion
    End If
End Property

Friend Function MaterializeWrapped() As Collection
    Dim entryIndex As Long
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
    If mRole = ROLE_DATA_VIEW Then
        Set MaterializeWrapped = MaterializeDataView
        Exit Function
    End If
    If IsGenericCollectionRole(mRole) Then
        If mCollectionKind = COLLECTION_READ_ONLY Then
            Set MaterializeWrapped = mSource.MaterializeWrapped
            Exit Function
        End If
        Set values = New Collection
        If IsDictionaryCollection Then
            For entryIndex = 1 To mItems.Count
                values.Add CreateDictionaryEntry( _
                    mKeys.Item(entryIndex), mItems.Item(entryIndex))
            Next entryIndex
        Else
            For Each wrapped In mItems
                values.Add wrapped
            Next wrapped
        End If
        Set MaterializeWrapped = values
        Exit Function
    End If

    If IsOrderingOperator(mOperator) Then
        Set MaterializeWrapped = ApplyOrderChain
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
        Case QUERY_TAKE_WHILE
            Set values = ApplyTakeWhile(sourceValues)
        Case QUERY_SKIP_WHILE
            Set values = ApplySkipWhile(sourceValues)
        Case QUERY_TAKE_LAST
            Set values = ApplyTakeLast(sourceValues)
        Case QUERY_SKIP_LAST
            Set values = ApplySkipLast(sourceValues)
        Case QUERY_DISTINCT
            Set values = ApplyDistinct(sourceValues)
        Case QUERY_DISTINCT_BY
            Set values = ApplyDistinctBy(sourceValues)
        Case QUERY_APPEND
            Set values = CopyWrappedCollection(sourceValues)
            values.Add mQueryValue
        Case QUERY_PREPEND
            Set values = New Collection
            values.Add mQueryValue
            AddWrappedRange values, sourceValues
        Case QUERY_CONCAT
            Set values = CopyWrappedCollection(sourceValues)
            AddWrappedRange values, mQueryValue.MaterializeWrapped
        Case QUERY_UNION
            Set values = ApplyUnion(sourceValues)
        Case QUERY_INTERSECT
            Set values = ApplyIntersect(sourceValues)
        Case QUERY_EXCEPT
            Set values = ApplyExcept(sourceValues)
        Case QUERY_DEFAULT_IF_EMPTY
            Set values = ApplyDefaultIfEmpty(sourceValues)
        Case QUERY_CHUNK
            Set values = ApplyChunk(sourceValues)
        Case QUERY_SELECT_MANY
            Set values = ApplySelectMany(sourceValues)
        Case QUERY_CAST
            Set values = ApplyCast(sourceValues, False)
        Case QUERY_OF_TYPE
            Set values = ApplyCast(sourceValues, True)
        Case QUERY_UNION_BY, QUERY_INTERSECT_BY, QUERY_EXCEPT_BY
            Set values = ApplySetBy(sourceValues)
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
            resolvedMember = Trim$(memberName)
            If Len(resolvedMember) > 0 Then
                Set created = New ROneCOne
                created.ConfigureMethod candidate, resolvedMember, isAction
                Set CreateCallableDelegate = created
                Exit Function
            End If
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

Private Sub RequireDbConnection( _
    ByVal connection As ROneCOne, _
    ByVal memberName As String _
)
    If connection Is Nothing Or connection.InternalRole <> ROLE_DB_CONNECTION Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbConnection."
    End If
End Sub

Private Sub RequireDbCommand( _
    ByVal command As ROneCOne, _
    ByVal memberName As String _
)
    If command Is Nothing Or command.InternalRole <> ROLE_DB_COMMAND Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbCommand."
    End If
End Sub

Private Sub RequireDbAdapter( _
    ByVal adapter As ROneCOne, _
    ByVal memberName As String _
)
    If adapter Is Nothing Or adapter.InternalRole <> ROLE_DB_ADAPTER Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DbDataAdapter."
    End If
End Sub

Private Function CreateProviderTask( _
    ByVal methodName As String, _
    Optional ByVal argument As Variant, _
    Optional ByVal cancellationToken As ROneCOne _
) As ROneCOne
    Dim arguments As Collection
    Dim taskValue As ROneCOne

    ValidateCancellationToken cancellationToken, methodName & "Async"
    Set arguments = New Collection
    If Not IsMissing(argument) Then
        arguments.Add WrapInvocationArgument(argument)
    End If
    Set taskValue = New ROneCOne
    taskValue.ConfigureProviderTask Me, methodName, arguments, cancellationToken
    Set CreateProviderTask = taskValue
End Function

Private Function CreateReadOnlyCollectionFromItems( _
    ByVal items As Collection _
) As ROneCOne
    Dim item As ROneCOne
    Dim values As ROneCOne

    Set values = ROneCOne.ListOf(Me)
    For Each item In items
        values.Add item
    Next item
    Set CreateReadOnlyCollectionFromItems = _
        ROneCOne.ReadOnlyCollectionOf(values)
End Function

Private Function RequireProviderTaskTarget( _
    ByVal arguments As Collection _
) As ROneCOne
    Dim candidate As Object
    Dim wrapped As ROneCOne

    If arguments Is Nothing Or arguments.Count <> 1 Then
        RaiseContractError ERROR_INVALID_OPERATION, "ProviderTask", _
            "The provider Task requires one data target."
    End If
    Set wrapped = arguments.Item(1)
    If Not wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProviderTask", _
            "The provider Task target must be a DataTable or DataSet."
    End If
    Set candidate = wrapped.InternalObjectValue
    If candidate Is Nothing Or Not TypeOf candidate Is ROneCOne Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProviderTask", _
            "The provider Task target must be a DataTable or DataSet."
    End If
    Set RequireProviderTaskTarget = candidate
End Function

Private Function BuildAdoCommand( _
    Optional ByVal dataRow As ROneCOne _
) As Object
    Dim commandObject As Object
    Dim connectionObject As Object
    Dim parameter As ROneCOne
    Dim parametersObject As Object
    Dim providerParameter As Object
    Dim rawValue As Variant
    Dim valueSize As Long
    Dim valueWrapped As ROneCOne

    RequireDbCommand Me, "Execute"
    mSource.Connect
    Set connectionObject = mSource.InternalProviderObject
    Set commandObject = CreateObject("ADODB.Command")
    CallByName commandObject, "ActiveConnection", VbSet, connectionObject
    CallByName commandObject, "CommandText", VbLet, mCommandText
    If mCommandTimeout > 0 Then
        CallByName commandObject, "CommandTimeout", VbLet, mCommandTimeout
    End If
    CallByName commandObject, "CommandType", VbLet, mCommandType
    Set parametersObject = CallByName(commandObject, "Parameters", VbGet)
    For Each parameter In mItems
        If dataRow Is Nothing Or Len(parameter.InternalSourceColumn) = 0 Then
            Set valueWrapped = parameter.InternalQueryValue
        Else
            Set valueWrapped = dataRow.DataRowVersionValue( _
                parameter.InternalSourceColumn, parameter.InternalSourceVersion)
        End If
        rawValue = Unwrap(valueWrapped)
        If VarType(rawValue) = vbString Then valueSize = Len(CStr(rawValue))
        If parameter.InternalProviderSize > 0 Then
            valueSize = parameter.InternalProviderSize
        ElseIf valueSize < 1 Then
            valueSize = 1
        End If
        Set providerParameter = CallByName( _
            commandObject, "CreateParameter", VbMethod, _
            parameter.InternalName, AdoTypeForParameter(parameter, valueWrapped), _
            parameter.InternalExpressionOption, valueSize, rawValue)
        If parameter.InternalProviderPrecision > 0 Then
            CallByName providerParameter, "Precision", VbLet, _
                parameter.InternalProviderPrecision
        End If
        If parameter.InternalProviderScale > 0 Then
            CallByName providerParameter, "NumericScale", VbLet, _
                parameter.InternalProviderScale
        End If
        CallByName parametersObject, "Append", VbMethod, providerParameter
    Next parameter
    Set BuildAdoCommand = commandObject
End Function

Private Function AdoTypeForParameter( _
    ByVal parameter As ROneCOne, _
    ByVal wrapped As ROneCOne _
) As Long
    If parameter.InternalProviderType <> 0 Then
        AdoTypeForParameter = parameter.InternalProviderType
    Else
        AdoTypeForParameter = AdoTypeForValue(wrapped)
    End If
End Function

Private Function NormalizeDataRowVersion( _
    ByVal sourceVersion As String, _
    ByVal memberName As String _
) As String
    Select Case UCase$(Trim$(sourceVersion))
        Case "CURRENT"
            NormalizeDataRowVersion = "Current"
        Case "ORIGINAL"
            NormalizeDataRowVersion = "Original"
        Case Else
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "Source version must be Current or Original."
    End Select
End Function

Friend Function ExecuteNonQueryForDataRow( _
    ByVal dataRow As ROneCOne _
) As Long
    Dim commandObject As Object
    Dim recordsAffected As Variant
    Dim recordset As Object

    RequireDbCommand Me, "Update"
    Set commandObject = BuildAdoCommand(dataRow)
    Set recordset = commandObject.Execute(recordsAffected)
    If Not recordset Is Nothing Then
        If CLng(CallByName(recordset, "State", VbGet)) <> 0 Then
            CallByName recordset, "Close", VbMethod
        End If
    End If
    ExecuteNonQueryForDataRow = CLng(recordsAffected)
End Function

Private Function UpdateDataTable(ByVal table As ROneCOne) As Long
    Dim command As ROneCOne
    Dim row As ROneCOne
    Dim rowState As String

    RequireDataTable table, "Update"
    For Each row In table.InternalDataRows
        rowState = row.RowState
        Set command = Nothing
        Select Case rowState
            Case "Added"
                Set command = mLeft
            Case "Modified"
                Set command = mRight
            Case "Deleted"
                Set command = mNext
        End Select
        If rowState <> "Unchanged" And rowState <> "Detached" Then
            If command Is Nothing Then
                RaiseContractError ERROR_INVALID_OPERATION, "Update", _
                    "The adapter has no command for " & rowState & " rows."
            End If
            command.ExecuteNonQueryForDataRow row
            UpdateDataTable = UpdateDataTable + 1
        End If
    Next row
    If UpdateDataTable > 0 Then table.AcceptChanges
End Function

Private Function AdoTypeForValue(ByVal wrapped As ROneCOne) As Long
    If wrapped.InternalHasObjectValue Then
        AdoTypeForValue = 12
        Exit Function
    End If
    Select Case VarType(wrapped.InternalScalarValue)
        Case vbByte
            AdoTypeForValue = 16
        Case vbInteger
            AdoTypeForValue = 2
        Case vbLong
            AdoTypeForValue = 3
        Case vbLongLong
            AdoTypeForValue = 20
        Case vbSingle
            AdoTypeForValue = 4
        Case vbDouble
            AdoTypeForValue = 5
        Case vbCurrency
            AdoTypeForValue = 6
        Case vbDate
            AdoTypeForValue = 7
        Case vbBoolean
            AdoTypeForValue = 11
        Case vbString
            AdoTypeForValue = 202
        Case Else
            AdoTypeForValue = 12
    End Select
End Function

Private Function ReadProviderField(ByVal index As Variant) As Variant
    Dim field As Object
    Dim fields As Object

    Set fields = CallByName(mProviderObject, "Fields", VbGet)
    Set field = CallByName(fields, "Item", VbGet, index)
    ReadProviderField = CallByName(field, "Value", VbGet)
End Function

Private Function FillTableFromReader( _
    ByVal table As ROneCOne, _
    ByVal reader As ROneCOne _
) As Long
    Dim columnIndex As Long
    Dim row As ROneCOne

    If table.Columns.Count = 0 Then
        For columnIndex = 0 To reader.FieldCount - 1
            table.AddColumn ROneCOne.DataColumn(reader.GetName(columnIndex), vbVariant)
        Next columnIndex
    End If
    Do While reader.Read
        Set row = table.NewRow
        For columnIndex = 0 To reader.FieldCount - 1
            row.Item(columnIndex) = reader.Item(columnIndex)
        Next columnIndex
        table.AddRow row
        FillTableFromReader = FillTableFromReader + 1
    Loop
End Function

Private Sub RequireDataTable( _
    ByVal table As ROneCOne, _
    ByVal memberName As String _
)
    If table Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataTable cannot be Nothing."
    End If
    If table.InternalRole <> ROLE_DATA_TABLE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataTable."
    End If
End Sub

Private Sub RequireDataColumn( _
    ByVal column As ROneCOne, _
    ByVal memberName As String _
)
    If column Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataColumn cannot be Nothing."
    End If
    If column.InternalRole <> ROLE_DATA_COLUMN Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataColumn."
    End If
End Sub

Private Sub RequireDataRow( _
    ByVal row As ROneCOne, _
    ByVal memberName As String _
)
    If row Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataRow cannot be Nothing."
    End If
    If row.InternalRole <> ROLE_DATA_ROW Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataRow."
    End If
End Sub

Private Sub RequireDataRelation( _
    ByVal relation As ROneCOne, _
    ByVal memberName As String _
)
    If relation Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The DataRelation cannot be Nothing."
    End If
    If relation.InternalRole <> ROLE_DATA_RELATION Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a DataRelation."
    End If
End Sub

Private Function FindNamedObjectIndex( _
    ByVal values As Collection, _
    ByVal objectName As String _
) As Long
    Dim index As Long
    Dim runtimeValue As ROneCOne

    FindNamedObjectIndex = -1
    For index = 1 To values.Count
        Set runtimeValue = values.Item(index)
        If StrComp(runtimeValue.InternalName, objectName, vbBinaryCompare) = 0 Then
            FindNamedObjectIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function CreateNamedDataCollection( _
    ByVal values As Collection _
) As ROneCOne
    Dim result As ROneCOne
    Dim runtimeValue As ROneCOne

    Set result = New ROneCOne
    result.ConfigureGenericCollection ROLE_COLLECTION, COLLECTION_KEYED, _
        vbObject, "ROneCOne", vbString
    For Each runtimeValue In values
        result.InternalAppendCollectionKey _
            CreateScalarResult(runtimeValue.InternalName)
        result.InternalAppendCollectionItem CreateObjectResult(runtimeValue)
    Next runtimeValue
    Set CreateNamedDataCollection = result
End Function

Private Function CreateObjectListSnapshot( _
    ByVal values As Collection _
) As ROneCOne
    Dim result As ROneCOne
    Dim runtimeValue As ROneCOne

    Set result = New ROneCOne
    result.ConfigureList vbObject, "ROneCOne"
    For Each runtimeValue In values
        result.InternalAddWrapped CreateObjectResult(runtimeValue)
    Next runtimeValue
    Set CreateObjectListSnapshot = result
End Function

Private Function MaterializeDataView() As Collection
    Dim query As ROneCOne
    Dim row As ROneCOne
    Dim rows As ROneCOne

    Set rows = ROneCOne.ListOf(Me)
    For Each row In mSource.InternalDataRows
        If row.RowState <> "Deleted" Then rows.Add row
    Next row
    Set query = rows
    If Not mQueryDelegate Is Nothing Then
        Set query = query.Where(mQueryDelegate)
    End If
    If Not mKeySelector Is Nothing Then
        If mExpressionOption = 0 Then
            Set query = query.OrderBy(mKeySelector)
        Else
            Set query = query.OrderByDescending(mKeySelector)
        End If
    End If
    Set MaterializeDataView = query.MaterializeWrapped
End Function

Private Function CloneDataTableSchema( _
    ByVal sourceTable As ROneCOne _
) As ROneCOne
    Dim clonedColumn As ROneCOne
    Dim column As ROneCOne
    Dim primaryColumns As ROneCOne
    Dim result As ROneCOne

    Set result = ROneCOne.DataTable(sourceTable.TableName)
    For Each column In sourceTable.InternalDataColumns
        Set clonedColumn = New ROneCOne
        clonedColumn.ConfigureDataColumn column.ColumnName, _
            column.InternalExpectedType, column.InternalExpectedObjectType
        clonedColumn.CopyDataColumnConfiguration column
        result.AddColumn clonedColumn
    Next column
    If sourceTable.InternalPrimaryKey.Count > 0 Then
        Set primaryColumns = ROneCOne.ListOf(vbObject)
        For Each column In sourceTable.InternalPrimaryKey
            primaryColumns.Add result.Columns(column.ColumnName)
        Next column
        Set result.PrimaryKey = primaryColumns
    End If
    Set CloneDataTableSchema = result
End Function

Private Sub CopyDataRowInto( _
    ByVal sourceRow As ROneCOne, _
    ByVal targetTable As ROneCOne _
)
    Dim index As Long
    Dim row As ROneCOne

    Set row = targetTable.NewRow
    For index = 0 To targetTable.Columns.Count - 1
        row.Item(index) = Unwrap(sourceRow.InternalDataRows.Item(index + 1))
    Next index
    targetTable.AddRow row
End Sub

Private Sub ValidateCompatibleDataTable( _
    ByVal sourceTable As ROneCOne, _
    ByVal memberName As String _
)
    Dim index As Long
    Dim sourceColumn As ROneCOne
    Dim targetColumn As ROneCOne

    If mKeys.Count <> sourceTable.InternalDataColumns.Count Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The table schemas have different column counts."
    End If
    For index = 1 To mKeys.Count
        Set targetColumn = mKeys.Item(index)
        Set sourceColumn = sourceTable.InternalDataColumns.Item(index)
        If targetColumn.ColumnName <> sourceColumn.ColumnName Or _
            targetColumn.InternalExpectedType <> sourceColumn.InternalExpectedType Or _
            targetColumn.InternalExpectedObjectType <> _
                sourceColumn.InternalExpectedObjectType Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "The table schemas are not compatible."
        End If
    Next index
End Sub

Friend Function DataFieldWrapped(ByVal columnName As String) As ROneCOne
    Dim columnIndex As Long

    RequireDataRow Me, "DataRow.Item"
    columnIndex = FindDataColumnIndex(mSource, columnName)
    If columnIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataRow.Item", _
            "The requested column was not found."
    End If
    Set DataFieldWrapped = mItems.Item(columnIndex + 1)
End Function

Private Function DataRowItem(ByVal index As Variant) As ROneCOne
    Dim columnIndex As Long

    If VarType(index) = vbString Then
        Set DataRowItem = DataFieldWrapped(CStr(index))
        Exit Function
    End If
    columnIndex = CLng(index)
    ValidateIndex columnIndex, mItems.Count, "DataRow.Item"
    Set DataRowItem = mItems.Item(columnIndex + 1)
End Function

Friend Function DataRowVersionValue( _
    ByVal columnName As String, _
    ByVal sourceVersion As String _
) As ROneCOne
    Dim columnIndex As Long

    RequireDataRow Me, "DataRowVersionValue"
    columnIndex = FindNamedObjectIndex(mSource.InternalDataColumns, columnName)
    If columnIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "DataRowVersionValue", _
            "The requested source column was not found."
    End If
    If NormalizeDataRowVersion(sourceVersion, "DataRowVersionValue") = _
        "Original" And mOriginalItems.Count > 0 Then
        Set DataRowVersionValue = mOriginalItems.Item(columnIndex + 1)
    Else
        Set DataRowVersionValue = mItems.Item(columnIndex + 1)
    End If
End Function

Private Sub SetDataRowItem(ByVal index As Variant, ByVal value As Variant)
    Dim column As ROneCOne
    Dim columnIndex As Long
    Dim previous As ROneCOne
    Dim wrapped As ROneCOne

    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    RequireDataRow Me, "DataRow.Item"
    If VarType(index) = vbString Then
        columnIndex = FindDataColumnIndex(mSource, CStr(index))
    Else
        columnIndex = CLng(index)
    End If
    ValidateIndex columnIndex, mItems.Count, "DataRow.Item"
    Set column = mSource.InternalDataColumns.Item(columnIndex + 1)
    Set wrapped = WrapDataValue(value)
    ValidateDataColumnValue column, wrapped, "DataRow.Item"
    Set previous = mItems.Item(columnIndex + 1)
    ReplaceDataValue columnIndex, wrapped
    On Error GoTo RestorePrevious
    If mRowState <> ROW_STATE_DETACHED Then
        mSource.ValidateDataRowConstraints Me, column
        mSource.ValidateForeignKeyConstraints Me
    End If
    If mRowState = ROW_STATE_UNCHANGED Then mRowState = ROW_STATE_MODIFIED
    mSource.TouchDataVersion
    On Error GoTo 0
    Exit Sub

RestorePrevious:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    On Error GoTo 0
    ReplaceDataValue columnIndex, previous
    Err.Raise capturedNumber, capturedSource, capturedDescription
End Sub

Private Function FindDataColumnIndex( _
    ByVal table As ROneCOne, _
    ByVal columnName As String _
) As Long
    FindDataColumnIndex = FindNamedObjectIndex( _
        table.InternalDataColumns, columnName)
End Function

Private Function WrapDataValue(ByVal value As Variant) As ROneCOne
    Dim valueObject As Object

    If IsObject(value) Then
        Set valueObject = value
        Set WrapDataValue = CreateObjectResult(valueObject)
    Else
        Set WrapDataValue = CreateScalarResult(value)
    End If
End Function

Private Function DefaultDataColumnValue( _
    ByVal column As ROneCOne _
) As ROneCOne
    Dim defaultValue As ROneCOne

    If column.InternalAutoIncrement Then
        Set DefaultDataColumnValue = column.NextAutoIncrementValue
        Exit Function
    End If
    Set defaultValue = column.InternalDefaultValue
    If Not defaultValue Is Nothing Then
        Set DefaultDataColumnValue = defaultValue
        Exit Function
    End If
    If column.InternalAllowDBNull Then
        Set DefaultDataColumnValue = CreateScalarResult(Null)
        Exit Function
    End If
    Select Case column.InternalExpectedType
        Case vbString
            Set DefaultDataColumnValue = CreateScalarResult(vbNullString)
        Case vbBoolean
            Set DefaultDataColumnValue = CreateScalarResult(False)
        Case vbDate
            Set DefaultDataColumnValue = CreateScalarResult(CDate(0))
        Case vbByte
            Set DefaultDataColumnValue = CreateScalarResult(CByte(0))
        Case vbInteger
            Set DefaultDataColumnValue = CreateScalarResult(CInt(0))
        Case vbLong
            Set DefaultDataColumnValue = CreateScalarResult(CLng(0))
        Case vbLongLong
            Set DefaultDataColumnValue = CreateScalarResult(CLngLng(0))
        Case vbSingle
            Set DefaultDataColumnValue = CreateScalarResult(CSng(0))
        Case vbDouble
            Set DefaultDataColumnValue = CreateScalarResult(CDbl(0))
        Case vbCurrency
            Set DefaultDataColumnValue = CreateScalarResult(CCur(0))
        Case vbObject
            Set DefaultDataColumnValue = CreateObjectResult(Nothing)
        Case Else
            Set DefaultDataColumnValue = CreateScalarResult(Empty)
    End Select
End Function

Private Sub ValidateDataColumnValue( _
    ByVal column As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    If Not wrapped.InternalHasObjectValue Then
        If IsNull(wrapped.InternalScalarValue) Then
            If column.InternalAllowDBNull Then Exit Sub
            RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
                "The DataColumn does not allow Null."
        End If
    End If
    ValidateWrappedType wrapped, column.InternalExpectedType, _
        column.InternalExpectedObjectType, memberName, "value"
End Sub

Private Sub ReplaceDataValue( _
    ByVal columnIndex As Long, _
    ByVal wrapped As ROneCOne _
)
    Dim current As ROneCOne
    Dim index As Long
    Dim replacement As Collection

    Set replacement = New Collection
    For index = 1 To mItems.Count
        If index = columnIndex + 1 Then
            replacement.Add wrapped
        Else
            Set current = mItems.Item(index)
            replacement.Add current
        End If
    Next index
    Set mItems = replacement
End Sub

Friend Sub ValidateDataRowConstraints( _
    ByVal row As ROneCOne, _
    ByVal changedColumn As ROneCOne _
)
    Dim column As ROneCOne
    Dim columnIndex As Long
    Dim existing As ROneCOne
    Dim wrapped As ROneCOne

    columnIndex = 0
    For Each column In mKeys
        Set wrapped = row.InternalDataRows.Item(columnIndex + 1)
        ValidateDataColumnValue column, wrapped, "AddRow"
        If column.InternalUnique Then
            For Each existing In mItems
                If Not existing Is row Then
                    If DataValuesEqual( _
                        existing.InternalDataRows.Item(columnIndex + 1), _
                        wrapped, "Unique") Then
                        RaiseContractError ERROR_INVALID_OPERATION, "AddRow", _
                            "A unique DataColumn contains a duplicate value."
                    End If
                End If
            Next existing
        End If
        columnIndex = columnIndex + 1
    Next column
End Sub

Friend Sub ValidateForeignKeyConstraints(ByVal row As ROneCOne)
    Dim dataSetValue As ROneCOne
    Dim relation As ROneCOne

    If mSource Is Nothing Then Exit Sub
    If mSource.InternalRole <> ROLE_DATA_SET Then Exit Sub
    Set dataSetValue = mSource
    If Not dataSetValue.InternalEnforceConstraints Then Exit Sub
    For Each relation In dataSetValue.InternalDataRelations
        If relation.InternalExpressionOption <> 0 Then
            If relation.InternalRelationChildColumn.InternalSource Is Me Then
                ValidateChildRowAgainstRelation row, relation
            End If
        End If
    Next relation
End Sub

Private Function IsNullDataValue(ByVal wrapped As ROneCOne) As Boolean
    If wrapped.InternalHasObjectValue Then
        IsNullDataValue = (wrapped.InternalObjectValue Is Nothing)
    Else
        IsNullDataValue = IsNull(wrapped.InternalScalarValue)
    End If
End Function

Private Function DataValuesEqual( _
    ByVal leftValue As ROneCOne, _
    ByVal rightValue As ROneCOne, _
    ByVal memberName As String _
) As Boolean
    DataValuesEqual = ValuesEqualWithComparer( _
        leftValue, rightValue, Nothing, memberName)
End Function

Private Sub AcceptRowChanges()
    Dim wrapped As ROneCOne

    Set mOriginalItems = New Collection
    For Each wrapped In mItems
        mOriginalItems.Add wrapped
    Next wrapped
    mRowState = ROW_STATE_UNCHANGED
End Sub

Private Sub RejectRowChanges()
    Dim wrapped As ROneCOne

    If mRowState = ROW_STATE_ADDED Then
        mRowState = ROW_STATE_DETACHED
        Exit Sub
    End If
    If mRowState = ROW_STATE_UNCHANGED Then Exit Sub
    Set mItems = New Collection
    For Each wrapped In mOriginalItems
        mItems.Add wrapped
    Next wrapped
    mRowState = ROW_STATE_UNCHANGED
End Sub

Private Sub ValidateRelationOwnership(ByVal relation As ROneCOne)
    Dim childTable As ROneCOne
    Dim childRow As ROneCOne
    Dim parentTable As ROneCOne

    Set parentTable = relation.InternalRelationParentColumn.InternalSource
    Set childTable = relation.InternalRelationChildColumn.InternalSource
    If parentTable Is Nothing Or childTable Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "Relation columns must belong to DataTables."
    End If
    If Not parentTable.InternalSource Is Me Or _
        Not childTable.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "Both relation tables must belong to this DataSet."
    End If
    If relation.InternalExpressionOption <> 0 Then
        ValidateExistingUniqueColumn relation.InternalRelationParentColumn
        For Each childRow In childTable.InternalDataRows
            ValidateChildRowAgainstRelation childRow, relation
        Next childRow
        relation.InternalRelationParentColumn.ConfigureUniqueColumn
    End If
End Sub

Private Sub ValidateExistingUniqueColumn(ByVal column As ROneCOne)
    Dim firstIndex As Long
    Dim firstRow As ROneCOne
    Dim firstValue As ROneCOne
    Dim rows As Collection
    Dim secondIndex As Long
    Dim secondRow As ROneCOne

    Set rows = column.InternalSource.InternalDataRows
    For firstIndex = 1 To rows.Count
        Set firstRow = rows.Item(firstIndex)
        Set firstValue = firstRow.DataFieldWrapped(column.ColumnName)
        For secondIndex = firstIndex + 1 To rows.Count
            Set secondRow = rows.Item(secondIndex)
            If DataValuesEqual(firstValue, _
                secondRow.DataFieldWrapped(column.ColumnName), _
                "DataRelation") Then
                RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
                    "The parent DataColumn contains duplicate values."
            End If
        Next secondIndex
    Next firstIndex
End Sub

Private Sub ValidateChildRowAgainstRelation( _
    ByVal childRow As ROneCOne, _
    ByVal relation As ROneCOne _
)
    Dim childColumn As ROneCOne
    Dim childValue As ROneCOne
    Dim foundParent As Boolean
    Dim parentColumn As ROneCOne
    Dim parentRow As ROneCOne

    Set childColumn = relation.InternalRelationChildColumn
    Set childValue = childRow.DataFieldWrapped(childColumn.ColumnName)
    If IsNullDataValue(childValue) Then Exit Sub
    Set parentColumn = relation.InternalRelationParentColumn
    For Each parentRow In parentColumn.InternalSource.InternalDataRows
        If parentRow.RowState <> "Deleted" Then
            If DataValuesEqual(childValue, _
                parentRow.DataFieldWrapped(parentColumn.ColumnName), _
                "ForeignKey") Then
                foundParent = True
                Exit For
            End If
        End If
    Next parentRow
    If Not foundParent Then
        RaiseContractError ERROR_INVALID_OPERATION, "AddRelation", _
            "A DataRelation foreign-key constraint has no parent row."
    End If
End Sub

Private Function NavigateRelation( _
    ByVal relationName As String, _
    ByVal findChildren As Boolean _
) As ROneCOne
    Dim candidate As ROneCOne
    Dim dataSetValue As ROneCOne
    Dim relation As ROneCOne
    Dim relationIndex As Long
    Dim result As ROneCOne
    Dim sourceColumn As ROneCOne
    Dim sourceValue As ROneCOne
    Dim targetColumn As ROneCOne
    Dim targetTable As ROneCOne

    Set dataSetValue = mSource.InternalSource
    If dataSetValue Is Nothing Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relation navigation", _
            "The DataTable does not belong to a DataSet."
    End If
    relationIndex = FindNamedObjectIndex( _
        dataSetValue.InternalDataRelations, relationName)
    If relationIndex < 0 Then
        RaiseContractError ERROR_INVALID_ARGUMENT, "Relation navigation", _
            "The requested DataRelation was not found."
    End If
    Set relation = dataSetValue.InternalDataRelations.Item(relationIndex + 1)
    If findChildren Then
        Set sourceColumn = relation.InternalRelationParentColumn
        Set targetColumn = relation.InternalRelationChildColumn
    Else
        Set sourceColumn = relation.InternalRelationChildColumn
        Set targetColumn = relation.InternalRelationParentColumn
    End If
    If Not sourceColumn.InternalSource Is mSource Then
        RaiseContractError ERROR_INVALID_OPERATION, "Relation navigation", _
            "The DataRow is on the wrong side of this relation."
    End If
    Set sourceValue = DataFieldWrapped(sourceColumn.ColumnName)
    Set targetTable = targetColumn.InternalSource
    Set result = ROneCOne.ListOf(Me)
    For Each candidate In targetTable.InternalDataRows
        If DataValuesEqual(sourceValue, _
            candidate.DataFieldWrapped(targetColumn.ColumnName), _
            "Relation navigation") Then result.Add candidate
    Next candidate
    Set NavigateRelation = result
End Function

Private Sub ValidateTaskWork( _
    ByVal work As ROneCOne, _
    ByVal memberName As String _
)
    If work Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The Task body cannot be Nothing."
    End If
    RequireDelegate work, memberName
    If work.Arity <> 0 And work.Arity <> -1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The Task body must accept no arguments."
    End If
End Sub

Private Sub ValidateProgressHandler( _
    ByVal handler As ROneCOne, _
    ByVal valueType As VbVarType, _
    ByVal objectType As String _
)
    Dim descriptor As ROneCOne

    ValidateAction handler, "ProgressOf", False
    If Not handler.InternalHasParameterSignature Or handler.Arity <> 1 Then
        RaiseContractError ERROR_ARITY_MISMATCH, "ProgressOf", _
            "The progress handler must declare exactly one parameter."
    End If
    Set descriptor = handler.InternalParameters.Item(1)
    If descriptor.InternalExpectedType <> valueType Or _
        StrComp(descriptor.InternalExpectedObjectType, objectType, _
            vbTextCompare) <> 0 Then
        RaiseContractError ERROR_TYPE_MISMATCH, "ProgressOf", _
            "The progress handler parameter does not match Progress<T>."
    End If
End Sub

Private Sub RequireTaskCompletionSource(ByVal memberName As String)
    If mRole = ROLE_TASK_COMPLETION_SOURCE Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a TaskCompletionSource<T>."
End Sub

Private Function WrapTypedCompletionValue(ByVal value As Variant) As ROneCOne
    Dim wrapped As ROneCOne

    Set wrapped = WrapInvocationArgument(value)
    ValidateWrappedType wrapped, mExpectedType, mExpectedObjectType, _
        "TaskCompletionSource", "result"
    Set WrapTypedCompletionValue = wrapped
End Function

Private Sub ValidateCancellationToken( _
    ByVal tokenValue As ROneCOne, _
    ByVal memberName As String _
)
    If tokenValue Is Nothing Then Exit Sub
    If tokenValue.InternalRole <> ROLE_CANCELLATION_TOKEN Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The cancellation value must be a CancellationToken."
    End If
End Sub

Private Sub RequireTask(ByVal taskValue As ROneCOne, ByVal memberName As String)
    If taskValue Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The Task cannot be Nothing."
    End If
    If taskValue.InternalRole <> ROLE_TASK Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a Task."
    End If
End Sub

Private Function CreateTaskCombinator( _
    ByVal tasks As Variant, _
    ByVal taskKind As Long, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidate As Object
    Dim child As ROneCOne
    Dim rawTask As Variant
    Dim sequenceValue As ROneCOne
    Dim taskValue As ROneCOne

    Set taskValue = New ROneCOne
    taskValue.ConfigureTask taskKind, Nothing, Nothing
    If IsObject(tasks) Then
        Set candidate = tasks
        If Not candidate Is Nothing Then
            If TypeOf candidate Is ROneCOne Then
                Set sequenceValue = candidate
                RequireSequence sequenceValue, memberName
                For Each rawTask In sequenceValue
                    Set child = RequireROneCOne(rawTask, memberName)
                    RequireTask child, memberName
                    taskValue.InternalAddTaskChild child
                Next rawTask
                Set CreateTaskCombinator = taskValue
                Exit Function
            End If
        End If
    End If
    If Not IsArray(tasks) Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "Tasks must be supplied as a typed sequence or array."
    End If
    For Each rawTask In tasks
        Set child = RequireROneCOne(rawTask, memberName)
        RequireTask child, memberName
        taskValue.InternalAddTaskChild child
    Next rawTask
    Set CreateTaskCombinator = taskValue
End Function

Friend Sub EnsureTaskCompleted()
    Call WaitForTaskCompletion(-1)
End Sub

Friend Sub AdvanceTask()
Attribute AdvanceTask.VB_MemberFlags = "40"
    Dim arguments As Collection
    Dim allCompleted As Boolean
    Dim capturedDescription As String
    Dim capturedHelpContext As Long
    Dim capturedHelpFile As String
    Dim capturedNumber As Long
    Dim capturedSource As String
    Dim child As ROneCOne
    Dim dataTarget As ROneCOne
    Dim results As ROneCOne
    Dim sawCanceled As Boolean

    If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Sub
    If TaskCancellationRequested Then
        mTaskStatus = TASK_STATUS_CANCELED
        Exit Sub
    End If
    If mOperator = TASK_PROMISE Then Exit Sub
    mTaskStatus = TASK_STATUS_RUNNING
    On Error GoTo TaskFailed
    Select Case mOperator
        Case TASK_WORK
            Set arguments = New Collection
            Set mTaskResult = mTaskBody.InvokeValue(arguments)
        Case TASK_DELAY
            AdvanceDelayTask
            Exit Sub
        Case TASK_WHEN_ALL
            allCompleted = True
            For Each child In mTaskChildren
                child.AdvanceTask
                If Not child.IsCompleted Then allCompleted = False
            Next child
            If Not allCompleted Then Exit Sub
            Set results = ROneCOne.ListOf(vbVariant)
            For Each child In mTaskChildren
                If child.IsFaulted Then
                    If mTaskException Is Nothing Then
                        Set mTaskException = child.Exception
                    End If
                End If
                If child.IsCanceled Then sawCanceled = True
                If Not child.IsFaulted And Not child.IsCanceled Then
                    results.Add child.Result
                End If
            Next child
            If Not mTaskException Is Nothing Then
                mTaskStatus = TASK_STATUS_FAULTED
                Exit Sub
            End If
            If sawCanceled Then
                mTaskStatus = TASK_STATUS_CANCELED
                Exit Sub
            End If
            Set mTaskResult = CreateObjectResult(results)
        Case TASK_WHEN_ANY
            If mTaskChildren.Count = 0 Then
                RaiseContractError ERROR_INVALID_ARGUMENT, "WhenAny", _
                    "WhenAny requires at least one Task."
            End If
            For Each child In mTaskChildren
                child.AdvanceTask
                If child.IsCompleted Then
                    Set mTaskResult = CreateObjectResult(child)
                    mTaskStatus = TASK_STATUS_COMPLETED
                    Exit Sub
                End If
            Next child
            Exit Sub
        Case TASK_CONTINUATION
            mSource.AdvanceTask
            If Not mSource.IsCompleted Then Exit Sub
            Set arguments = New Collection
            arguments.Add CreateObjectResult(mSource)
            Set mTaskResult = mTaskBody.InvokeValue(arguments)
        Case TASK_PROVIDER_CALL
            Select Case mMethodName
                Case "Connect"
                    mSource.Connect
                    Set mTaskResult = CreateScalarResult(Empty)
                Case "ExecuteReader"
                    Set mTaskResult = CreateObjectResult(mSource.ExecuteReader)
                Case "ExecuteNonQuery"
                    Set mTaskResult = CreateScalarResult(mSource.ExecuteNonQuery)
                Case "ExecuteScalar"
                    Set mTaskResult = WrapDataValue(mSource.ExecuteScalar)
                Case "Read"
                    Set mTaskResult = CreateScalarResult(mSource.Read)
                Case "Fill", "Update"
                    Set dataTarget = RequireProviderTaskTarget(mTaskChildren)
                    If mMethodName = "Fill" Then
                        Set mTaskResult = CreateScalarResult(mSource.Fill(dataTarget))
                    Else
                        Set mTaskResult = CreateScalarResult(mSource.Update(dataTarget))
                    End If
                Case Else
                    RaiseContractError ERROR_INVALID_OPERATION, "ProviderTask", _
                        "Unknown provider Task operation."
            End Select
    End Select
    mTaskStatus = TASK_STATUS_COMPLETED
    On Error GoTo 0
    Exit Sub

TaskFailed:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    capturedHelpFile = Err.HelpFile
    capturedHelpContext = Err.HelpContext
    On Error GoTo 0
    Set mTaskException = New ROneCOne
    mTaskException.ConfigureError capturedNumber, capturedSource, _
        capturedDescription, capturedHelpFile, capturedHelpContext
    mTaskStatus = TASK_STATUS_FAULTED
End Sub

Private Function WaitForTaskCompletion( _
    ByVal timeoutMilliseconds As Long _
) As Boolean
    Dim elapsed As Double
    Dim startedAt As Double

    startedAt = Timer
    Do While mTaskStatus < TASK_STATUS_COMPLETED
        AdvanceTask
        If mTaskStatus >= TASK_STATUS_COMPLETED Then Exit Do
        DoEvents
        If timeoutMilliseconds >= 0 Then
            elapsed = ElapsedMilliseconds(startedAt)
            If elapsed >= CDbl(timeoutMilliseconds) Then Exit Function
        End If
    Loop
    WaitForTaskCompletion = True
End Function

Private Sub AdvanceDelayTask()
    If mExpressionOption = 0 Then
        mScalarValue = Timer
        mExpressionOption = 1
    End If
    If TaskCancellationRequested Then
        mTaskStatus = TASK_STATUS_CANCELED
        Exit Sub
    End If
    If ElapsedMilliseconds(CDbl(mScalarValue)) >= CDbl(mQueryCount) Then
        Set mTaskResult = CreateScalarResult(Empty)
        mTaskStatus = TASK_STATUS_COMPLETED
    End If
End Sub

Private Function ElapsedMilliseconds(ByVal startedAt As Double) As Double
    Dim elapsed As Double

    elapsed = Timer - startedAt
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ElapsedMilliseconds = elapsed * 1000#
End Function

Private Function TaskCancellationRequested() As Boolean
    If mCancellationSource Is Nothing Then Exit Function
    TaskCancellationRequested = mCancellationSource.IsCancellationRequested
End Function

Private Sub EvaluateScheduledCancellation()
    If mCancellationRequested Or mQueryCount < 0 Then Exit Sub
    If ElapsedMilliseconds(CDbl(mScalarValue)) >= CDbl(mQueryCount) Then
        CancelCancellationSource
    End If
End Sub

Private Sub CancelCancellationSource()
    Dim arguments As Collection
    Dim callback As ROneCOne
    Dim callbacks As Collection
    Dim ignored As ROneCOne

    If mCancellationRequested Then Exit Sub
    mCancellationRequested = True
    Set callbacks = CopyWrappedCollection(mInvocationList)
    Set arguments = New Collection
    For Each callback In callbacks
        Set ignored = callback.InvokeValue(arguments)
    Next callback
End Sub

Private Sub RaiseTaskTerminalState(ByVal memberName As String)
    If mTaskStatus = TASK_STATUS_CANCELED Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The Task was canceled."
    End If
    If mTaskStatus = TASK_STATUS_FAULTED Then
        RaiseCapturedError mTaskException
    End If
End Sub

Private Function IsGenericCollectionRole(ByVal roleValue As Long) As Boolean
    Select Case roleValue
        Case ROLE_DICTIONARY, ROLE_HASH_SET, ROLE_COLLECTION
            IsGenericCollectionRole = True
        Case ROLE_IMMUTABLE, ROLE_CONCURRENT
            IsGenericCollectionRole = True
    End Select
End Function

Friend Function IsDictionaryCollection() As Boolean
    Select Case mCollectionKind
        Case COLLECTION_DICTIONARY, COLLECTION_SORTED_LIST
            IsDictionaryCollection = True
        Case COLLECTION_SORTED_DICTIONARY, COLLECTION_ORDERED_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_CONCURRENT_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_IMMUTABLE_DICTIONARY
            IsDictionaryCollection = True
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            IsDictionaryCollection = True
    End Select
End Function

Private Function IsSetCollection() As Boolean
    Select Case mCollectionKind
        Case COLLECTION_HASH_SET, COLLECTION_SORTED_SET
            IsSetCollection = True
        Case COLLECTION_IMMUTABLE_HASH_SET, COLLECTION_IMMUTABLE_SORTED_SET
            IsSetCollection = True
    End Select
End Function

Private Function IsQueueKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_QUEUE, COLLECTION_CONCURRENT_QUEUE
            IsQueueKind = True
        Case COLLECTION_BLOCKING, COLLECTION_IMMUTABLE_QUEUE
            IsQueueKind = True
    End Select
End Function

Private Function IsStackKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_STACK, COLLECTION_CONCURRENT_STACK
            IsStackKind = True
        Case COLLECTION_IMMUTABLE_STACK
            IsStackKind = True
    End Select
End Function

Private Function IsPriorityQueueKind(ByVal collectionKind As Long) As Boolean
    IsPriorityQueueKind = (collectionKind = COLLECTION_PRIORITY_QUEUE)
End Function

Private Function MutableRoleForKind(ByVal collectionKind As Long) As Long
    Select Case collectionKind
        Case COLLECTION_IMMUTABLE_DICTIONARY
            MutableRoleForKind = ROLE_DICTIONARY
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            MutableRoleForKind = ROLE_DICTIONARY
        Case COLLECTION_IMMUTABLE_HASH_SET, COLLECTION_IMMUTABLE_SORTED_SET
            MutableRoleForKind = ROLE_HASH_SET
        Case Else
            MutableRoleForKind = ROLE_COLLECTION
    End Select
End Function

Private Sub RequireMutableCollection(ByVal memberName As String)
    If Not IsGenericCollectionRole(mRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This operation requires a generic collection."
    End If
    If mRole = ROLE_IMMUTABLE Or mCollectionKind = COLLECTION_READ_ONLY Or _
        mCollectionKind = COLLECTION_LOOKUP Or _
        mCollectionKind = COLLECTION_GROUPING Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "This collection is read-only."
    End If
End Sub

Private Sub RequireKeyedCollection(ByVal memberName As String)
    If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Or _
        mCollectionKind = COLLECTION_LOOKUP Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a keyed collection."
End Sub

Private Sub RequireSetCollection(ByVal memberName As String)
    If IsSetCollection Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a set."
End Sub

Private Sub RequireConcurrentDictionary(ByVal memberName As String)
    If mRole = ROLE_CONCURRENT And _
        mCollectionKind = COLLECTION_CONCURRENT_DICTIONARY Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a ConcurrentDictionary."
End Sub

Private Sub RequireBlockingCollection(ByVal memberName As String)
    If mRole = ROLE_CONCURRENT And _
        mCollectionKind = COLLECTION_BLOCKING Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a BlockingCollection."
End Sub

Private Sub RequirePriorityQueue(ByVal memberName As String)
    If IsPriorityQueueKind(mCollectionKind) Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a PriorityQueue."
End Sub

Private Sub RequireImmutableCollection(ByVal memberName As String)
    If mRole = ROLE_IMMUTABLE Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires an immutable collection."
End Sub

Private Sub RequireCollectionNode(ByVal memberName As String)
    If mRole = ROLE_NODE And Not mSource Is Nothing Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires an attached LinkedListNode."
End Sub

Private Sub RequireOwnedNode( _
    ByVal node As ROneCOne, _
    ByVal memberName As String _
)
    If node Is Nothing Or node.InternalRole <> ROLE_NODE Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The node cannot be Nothing and must be a LinkedListNode."
    End If
    If Not node.InternalSource Is Me Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The node belongs to a different linked list."
    End If
End Sub

Private Sub RequireLinkedList(ByVal memberName As String)
    If mCollectionKind = COLLECTION_LINKED_LIST Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a LinkedList<T>."
End Sub

Private Sub RequireSortedSet(ByVal memberName As String)
    If mCollectionKind = COLLECTION_SORTED_SET Or _
        mCollectionKind = COLLECTION_IMMUTABLE_SORTED_SET Then Exit Sub
    RaiseContractError ERROR_INVALID_OPERATION, memberName, _
        "This operation requires a sorted set."
End Sub

Private Sub RequireReference( _
    ByVal referenceValue As ROneCOne, _
    ByVal memberName As String _
)
    If referenceValue Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
            "The output reference cannot be Nothing."
    End If
    If referenceValue.InternalRole <> ROLE_REFERENCE Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The output must be created by a ROneCOne.Ref method."
    End If
End Sub

Friend Function WrapKey( _
    ByVal value As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim valueObject As Object
    Dim wrapped As ROneCOne

    If IsObject(value) Then
        Set valueObject = value
        Set wrapped = CreateObjectResult(valueObject)
    Else
        If IsNull(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A collection key cannot be Null."
        End If
        Set wrapped = CreateScalarResult(value)
    End If
    ValidateWrappedType wrapped, mKeyType, mKeyObjectType, memberName, "key"
    Set WrapKey = wrapped
End Function

Private Sub ValidateWrappedType( _
    ByVal wrapped As ROneCOne, _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String, _
    ByVal memberName As String, _
    ByVal valueName As String _
)
    Dim actualObject As Object

    If expectedType = vbVariant Then Exit Sub
    If expectedType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then GoTo TypeFailed
        Set actualObject = wrapped.InternalObjectValue
        If actualObject Is Nothing Then Exit Sub
        If TypeName(actualObject) = expectedObjectType Then Exit Sub
        GoTo TypeFailed
    End If
    If wrapped.InternalHasObjectValue Then GoTo TypeFailed
    If VarType(wrapped.InternalScalarValue) = expectedType Then Exit Sub
TypeFailed:
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "The collection " & valueName & " does not match its generic type."
End Sub

Private Function BuildCollectionTypeName() As String
    Dim keyName As String
    Dim valueName As String

    If Len(mElementObjectType) > 0 Then
        valueName = mElementObjectType
    Else
        valueName = TypeNameForVarType(mElementType)
    End If
    If Len(mKeyObjectType) > 0 Then
        keyName = mKeyObjectType
    Else
        keyName = TypeNameForVarType(mKeyType)
    End If
    Select Case mCollectionKind
        Case COLLECTION_DICTIONARY
            BuildCollectionTypeName = "Dictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_HASH_SET
            BuildCollectionTypeName = "HashSet<" & valueName & ">"
        Case COLLECTION_QUEUE
            BuildCollectionTypeName = "Queue<" & valueName & ">"
        Case COLLECTION_STACK
            BuildCollectionTypeName = "Stack<" & valueName & ">"
        Case COLLECTION_LINKED_LIST
            BuildCollectionTypeName = "LinkedList<" & valueName & ">"
        Case COLLECTION_SORTED_LIST
            BuildCollectionTypeName = "SortedList<" & keyName & ", " & valueName & ">"
        Case COLLECTION_SORTED_DICTIONARY
            BuildCollectionTypeName = _
                "SortedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_SORTED_SET
            BuildCollectionTypeName = "SortedSet<" & valueName & ">"
        Case COLLECTION_ORDERED_DICTIONARY
            BuildCollectionTypeName = _
                "OrderedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_PRIORITY_QUEUE
            BuildCollectionTypeName = _
                "PriorityQueue<" & valueName & ", " & keyName & ">"
        Case COLLECTION_OBSERVABLE
            BuildCollectionTypeName = "ObservableCollection<" & valueName & ">"
        Case COLLECTION_READ_ONLY
            BuildCollectionTypeName = "ReadOnlyCollection<" & valueName & ">"
        Case COLLECTION_KEYED
            BuildCollectionTypeName = _
                "KeyedCollection<" & keyName & ", " & valueName & ">"
        Case COLLECTION_LOOKUP
            BuildCollectionTypeName = _
                "Lookup<" & keyName & ", " & valueName & ">"
        Case COLLECTION_GROUPING
            BuildCollectionTypeName = _
                "IGrouping<" & keyName & ", " & valueName & ">"
        Case COLLECTION_CONCURRENT_DICTIONARY
            BuildCollectionTypeName = _
                "ConcurrentDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_CONCURRENT_QUEUE
            BuildCollectionTypeName = "ConcurrentQueue<" & valueName & ">"
        Case COLLECTION_CONCURRENT_STACK
            BuildCollectionTypeName = "ConcurrentStack<" & valueName & ">"
        Case COLLECTION_CONCURRENT_BAG
            BuildCollectionTypeName = "ConcurrentBag<" & valueName & ">"
        Case COLLECTION_BLOCKING
            BuildCollectionTypeName = "BlockingCollection<" & valueName & ">"
        Case Else
            BuildCollectionTypeName = BuildImmutableTypeName(keyName, valueName)
    End Select
End Function

Private Function BuildImmutableTypeName( _
    ByVal keyName As String, _
    ByVal valueName As String _
) As String
    Select Case mCollectionKind
        Case COLLECTION_IMMUTABLE_ARRAY
            BuildImmutableTypeName = "ImmutableArray<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_LIST
            BuildImmutableTypeName = "ImmutableList<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_DICTIONARY
            BuildImmutableTypeName = _
                "ImmutableDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_IMMUTABLE_HASH_SET
            BuildImmutableTypeName = "ImmutableHashSet<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_QUEUE
            BuildImmutableTypeName = "ImmutableQueue<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_STACK
            BuildImmutableTypeName = "ImmutableStack<" & valueName & ">"
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            BuildImmutableTypeName = _
                "ImmutableSortedDictionary<" & keyName & ", " & valueName & ">"
        Case COLLECTION_IMMUTABLE_SORTED_SET
            BuildImmutableTypeName = "ImmutableSortedSet<" & valueName & ">"
    End Select
End Function

Friend Sub AddMutable( _
    ByVal keyOrValue As Variant, _
    Optional ByVal value As Variant, _
    Optional ByVal memberName As String = "Add" _
)
    Dim keyWrapped As ROneCOne
    Dim valueWrapped As ROneCOne

    RequireMutableCollection memberName
    If IsDictionaryCollection Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A dictionary value is required."
        End If
        Set keyWrapped = WrapKey(keyOrValue, memberName)
        If FindKeyIndex(keyWrapped, memberName) >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "An item with the same key already exists."
        End If
        AddDictionaryPair keyWrapped, WrapElement(value, memberName), memberName
        Exit Sub
    End If
    If IsSetCollection Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A set accepts one value."
        End If
        Set valueWrapped = WrapElement(keyOrValue, memberName)
        If FindElementIndex(valueWrapped, memberName) < 0 Then
            AddSetValue valueWrapped, memberName
        End If
        Exit Sub
    End If
    If mCollectionKind = COLLECTION_KEYED Then
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A keyed collection accepts one item."
        End If
        Set valueWrapped = WrapElement(keyOrValue, memberName)
        Set keyWrapped = InvokeUnary(mKeySelector, valueWrapped)
        ValidateWrappedType keyWrapped, mKeyType, mKeyObjectType, _
            memberName, "key"
        If FindKeyIndex(keyWrapped, memberName) >= 0 Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "An item with the same key already exists."
        End If
        mKeys.Add keyWrapped
        mItems.Add valueWrapped
    ElseIf IsPriorityQueueKind(mCollectionKind) Then
        If IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "A priority value is required."
        End If
        AddPriorityValue WrapElement(keyOrValue, memberName), _
            WrapKey(value, memberName)
        Exit Sub
    ElseIf IsStackKind(mCollectionKind) Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "Use Push for a stack."
    Else
        If Not IsMissing(value) Then
            RaiseContractError ERROR_INVALID_ARGUMENT, memberName, _
                "This collection accepts one item."
        End If
        If mCollectionKind = COLLECTION_BLOCKING And mAddingCompleted Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "Adding has already been completed."
        End If
        If mCollectionKind = COLLECTION_BLOCKING And mQueryCount >= 0 Then
            If mItems.Count >= mQueryCount Then
                RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                    "The bounded collection is full."
            End If
        End If
        mItems.Add WrapElement(keyOrValue, memberName)
    End If
    mVersion = mVersion + 1
    NotifyCollectionChanged "Add"
End Sub

Friend Sub AddDictionaryPair( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim insertAt As Long

    If IsSortedMapKind(mCollectionKind) Then
        insertAt = FindSortedInsertIndex(mKeys, keyWrapped, memberName)
        If insertAt <= mKeys.Count Then
            mKeys.Add keyWrapped, Before:=insertAt
            mItems.Add valueWrapped, Before:=insertAt
        Else
            mKeys.Add keyWrapped
            mItems.Add valueWrapped
        End If
    Else
        mKeys.Add keyWrapped
        mItems.Add valueWrapped
    End If
    mVersion = mVersion + 1
End Sub

Friend Sub AddSetValue( _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim insertAt As Long

    If IsSortedSetKind(mCollectionKind) Then
        insertAt = FindSortedInsertIndex(mItems, wrapped, memberName)
        If insertAt <= mItems.Count Then
            mItems.Add wrapped, Before:=insertAt
        Else
            mItems.Add wrapped
        End If
    Else
        mItems.Add wrapped
    End If
    mVersion = mVersion + 1
End Sub

Friend Sub AddPriorityValue( _
    ByVal wrapped As ROneCOne, _
    ByVal priorityWrapped As ROneCOne _
)
    Dim insertAt As Long

    insertAt = FindSortedInsertIndex(mPriorities, priorityWrapped, "Enqueue")
    If insertAt <= mPriorities.Count Then
        mPriorities.Add priorityWrapped, Before:=insertAt
        mItems.Add wrapped, Before:=insertAt
    Else
        mPriorities.Add priorityWrapped
        mItems.Add wrapped
    End If
    mVersion = mVersion + 1
End Sub

Private Function IsSortedMapKind(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_SORTED_LIST, COLLECTION_SORTED_DICTIONARY
            IsSortedMapKind = True
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
            IsSortedMapKind = True
    End Select
End Function

Private Function IsSortedSetKind(ByVal collectionKind As Long) As Boolean
    IsSortedSetKind = (collectionKind = COLLECTION_SORTED_SET Or _
        collectionKind = COLLECTION_IMMUTABLE_SORTED_SET)
End Function

Private Function FindSortedInsertIndex( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    For index = 1 To values.Count
        If CompareWithComparer(candidate, values.Item(index), _
            mQueryComparer, memberName) < 0 Then
            FindSortedInsertIndex = index
            Exit Function
        End If
    Next index
    FindSortedInsertIndex = values.Count + 1
End Function

Friend Function FindKeyIndex( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    FindKeyIndex = -1
    For index = 1 To mKeys.Count
        If CollectionUsesEquality(mCollectionKind) Then
            If ValuesEqualWithComparer(mKeys.Item(index), candidate, _
                mQueryComparer, memberName) Then
                FindKeyIndex = index - 1
                Exit Function
            End If
        ElseIf CompareWithComparer(mKeys.Item(index), candidate, _
            mQueryComparer, memberName) = 0 Then
            FindKeyIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Friend Function FindElementIndex( _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    FindElementIndex = FindWrappedInCollection(mItems, candidate, memberName)
End Function

Private Function FindWrappedInCollection( _
    ByVal values As Collection, _
    ByVal candidate As ROneCOne, _
    ByVal memberName As String _
) As Long
    Dim index As Long

    FindWrappedInCollection = -1
    For index = 1 To values.Count
        If CollectionUsesEquality(mCollectionKind) Then
            If ValuesEqualWithComparer(values.Item(index), candidate, _
                mQueryComparer, memberName) Then
                FindWrappedInCollection = index - 1
                Exit Function
            End If
        ElseIf CompareWithComparer(values.Item(index), candidate, _
            mQueryComparer, memberName) = 0 Then
            FindWrappedInCollection = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function DistinctWrappedValues( _
    ByVal values As Collection, _
    ByVal memberName As String _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In values
        If FindWrappedInCollection(output, wrapped, memberName) < 0 Then
            output.Add wrapped
        End If
    Next wrapped
    Set DistinctWrappedValues = output
End Function

Private Function IsSetSubset( _
    ByVal other As Variant, _
    ByVal proper As Boolean, _
    ByVal memberName As String _
) As Boolean
    Dim incoming As Collection
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection memberName
    Set incoming = BuildIncomingItems(other, memberName)
    Set uniqueValues = DistinctWrappedValues(incoming, memberName)
    If proper And mItems.Count >= uniqueValues.Count Then Exit Function
    For Each wrapped In mItems
        If FindWrappedInCollection(uniqueValues, wrapped, memberName) < 0 Then
            Exit Function
        End If
    Next wrapped
    IsSetSubset = True
End Function

Private Function IsSetSuperset( _
    ByVal other As Variant, _
    ByVal proper As Boolean, _
    ByVal memberName As String _
) As Boolean
    Dim incoming As Collection
    Dim uniqueValues As Collection
    Dim wrapped As ROneCOne

    RequireSetCollection memberName
    Set incoming = BuildIncomingItems(other, memberName)
    Set uniqueValues = DistinctWrappedValues(incoming, memberName)
    If proper And mItems.Count <= uniqueValues.Count Then Exit Function
    For Each wrapped In uniqueValues
        If FindElementIndex(wrapped, memberName) < 0 Then Exit Function
    Next wrapped
    IsSetSuperset = True
End Function

Private Sub SetItemValue(ByVal index As Variant, ByVal value As Variant)
    Dim itemIndex As Long
    Dim keyWrapped As ROneCOne
    Dim replacement As ROneCOne

    If mRole = ROLE_DATA_ROW Then
        SetDataRowItem index, value
        Exit Sub
    End If
    If IsDictionaryCollection Then
        RequireMutableCollection "Item"
        Set keyWrapped = WrapKey(index, "Item")
        Set replacement = WrapElement(value, "Item")
        itemIndex = FindKeyIndex(keyWrapped, "Item")
        If itemIndex < 0 Then
            AddDictionaryPair keyWrapped, replacement, "Item"
        Else
            ReplaceCollectionItem itemIndex, replacement
        End If
        Exit Sub
    End If
    If IsGenericCollectionRole(mRole) Then
        RequireMutableCollection "Item"
        itemIndex = CLng(index)
        ValidateIndex itemIndex, mItems.Count, "Item"
        ReplaceCollectionItem itemIndex, WrapElement(value, "Item")
        Exit Sub
    End If
    ReplaceItem CLng(index), WrapElement(value, "Item")
End Sub

Friend Sub ReplaceCollectionItem( _
    ByVal index As Long, _
    ByVal wrapped As ROneCOne _
)
    Dim current As ROneCOne
    Dim position As Long
    Dim replacement As Collection

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
    mVersion = mVersion + 1
    NotifyCollectionChanged "Replace"
End Sub

Private Function RemoveCollectionValue( _
    ByVal target As ROneCOne, _
    ByVal value As Variant, _
    ByVal memberName As String _
) As Boolean
    Dim itemIndex As Long

    If target.IsDictionaryCollection Or _
        target.InternalCollectionKind = COLLECTION_KEYED Then
        itemIndex = target.FindKeyIndex(target.WrapKey(value, memberName), memberName)
    Else
        itemIndex = target.FindElementIndex( _
            target.WrapElement(value, memberName), memberName)
    End If
    If itemIndex < 0 Then Exit Function
    target.RemoveCollectionAt itemIndex
    RemoveCollectionValue = True
End Function

Friend Sub RemoveCollectionAt(ByVal index As Long)
    mItems.Remove index + 1
    If IsDictionaryCollection Or mCollectionKind = COLLECTION_KEYED Then
        mKeys.Remove index + 1
    ElseIf IsPriorityQueueKind(mCollectionKind) Then
        mPriorities.Remove index + 1
    End If
    mVersion = mVersion + 1
    NotifyCollectionChanged "Remove"
End Sub

Private Function CloneGenericCollection( _
    ByVal source As ROneCOne _
) As ROneCOne
    Dim clone As ROneCOne
    Dim wrapped As ROneCOne

    Set clone = New ROneCOne
    clone.ConfigureGenericCollection source.InternalRole, _
        source.InternalCollectionKind, source.InternalValueType, _
        source.InternalValueObjectType, source.InternalKeyType, _
        source.InternalKeyObjectType, source.InternalQueryComparer
    For Each wrapped In source.InternalCollectionItems
        clone.InternalAppendCollectionItem wrapped
    Next wrapped
    For Each wrapped In source.InternalCollectionKeys
        clone.InternalAppendCollectionKey wrapped
    Next wrapped
    For Each wrapped In source.InternalCollectionPriorities
        clone.InternalAppendCollectionPriority wrapped
    Next wrapped
    clone.ConfigureCollectionKeySelector source.InternalKeySelector
    If Not source.InternalCollectionChanged Is Nothing Then
        clone.ConfigureObservableCollection
    End If
    clone.ConfigureReadOnlySource source.InternalSource
    clone.ConfigureCollectionCapacity source.InternalCollectionCapacity
    clone.ConfigureCollectionVersion source.InternalCollectionVersion
    clone.ConfigureCollectionBuilder source.InternalIsBuilder
    Set CloneGenericCollection = clone
End Function

Private Function CreateCollectionNode(ByVal index As Long) As ROneCOne
    Dim node As ROneCOne

    ValidateIndex index, mItems.Count, "LinkedListNode"
    Set node = New ROneCOne
    node.ConfigureCollectionNode mItems.Item(index + 1), Me
    Set CreateCollectionNode = node
End Function

Friend Function CreateNodeAt(ByVal index As Long) As ROneCOne
    Set CreateNodeAt = CreateCollectionNode(index)
End Function

Friend Function FindNodeWrappedIndex(ByVal wrapped As ROneCOne) As Long
    Dim index As Long

    FindNodeWrappedIndex = -1
    For index = 1 To mItems.Count
        If mItems.Item(index) Is wrapped Then
            FindNodeWrappedIndex = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function InsertRelativeNode( _
    ByVal node As ROneCOne, _
    ByVal value As Variant, _
    ByVal insertAfter As Boolean, _
    ByVal memberName As String _
) As ROneCOne
    Dim index As Long
    Dim insertAt As Long
    Dim wrapped As ROneCOne

    RequireLinkedList memberName
    RequireMutableCollection memberName
    RequireOwnedNode node, memberName
    index = FindNodeWrappedIndex(node.InternalQueryValue)
    If index < 0 Then
        RaiseContractError ERROR_INVALID_OPERATION, memberName, _
            "The node no longer belongs to this list."
    End If
    insertAt = index
    If insertAfter Then insertAt = insertAt + 1
    Set wrapped = WrapElement(value, memberName)
    If insertAt >= mItems.Count Then
        mItems.Add wrapped
    Else
        mItems.Add wrapped, Before:=insertAt + 1
    End If
    mVersion = mVersion + 1
    Set InsertRelativeNode = CreateCollectionNode(insertAt)
End Function

Private Function CreateDictionaryEntry( _
    ByVal keyWrapped As ROneCOne, _
    ByVal valueWrapped As ROneCOne _
) As ROneCOne
    Dim entry As ROneCOne

    Set entry = New ROneCOne
    entry.ConfigureDictionaryEntry keyWrapped, valueWrapped
    Set CreateDictionaryEntry = entry
End Function

Private Sub NotifyCollectionChanged(ByVal actionName As String)
    If mCollectionKind <> COLLECTION_OBSERVABLE Then Exit Sub
    If mCollectionChanged Is Nothing Then Exit Sub
    mCollectionChanged.Emit actionName
End Sub

Private Sub AssignReferenceWrapped( _
    ByVal referenceValue As ROneCOne, _
    ByVal wrapped As ROneCOne, _
    ByVal memberName As String _
)
    Dim byteValue As Byte
    Dim booleanValue As Boolean
    Dim currencyValue As Currency
    Dim dateValue As Date
    Dim doubleValue As Double
    Dim integerValue As Integer
    Dim longLongValue As LongLong
    Dim longValue As Long
    Dim singleValue As Single
    Dim scalarValue As Variant

    If referenceValue.InternalReferenceAddress = 0 Then
        ValidateWrappedType wrapped, referenceValue.InternalExpectedType, _
            referenceValue.InternalExpectedObjectType, memberName, "output"
        referenceValue.ConfigureManagedReferenceValue wrapped
        Exit Sub
    End If

    If wrapped.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "The output reference cannot receive an object."
    End If
    scalarValue = wrapped.InternalScalarValue
    If VarType(scalarValue) <> referenceValue.InternalExpectedType Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "The output reference type does not match the value type."
    End If
    Select Case referenceValue.InternalExpectedType
        Case vbByte
            byteValue = CByte(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, byteValue, 1
        Case vbInteger
            integerValue = CInt(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, integerValue, 2
        Case vbLong
            longValue = CLng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, longValue, 4
        Case vbLongLong
            longLongValue = CLngLng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, longLongValue, 8
        Case vbSingle
            singleValue = CSng(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, singleValue, 4
        Case vbDouble
            doubleValue = CDbl(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, doubleValue, 8
        Case vbCurrency
            currencyValue = CCur(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, currencyValue, 8
        Case vbDate
            dateValue = CDate(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, dateValue, 8
        Case vbBoolean
            booleanValue = CBool(scalarValue)
            CopyMemory ByVal referenceValue.InternalReferenceAddress, booleanValue, 2
    End Select
End Sub

Private Sub AssignDefaultReferenceResult(ByRef destination As Variant)
    If mExpectedType = vbObject Then
        Set destination = Nothing
        Exit Sub
    End If
    Select Case mExpectedType
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

Private Function CreateSequenceCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal typeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim elementObjectType As String
    Dim elementType As VbVarType
    Dim result As ROneCOne

    ResolveTypeDescriptor typeToken, elementType, elementObjectType, _
        "Collection factory"
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, CollectionUsesEquality(collectionKind), _
            "Collection comparer")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection roleValue, collectionKind, _
        elementType, elementObjectType, vbEmpty, vbNullString, comparerValue
    Set CreateSequenceCollection = result
End Function

Private Function CreateMapCollection( _
    ByVal roleValue As Long, _
    ByVal collectionKind As Long, _
    ByVal keyTypeToken As Variant, _
    ByVal valueTypeToken As Variant, _
    Optional ByVal comparer As Variant _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim keyObjectType As String
    Dim keyType As VbVarType
    Dim result As ROneCOne
    Dim valueObjectType As String
    Dim valueType As VbVarType

    ResolveTypeDescriptor keyTypeToken, keyType, keyObjectType, _
        "Collection key type"
    ResolveTypeDescriptor valueTypeToken, valueType, valueObjectType, _
        "Collection value type"
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer( _
            comparer, CollectionUsesEquality(collectionKind), _
            "Collection comparer")
    End If
    Set result = New ROneCOne
    result.ConfigureGenericCollection roleValue, collectionKind, _
        valueType, valueObjectType, keyType, keyObjectType, comparerValue
    Set CreateMapCollection = result
End Function

Private Function CollectionUsesEquality(ByVal collectionKind As Long) As Boolean
    Select Case collectionKind
        Case COLLECTION_SORTED_LIST, COLLECTION_SORTED_DICTIONARY
        Case COLLECTION_SORTED_SET, COLLECTION_PRIORITY_QUEUE
        Case COLLECTION_IMMUTABLE_SORTED_DICTIONARY
        Case COLLECTION_IMMUTABLE_SORTED_SET
        Case Else
            CollectionUsesEquality = True
    End Select
End Function

Private Sub RequireCompatibleSequence( _
    ByVal other As ROneCOne, _
    ByVal memberName As String _
)
    If InternalElementType <> other.InternalElementType Or _
        InternalElementObjectType <> other.InternalElementObjectType Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "Both sequences must have the same generic element type."
    End If
End Sub

Private Function CreateSetQuery( _
    ByVal queryOperator As Long, _
    ByVal other As Variant, _
    Optional ByVal comparer As Variant, _
    Optional ByVal memberName As String = "Set operator" _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim otherSequence As ROneCOne

    Set otherSequence = NormalizeMembershipSequence(other, memberName)
    RequireCompatibleSequence otherSequence, memberName
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, memberName)
    End If
    Set CreateSetQuery = CreateQuery( _
        queryOperator, Nothing, otherSequence, 0, _
        InternalElementType, InternalElementObjectType, comparerValue)
End Function

Private Function CreateSetByQuery( _
    ByVal queryOperator As Long, _
    ByVal other As Variant, _
    ByVal keySelector As Variant, _
    Optional ByVal comparer As Variant, _
    Optional ByVal memberName As String = "SetBy", _
    Optional ByVal requireCompatible As Boolean = False _
) As ROneCOne
    Dim comparerValue As ROneCOne
    Dim otherSequence As ROneCOne
    Dim selectorValue As ROneCOne

    Set selectorValue = NormalizeContextUnary(keySelector, memberName, False)
    Set otherSequence = NormalizeMembershipSequence(other, memberName)
    If requireCompatible Then RequireCompatibleSequence otherSequence, memberName
    If Not IsMissing(comparer) Then
        Set comparerValue = NormalizeComparer(comparer, True, memberName)
    End If
    Set CreateSetByQuery = CreateQuery( _
        queryOperator, selectorValue, otherSequence, 0, _
        InternalElementType, InternalElementObjectType, comparerValue)
End Function

Private Function ExecuteJoin( _
    ByVal inner As Variant, _
    ByVal outerKeySelector As Variant, _
    ByVal innerKeySelector As Variant, _
    ByVal resultSelector As Variant, _
    ByVal grouped As Boolean _
) As ROneCOne
    Dim group As ROneCOne
    Dim groupIndex As Long
    Dim innerSequence As ROneCOne
    Dim innerValue As ROneCOne
    Dim key As ROneCOne
    Dim lookup As ROneCOne
    Dim outerKey As ROneCOne
    Dim outerValue As ROneCOne
    Dim outerValues As Collection
    Dim result As ROneCOne
    Dim resultSelectorValue As ROneCOne
    Dim selected As ROneCOne

    Set innerSequence = NormalizeMembershipSequence(inner, "Join")
    Set outerKey = NormalizeContextUnary(outerKeySelector, "Join", False)
    Set lookup = innerSequence.ToLookup(innerKeySelector)
    Set resultSelectorValue = NormalizeBinaryDelegate(resultSelector, "Join")
    Set result = ROneCOne.ListOf(vbVariant)
    Set outerValues = MaterializeWrapped
    For Each outerValue In outerValues
        Set key = InvokeUnary(outerKey, outerValue)
        groupIndex = lookup.FindKeyIndex(key, "Join")
        If grouped Then
            If groupIndex < 0 Then
                Set group = New ROneCOne
                group.ConfigureEmptyGrouping key
            Else
                Set group = lookup.InternalCollectionItem(groupIndex)
            End If
            Set selected = InvokeBinary(resultSelectorValue, outerValue, _
                CreateObjectResult(group))
            result.InternalAddWrapped selected
        ElseIf groupIndex >= 0 Then
            Set group = lookup.InternalCollectionItem(groupIndex)
            For Each innerValue In group.MaterializeWrapped
                Set selected = InvokeBinary( _
                    resultSelectorValue, outerValue, innerValue)
                result.InternalAddWrapped selected
            Next innerValue
        End If
    Next outerValue
    Set ExecuteJoin = result
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

Private Function CreatePrimaryOrderQuery( _
    ByVal queryOperator As Long, _
    ByVal queryDelegate As ROneCOne, _
    ByVal elementType As VbVarType, _
    ByVal elementObjectType As String, _
    Optional ByVal queryComparer As ROneCOne _
) As ROneCOne
    Dim baseSource As ROneCOne
    Dim query As ROneCOne

    RequireSequence Me, "Order"
    Set baseSource = OrderingBaseSource
    Set query = New ROneCOne
    query.ConfigureQuery baseSource, queryOperator, queryDelegate, Nothing, 0, _
        elementType, elementObjectType, queryComparer
    Set CreatePrimaryOrderQuery = query
End Function

Private Function OrderingBaseSource() As ROneCOne
    Dim sourceValue As ROneCOne

    Set sourceValue = Me
    Do While sourceValue.InternalRole = ROLE_QUERY And _
        IsOrderingOperator(sourceValue.InternalOperator)
        Set sourceValue = sourceValue.InternalSource
    Loop
    Set OrderingBaseSource = sourceValue
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

Friend Function WrapElement( _
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
    If candidate.InternalRole <> ROLE_LIST And _
        candidate.InternalRole <> ROLE_QUERY And _
        candidate.InternalRole <> ROLE_DATA_VIEW And _
        Not IsGenericCollectionRole(candidate.InternalRole) Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            "This operation requires a typed sequence."
    End If
End Sub

Private Sub RequireOrderedQuery( _
    ByVal candidate As ROneCOne, _
    ByVal context As String _
)
    If candidate Is Nothing Then
        RaiseContractError ERROR_INVALID_ARGUMENT, context, _
            "The ordered query cannot be Nothing."
    End If
    If candidate.InternalRole <> ROLE_QUERY Or _
        Not IsOrderingOperator(candidate.InternalOperator) Then
        RaiseContractError ERROR_INVALID_OPERATION, context, _
            context & " is available only after an Order, OrderBy, or ThenBy operation."
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

Private Function NormalizeBinaryDelegate( _
    ByVal candidate As Variant, _
    ByVal memberName As String _
) As ROneCOne
    Dim candidateValue As ROneCOne
    Dim parameters As Collection
    Dim normalized As ROneCOne

    Set candidateValue = RequireROneCOne(candidate, memberName)
    If IsDelegateRole(candidateValue.InternalRole) Then
        If candidateValue.IsAction Then
            RaiseContractError ERROR_INVALID_OPERATION, memberName, _
                "The operation requires a Func."
        End If
        If candidateValue.Arity <> 2 And candidateValue.Arity <> -1 Then
            RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
                "The Func must accept exactly two values."
        End If
        Set NormalizeBinaryDelegate = candidateValue
        Exit Function
    End If
    RequireExpression candidateValue, memberName
    Set parameters = New Collection
    CollectExpressionParameters candidateValue, parameters
    If parameters.Count <> 2 Then
        RaiseContractError ERROR_ARITY_MISMATCH, memberName, _
            "The expression must reference exactly two parameters."
    End If
    Set normalized = New ROneCOne
    normalized.ConfigureLambda candidateValue, parameters
    Set NormalizeBinaryDelegate = normalized
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

Private Function ApplyTakeWhile( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not InvokePredicate(mQueryDelegate, wrapped, "TakeWhile") Then Exit For
        output.Add wrapped
    Next wrapped
    Set ApplyTakeWhile = output
End Function

Private Function ApplySkipWhile( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection
    Dim skipping As Boolean
    Dim wrapped As ROneCOne

    skipping = True
    Set output = New Collection
    For Each wrapped In sourceValues
        If skipping Then
            skipping = InvokePredicate(mQueryDelegate, wrapped, "SkipWhile")
        End If
        If Not skipping Then output.Add wrapped
    Next wrapped
    Set ApplySkipWhile = output
End Function

Private Function ApplyTakeLast( _
    ByVal sourceValues As Collection _
) As Collection
    Dim firstIndex As Long
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    firstIndex = sourceValues.Count - mQueryCount + 1
    If firstIndex < 1 Then firstIndex = 1
    For index = firstIndex To sourceValues.Count
        output.Add sourceValues.Item(index)
    Next index
    Set ApplyTakeLast = output
End Function

Private Function ApplySkipLast( _
    ByVal sourceValues As Collection _
) As Collection
    Dim index As Long
    Dim lastIndex As Long
    Dim output As Collection

    Set output = New Collection
    lastIndex = sourceValues.Count - mQueryCount
    If lastIndex < 0 Then lastIndex = 0
    For index = 1 To lastIndex
        output.Add sourceValues.Item(index)
    Next index
    Set ApplySkipLast = output
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

Private Function ApplyUnion(ByVal sourceValues As Collection) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set otherValues = mQueryValue.MaterializeWrapped
    For Each wrapped In otherValues
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyUnion = output
End Function

Private Function ApplyIntersect( _
    ByVal sourceValues As Collection _
) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set otherValues = mQueryValue.MaterializeWrapped
    Set output = New Collection
    For Each wrapped In sourceValues
        If WrappedCollectionContains(otherValues, wrapped, mQueryComparer) And _
            Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyIntersect = output
End Function

Private Function ApplyExcept(ByVal sourceValues As Collection) As Collection
    Dim otherValues As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set otherValues = mQueryValue.MaterializeWrapped
    Set output = New Collection
    For Each wrapped In sourceValues
        If Not WrappedCollectionContains(otherValues, wrapped, mQueryComparer) And _
            Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set ApplyExcept = output
End Function

Private Function ApplyDefaultIfEmpty( _
    ByVal sourceValues As Collection _
) As Collection
    Dim output As Collection

    If sourceValues.Count > 0 Then
        Set ApplyDefaultIfEmpty = CopyWrappedCollection(sourceValues)
        Exit Function
    End If
    Set output = New Collection
    If mQueryValue Is Nothing Then
        output.Add DefaultWrappedElement
    Else
        output.Add mQueryValue
    End If
    Set ApplyDefaultIfEmpty = output
End Function

Private Function ApplyChunk(ByVal sourceValues As Collection) As Collection
    Dim chunk As ROneCOne
    Dim index As Long
    Dim output As Collection

    Set output = New Collection
    For index = 1 To sourceValues.Count
        If (index - 1) Mod mQueryCount = 0 Then
            Set chunk = New ROneCOne
            chunk.ConfigureList mSource.InternalElementType, _
                mSource.InternalElementObjectType
            output.Add CreateObjectResult(chunk)
        End If
        chunk.InternalAddWrapped sourceValues.Item(index)
    Next index
    Set ApplyChunk = output
End Function

Private Function ApplySelectMany( _
    ByVal sourceValues As Collection _
) As Collection
    Dim nested As ROneCOne
    Dim nestedObject As Object
    Dim nestedValues As Collection
    Dim output As Collection
    Dim selected As ROneCOne
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        Set selected = InvokeUnary(mQueryDelegate, wrapped)
        If Not selected.InternalHasObjectValue Then
            RaiseContractError ERROR_TYPE_MISMATCH, "SelectMany", _
                "The selector must return a ROneCOne sequence."
        End If
        Set nestedObject = selected.InternalObjectValue
        If nestedObject Is Nothing Or Not TypeOf nestedObject Is ROneCOne Then
            RaiseContractError ERROR_TYPE_MISMATCH, "SelectMany", _
                "The selector must return a ROneCOne sequence."
        End If
        Set nested = nestedObject
        RequireSequence nested, "SelectMany"
        Set nestedValues = nested.MaterializeWrapped
        For Each selected In nestedValues
            ValidateElementWrapped selected, "SelectMany"
            output.Add selected
        Next selected
    Next wrapped
    Set ApplySelectMany = output
End Function

Private Function ApplyCast( _
    ByVal sourceValues As Collection, _
    ByVal filterOnly As Boolean _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In sourceValues
        If WrappedMatchesType(wrapped, mElementType, mElementObjectType) Then
            output.Add wrapped
        ElseIf Not filterOnly Then
            RaiseContractError ERROR_TYPE_MISMATCH, "Cast", _
                "A sequence value cannot be cast to the requested type."
        End If
    Next wrapped
    Set ApplyCast = output
End Function

Private Function WrappedMatchesType( _
    ByVal wrapped As ROneCOne, _
    ByVal expectedType As VbVarType, _
    ByVal expectedObjectType As String _
) As Boolean
    If expectedType = vbVariant Then
        WrappedMatchesType = True
        Exit Function
    End If
    If expectedType = vbObject Then
        If Not wrapped.InternalHasObjectValue Then Exit Function
        If wrapped.InternalObjectValue Is Nothing Then
            WrappedMatchesType = True
        ElseIf Len(expectedObjectType) = 0 Then
            WrappedMatchesType = True
        Else
            WrappedMatchesType = (StrComp( _
                TypeName(wrapped.InternalObjectValue), expectedObjectType, _
                vbTextCompare) = 0)
        End If
        Exit Function
    End If
    If wrapped.InternalHasObjectValue Then Exit Function
    WrappedMatchesType = (VarType(wrapped.InternalScalarValue) = expectedType)
End Function

Private Function ApplySetBy(ByVal sourceValues As Collection) As Collection
    Dim candidateKeys As Collection
    Dim key As ROneCOne
    Dim otherValues As Collection
    Dim output As Collection
    Dim seenKeys As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    Set seenKeys = New Collection
    Set otherValues = mQueryValue.MaterializeWrapped
    If mOperator = QUERY_UNION_BY Then
        AddDistinctSelectedValues sourceValues, output, seenKeys
        AddDistinctSelectedValues otherValues, output, seenKeys
        Set ApplySetBy = output
        Exit Function
    End If

    Set candidateKeys = DistinctQueryKeys(otherValues)
    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains( _
            seenKeys, key, mQueryComparer) Then
            If mOperator = QUERY_INTERSECT_BY Then
                If WrappedCollectionContains( _
                    candidateKeys, key, mQueryComparer) Then
                    output.Add wrapped
                    seenKeys.Add key
                End If
            ElseIf Not WrappedCollectionContains( _
                candidateKeys, key, mQueryComparer) Then
                output.Add wrapped
                seenKeys.Add key
            End If
        End If
    Next wrapped
    Set ApplySetBy = output
End Function

Private Sub AddDistinctSelectedValues( _
    ByVal sourceValues As Collection, _
    ByVal output As Collection, _
    ByVal seenKeys As Collection _
)
    Dim key As ROneCOne
    Dim wrapped As ROneCOne

    For Each wrapped In sourceValues
        Set key = InvokeUnary(mQueryDelegate, wrapped)
        If Not WrappedCollectionContains(seenKeys, key, mQueryComparer) Then
            seenKeys.Add key
            output.Add wrapped
        End If
    Next wrapped
End Sub

Private Function DistinctQueryKeys( _
    ByVal values As Collection _
) As Collection
    Dim output As Collection
    Dim wrapped As ROneCOne

    Set output = New Collection
    For Each wrapped In values
        If Not WrappedCollectionContains(output, wrapped, mQueryComparer) Then
            output.Add wrapped
        End If
    Next wrapped
    Set DistinctQueryKeys = output
End Function

Private Function DefaultWrappedElement() As ROneCOne
    If mElementType = vbObject Then
        Set DefaultWrappedElement = CreateObjectResult(Nothing)
        Exit Function
    End If
    Select Case mElementType
        Case vbString
            Set DefaultWrappedElement = CreateScalarResult(vbNullString)
        Case vbBoolean
            Set DefaultWrappedElement = CreateScalarResult(False)
        Case vbDate
            Set DefaultWrappedElement = CreateScalarResult(CDate(0))
        Case vbByte
            Set DefaultWrappedElement = CreateScalarResult(CByte(0))
        Case vbInteger
            Set DefaultWrappedElement = CreateScalarResult(CInt(0))
        Case vbLong
            Set DefaultWrappedElement = CreateScalarResult(CLng(0))
        Case vbLongLong
            Set DefaultWrappedElement = CreateScalarResult(CLngLng(0))
        Case vbSingle
            Set DefaultWrappedElement = CreateScalarResult(CSng(0))
        Case vbDouble
            Set DefaultWrappedElement = CreateScalarResult(CDbl(0))
        Case vbCurrency
            Set DefaultWrappedElement = CreateScalarResult(CCur(0))
        Case Else
            Set DefaultWrappedElement = CreateScalarResult(Empty)
    End Select
End Function

Private Function ApplyOrderChain() As Collection
    Dim comparers() As ROneCOne
    Dim current As ROneCOne
    Dim descending() As Boolean
    Dim indexes() As Long
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim items() As ROneCOne
    Dim key As ROneCOne
    Dim keyScalars() As Variant
    Dim keyTypes() As VbVarType
    Dim keys() As ROneCOne
    Dim level As ROneCOne
    Dim levelCount As Long
    Dim levelIndex As Long
    Dim levels As Collection
    Dim output As Collection
    Dim sourceValues As Collection
    Dim workspace() As Long

    Set levels = New Collection
    Set current = Me
    Do While current.InternalRole = ROLE_QUERY And _
        IsOrderingOperator(current.InternalOperator)
        If levels.Count = 0 Then
            levels.Add current
        Else
            levels.Add current, Before:=1
        End If
        Set current = current.InternalSource
    Loop

    Set sourceValues = current.MaterializeWrapped
    Set output = New Collection
    itemCount = sourceValues.Count
    If itemCount = 0 Then
        Set ApplyOrderChain = output
        Exit Function
    End If

    levelCount = levels.Count
    ReDim items(1 To itemCount)
    ReDim keys(1 To levelCount, 1 To itemCount)
    ReDim keyScalars(1 To levelCount, 1 To itemCount)
    ReDim keyTypes(1 To levelCount, 1 To itemCount)
    ReDim comparers(1 To levelCount)
    ReDim descending(1 To levelCount)
    ReDim indexes(1 To itemCount)
    ReDim workspace(1 To itemCount)

    For itemIndex = 1 To itemCount
        Set items(itemIndex) = sourceValues.Item(itemIndex)
        indexes(itemIndex) = itemIndex
    Next itemIndex
    For levelIndex = 1 To levelCount
        Set level = levels.Item(levelIndex)
        Set comparers(levelIndex) = level.InternalQueryComparer
        descending(levelIndex) = _
            IsDescendingOrderingOperator(level.InternalOperator)
        For itemIndex = 1 To itemCount
            Set key = InvokeUnary(level.InternalQueryDelegate, items(itemIndex))
            Set keys(levelIndex, itemIndex) = key
            If comparers(levelIndex) Is Nothing Then
                If key.InternalHasObjectValue Then
                    RaiseContractError ERROR_TYPE_MISMATCH, "Order", _
                        "Object values require an explicit ordering comparer."
                End If
                keyScalars(levelIndex, itemIndex) = key.InternalScalarValue
                keyTypes(levelIndex, itemIndex) = _
                    VarType(keyScalars(levelIndex, itemIndex))
            End If
        Next itemIndex
    Next levelIndex

    StableMergeOrdering indexes, workspace, keys, keyScalars, keyTypes, _
        comparers, descending, levelCount, itemCount
    For itemIndex = 1 To itemCount
        output.Add items(indexes(itemIndex))
    Next itemIndex
    Set ApplyOrderChain = output
End Function

Private Sub StableMergeOrdering( _
    ByRef indexes() As Long, _
    ByRef workspace() As Long, _
    ByRef keys() As ROneCOne, _
    ByRef keyScalars() As Variant, _
    ByRef keyTypes() As VbVarType, _
    ByRef comparers() As ROneCOne, _
    ByRef descending() As Boolean, _
    ByVal levelCount As Long, _
    ByVal itemCount As Long _
)
    Dim comparison As Long
    Dim leftEnd As Long
    Dim leftPosition As Long
    Dim leftStart As Long
    Dim outputPosition As Long
    Dim rightEnd As Long
    Dim rightPosition As Long
    Dim width As Long

    width = 1
    Do While width < itemCount
        leftStart = 1
        Do While leftStart <= itemCount
            leftEnd = leftStart + width - 1
            If leftEnd > itemCount Then leftEnd = itemCount
            rightEnd = leftStart + (width * 2) - 1
            If rightEnd > itemCount Then rightEnd = itemCount
            leftPosition = leftStart
            rightPosition = leftEnd + 1
            outputPosition = leftStart

            Do While leftPosition <= leftEnd And rightPosition <= rightEnd
                comparison = CompareOrderingIndexes( _
                    indexes(leftPosition), indexes(rightPosition), _
                    keys, keyScalars, keyTypes, comparers, descending, _
                    levelCount)
                If comparison <= 0 Then
                    workspace(outputPosition) = indexes(leftPosition)
                    leftPosition = leftPosition + 1
                Else
                    workspace(outputPosition) = indexes(rightPosition)
                    rightPosition = rightPosition + 1
                End If
                outputPosition = outputPosition + 1
            Loop
            Do While leftPosition <= leftEnd
                workspace(outputPosition) = indexes(leftPosition)
                leftPosition = leftPosition + 1
                outputPosition = outputPosition + 1
            Loop
            Do While rightPosition <= rightEnd
                workspace(outputPosition) = indexes(rightPosition)
                rightPosition = rightPosition + 1
                outputPosition = outputPosition + 1
            Loop
            For outputPosition = leftStart To rightEnd
                indexes(outputPosition) = workspace(outputPosition)
            Next outputPosition
            leftStart = leftStart + (width * 2)
        Loop
        width = width * 2
    Loop
End Sub

Private Function CompareOrderingIndexes( _
    ByVal leftIndex As Long, _
    ByVal rightIndex As Long, _
    ByRef keys() As ROneCOne, _
    ByRef keyScalars() As Variant, _
    ByRef keyTypes() As VbVarType, _
    ByRef comparers() As ROneCOne, _
    ByRef descending() As Boolean, _
    ByVal levelCount As Long _
) As Long
    Dim comparison As Long
    Dim levelIndex As Long

    For levelIndex = 1 To levelCount
        If comparers(levelIndex) Is Nothing Then
            comparison = CompareCachedOrderingKeys( _
                keyScalars(levelIndex, leftIndex), _
                keyTypes(levelIndex, leftIndex), _
                keyScalars(levelIndex, rightIndex), _
                keyTypes(levelIndex, rightIndex))
        Else
            comparison = CompareWithComparer( _
                keys(levelIndex, leftIndex), _
                keys(levelIndex, rightIndex), _
                comparers(levelIndex), "Order")
        End If
        If comparison < 0 Then
            comparison = -1
        ElseIf comparison > 0 Then
            comparison = 1
        End If
        If descending(levelIndex) Then
            comparison = -comparison
        End If
        If comparison <> 0 Then
            CompareOrderingIndexes = comparison
            Exit Function
        End If
    Next levelIndex

    If leftIndex < rightIndex Then
        CompareOrderingIndexes = -1
    ElseIf leftIndex > rightIndex Then
        CompareOrderingIndexes = 1
    End If
End Function

Private Function CompareCachedOrderingKeys( _
    ByVal leftValue As Variant, _
    ByVal leftType As VbVarType, _
    ByVal rightValue As Variant, _
    ByVal rightType As VbVarType _
) As Long
    If leftType = vbNull Or rightType = vbNull Then
        If leftType = rightType Then
            CompareCachedOrderingKeys = 0
        ElseIf leftType = vbNull Then
            CompareCachedOrderingKeys = -1
        Else
            CompareCachedOrderingKeys = 1
        End If
        Exit Function
    End If
    If IsOrderingNumericType(leftType) And _
        IsOrderingNumericType(rightType) Then
        CompareCachedOrderingKeys = CompareScalarValues( _
            leftValue, rightValue, "Order")
        Exit Function
    End If
    If leftType <> rightType Then
        RaiseIncompatibleOrderingTypes leftType, rightType, "Order"
    End If
    Select Case leftType
        Case vbEmpty
            CompareCachedOrderingKeys = 0
        Case vbBoolean
            If CBool(leftValue) = CBool(rightValue) Then
                CompareCachedOrderingKeys = 0
            ElseIf Not CBool(leftValue) Then
                CompareCachedOrderingKeys = -1
            Else
                CompareCachedOrderingKeys = 1
            End If
        Case vbString
            CompareCachedOrderingKeys = CLng(Sgn(StrComp( _
                CStr(leftValue), CStr(rightValue), vbBinaryCompare)))
        Case vbDate
            CompareCachedOrderingKeys = CompareScalarValues( _
                leftValue, rightValue, "Order")
        Case Else
            RaiseIncompatibleOrderingTypes leftType, rightType, "Order"
    End Select
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
    Dim leftType As VbVarType
    Dim rightScalar As Variant
    Dim rightType As VbVarType

    If leftValue.InternalHasObjectValue Or rightValue.InternalHasObjectValue Then
        RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
            "Object values require an explicit ordering comparer."
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

    leftType = VarType(leftScalar)
    rightType = VarType(rightScalar)
    If IsOrderingNumericType(leftType) And _
        IsOrderingNumericType(rightType) Then
        CompareWrapped = CompareScalarValues( _
            leftScalar, rightScalar, memberName)
        Exit Function
    End If
    If leftType <> rightType Then
        RaiseIncompatibleOrderingTypes leftType, rightType, memberName
    End If

    Select Case leftType
        Case vbEmpty
            CompareWrapped = 0
        Case vbBoolean
            If CBool(leftScalar) = CBool(rightScalar) Then
                CompareWrapped = 0
            ElseIf Not CBool(leftScalar) Then
                CompareWrapped = -1
            Else
                CompareWrapped = 1
            End If
        Case vbString
            CompareWrapped = CLng(Sgn(StrComp( _
                CStr(leftScalar), CStr(rightScalar), vbBinaryCompare)))
        Case vbDate
            CompareWrapped = CompareScalarValues( _
                leftScalar, rightScalar, memberName)
        Case Else
            RaiseIncompatibleOrderingTypes leftType, rightType, memberName
    End Select
End Function

Private Function CompareScalarValues( _
    ByVal leftScalar As Variant, _
    ByVal rightScalar As Variant, _
    ByVal memberName As String _
) As Long
    On Error GoTo ComparisonFailed
    If leftScalar = rightScalar Then
        CompareScalarValues = 0
    ElseIf leftScalar < rightScalar Then
        CompareScalarValues = -1
    Else
        CompareScalarValues = 1
    End If
    On Error GoTo 0
    Exit Function

ComparisonFailed:
    On Error GoTo 0
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "The selected keys cannot be compared safely."
End Function

Private Function IsOrderingNumericType( _
    ByVal valueType As VbVarType _
) As Boolean
    Select Case valueType
        Case vbByte, vbInteger, vbLong, vbLongLong
            IsOrderingNumericType = True
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            IsOrderingNumericType = True
    End Select
End Function

Private Sub RaiseIncompatibleOrderingTypes( _
    ByVal leftType As VbVarType, _
    ByVal rightType As VbVarType, _
    ByVal memberName As String _
)
    RaiseContractError ERROR_TYPE_MISMATCH, memberName, _
        "Incompatible Variant key types: " & _
        TypeNameForVarType(leftType) & " and " & _
        TypeNameForVarType(rightType) & "."
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

Friend Function NormalizeContextUnary( _
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
            candidate.InternalRole = ROLE_QUERY Or _
            IsGenericCollectionRole(candidate.InternalRole)
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

Private Function IsOrderingOperator(ByVal operatorCode As Long) As Boolean
    Select Case operatorCode
        Case QUERY_ORDER, QUERY_ORDER_DESCENDING
            IsOrderingOperator = True
        Case QUERY_ORDER_BY, QUERY_ORDER_BY_DESCENDING
            IsOrderingOperator = True
        Case QUERY_THEN_BY, QUERY_THEN_BY_DESCENDING
            IsOrderingOperator = True
    End Select
End Function

Private Function IsDescendingOrderingOperator( _
    ByVal operatorCode As Long _
) As Boolean
    IsDescendingOrderingOperator = _
        operatorCode = QUERY_ORDER_DESCENDING Or _
        operatorCode = QUERY_ORDER_BY_DESCENDING Or _
        operatorCode = QUERY_THEN_BY_DESCENDING
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
    Dim runtimeTarget As ROneCOne
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

    If TypeOf target Is ROneCOne Then
        Set runtimeTarget = target
        If runtimeTarget.InternalRole = ROLE_DATA_ROW Then
            Set ReadMember = runtimeTarget.DataFieldWrapped(memberName)
            Exit Function
        End If
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

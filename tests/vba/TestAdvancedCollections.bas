Attribute VB_Name = "TestAdvancedCollections"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mNextRow As Long
Private mCurrentTest As String

Public Sub RunROneCOneAdvancedCollectionTests()
    Dim capturedDescription As String
    Dim capturedNumber As Long
    Dim capturedSource As String

    On Error GoTo FatalFailure
    ResetResults

    mCurrentTest = "TestDictionary"
    TestDictionary
    mCurrentTest = "TestHashSet"
    TestHashSet
    mCurrentTest = "TestQueueStackAndPriorityQueue"
    TestQueueStackAndPriorityQueue
    mCurrentTest = "TestLinkedAndSortedCollections"
    TestLinkedAndSortedCollections
    mCurrentTest = "TestCollectionWrappers"
    TestCollectionWrappers
    mCurrentTest = "TestConcurrentCollections"
    TestConcurrentCollections
    mCurrentTest = "TestImmutableCollections"
    TestImmutableCollections
    mCurrentTest = "TestCompleteSetSurface"
    TestCompleteSetSurface
    mCurrentTest = "TestConcurrentAndBlockingSurface"
    TestConcurrentAndBlockingSurface
    mCurrentTest = "TestPriorityAndLinkedSurface"
    TestPriorityAndLinkedSurface
    mCurrentTest = "TestImmutableRangeSurface"
    TestImmutableRangeSurface
    mCurrentTest = vbNullString

    With ThisWorkbook.Worksheets("Advanced Collection Results")
        .Range("B2").Value2 = mPassed
        .Range("B3").Value2 = mFailed
        .Range("B4").Value2 = IIf(mFailed = 0, "PASS", "FAIL")
    End With
    Exit Sub

FatalFailure:
    capturedNumber = Err.Number
    capturedSource = Err.Source
    capturedDescription = Err.Description
    With ThisWorkbook.Worksheets("Advanced Collection Results")
        .Range("B4").Value2 = "ERROR"
        .Range("B5").Value2 = mCurrentTest & " | " & CStr(capturedNumber) & _
            " | " & capturedSource & " | " & capturedDescription
    End With
End Sub

Private Sub TestDictionary()
    Dim capacity As Long
    Dim dictionary As ROneCOne
    Dim keys As ROneCOne
    Dim returnedValue As Long
    Dim valueReference As ROneCOne

    Set dictionary = ROneCOne.DictionaryOf(vbString, vbLong)
    capacity = dictionary.EnsureCapacity(128&)
    AssertTrue "dictionary reserves hash capacity", capacity >= 128&
    dictionary.Add "Ada", CLng(36)
    dictionary.Item("Grace") = CLng(40)

    AssertEqual "dictionary count", CLng(2), dictionary.Count
    AssertEqual "dictionary item", CLng(36), dictionary("Ada")
    AssertTrue "dictionary contains key", dictionary.ContainsKey("Grace")
    AssertFalse "dictionary rejects duplicate TryAdd", _
        dictionary.TryAdd("Ada", CLng(99))

    Set valueReference = ROneCOne.RefLong(returnedValue)
    AssertTrue "dictionary TryGetValue", _
        dictionary.TryGetValue("Grace", valueReference)
    AssertEqual "dictionary TryGetValue result", CLng(40), returnedValue

    Set keys = dictionary.Keys
    AssertTrue "dictionary keys", keys.Contains("Ada")
    AssertTrue "dictionary remove", dictionary.Remove("Ada")
    AssertFalse "dictionary removed key", dictionary.ContainsKey("Ada")
    dictionary.TrimExcess
    AssertTrue "dictionary trims without losing keys", _
        dictionary.ContainsKey("Grace")
    AssertEqual "dictionary type name", "Dictionary<String, Long>", _
        dictionary.GenericTypeName
    dictionary.Clear
    dictionary.Add "Katherine", CLng(38)
    AssertEqual "dictionary reuses cleared hash index", CLng(38), _
        dictionary.Item("Katherine")
End Sub

Private Sub TestHashSet()
    Dim leftSet As ROneCOne
    Dim rightSet As ROneCOne

    Set leftSet = ROneCOne.HashSetOf(vbLong)
    AssertTrue "hash set reserves capacity", _
        leftSet.EnsureCapacity(64&) >= 64&
    AssertTrue "hash set first add", leftSet.TryAdd(CLng(1))
    AssertFalse "hash set duplicate add", leftSet.TryAdd(CLng(1))
    leftSet.UnionWith Array(CLng(2), CLng(3))

    Set rightSet = ROneCOne.HashSetOf(vbLong)
    rightSet.UnionWith Array(CLng(2), CLng(3), CLng(4))
    AssertTrue "hash set overlaps", leftSet.Overlaps(rightSet)
    AssertTrue "hash set set equals", _
        leftSet.SetEquals(Array(CLng(1), CLng(2), CLng(3)))
    leftSet.IntersectWith rightSet
    AssertEqual "hash set intersection", CLng(2), leftSet.Count
    leftSet.Clear
    AssertTrue "hash set reuses cleared hash index", leftSet.TryAdd(CLng(8))
    AssertTrue "hash set finds value after clear", leftSet.Contains(CLng(8))
End Sub

Private Sub TestQueueStackAndPriorityQueue()
    Dim priority As ROneCOne
    Dim queue As ROneCOne
    Dim stack As ROneCOne

    Set queue = ROneCOne.QueueOf(vbString)
    queue.Enqueue "first"
    queue.Enqueue "second"
    AssertEqual "queue peek", "first", queue.Peek
    AssertEqual "queue dequeue", "first", queue.Dequeue

    Set stack = ROneCOne.StackOf(vbString)
    stack.Push "first"
    stack.Push "second"
    AssertEqual "stack peek", "second", stack.Peek
    AssertEqual "stack pop", "second", stack.Pop

    Set priority = ROneCOne.PriorityQueueOf(vbString, vbLong)
    priority.Enqueue "later", CLng(20)
    priority.Enqueue "first", CLng(10)
    AssertEqual "priority queue dequeue", "first", priority.Dequeue
End Sub

Private Sub TestLinkedAndSortedCollections()
    Dim linked As ROneCOne
    Dim sortedDictionary As ROneCOne
    Dim sortedSet As ROneCOne

    Set linked = ROneCOne.LinkedListOf(vbString)
    linked.AddLast "middle"
    linked.AddFirst "first"
    linked.AddLast "last"
    AssertEqual "linked first", "first", linked.FirstNode.Value
    AssertEqual "linked last", "last", linked.LastNode.Value
    linked.RemoveFirst
    AssertEqual "linked remove first", "middle", linked.FirstNode.Value

    Set sortedDictionary = ROneCOne.SortedDictionaryOf(vbString, vbLong)
    sortedDictionary.Add "z", CLng(2)
    sortedDictionary.Add "a", CLng(1)
    AssertEqual "sorted dictionary first key", "a", _
        sortedDictionary.Keys.Item(0)

    Set sortedSet = ROneCOne.SortedSetOf(vbLong)
    sortedSet.UnionWith Array(CLng(5), CLng(1), CLng(3))
    AssertEqual "sorted set minimum", CLng(1), sortedSet.MinValue
    AssertEqual "sorted set maximum", CLng(5), sortedSet.MaxValue
End Sub

Private Sub TestCollectionWrappers()
    Dim changed As ROneCOne
    Dim keyed As ROneCOne
    Dim readOnlyValues As ROneCOne
    Dim values As ROneCOne

    Set values = ROneCOne.ListOf(vbLong, CLng(1), CLng(2))
    Set readOnlyValues = ROneCOne.ReadOnlyCollectionOf(values)
    AssertTrue "read-only reports state", readOnlyValues.IsReadOnly
    AssertEqual "read-only count", CLng(2), readOnlyValues.Count

    Set changed = ROneCOne.ObservableCollectionOf(vbLong)
    AssertEqual "observable starts empty", CLng(0), changed.Count
    AssertEqual "observable change event", CLng(0), _
        changed.CollectionChanged.HandlerCount

    Set keyed = ROneCOne.KeyedCollectionOf( _
        vbString, vbString, ROneCOne.Var(vbString).AsFunc)
    keyed.Add "Ada"
    AssertTrue "keyed collection contains key", keyed.ContainsKey("Ada")
End Sub

Private Sub TestConcurrentCollections()
    Dim bag As ROneCOne
    Dim dictionary As ROneCOne
    Dim queue As ROneCOne
    Dim returnedValue As Long
    Dim valueReference As ROneCOne

    Set dictionary = ROneCOne.ConcurrentDictionaryOf(vbString, vbLong)
    AssertTrue "concurrent dictionary TryAdd", _
        dictionary.TryAdd("Ada", CLng(36))
    AssertEqual "concurrent dictionary GetOrAdd", CLng(36), _
        dictionary.GetOrAdd("Ada", CLng(99))

    Set queue = ROneCOne.ConcurrentQueueOf(vbLong)
    queue.Enqueue CLng(7)
    Set valueReference = ROneCOne.RefLong(returnedValue)
    AssertTrue "concurrent queue TryDequeue", queue.TryDequeue(valueReference)
    AssertEqual "concurrent queue value", CLng(7), returnedValue

    Set bag = ROneCOne.ConcurrentBagOf(vbString)
    bag.Add "value"
    AssertEqual "concurrent bag count", CLng(1), bag.Count
End Sub

Private Sub TestImmutableCollections()
    Dim dictionary As ROneCOne
    Dim list As ROneCOne
    Dim nextDictionary As ROneCOne
    Dim nextList As ROneCOne
    Dim setValue As ROneCOne

    Set list = ROneCOne.ImmutableListOf(vbLong)
    Set nextList = list.Add(CLng(1))
    AssertEqual "immutable list source", CLng(0), list.Count
    AssertEqual "immutable list result", CLng(1), nextList.Count

    Set dictionary = ROneCOne.ImmutableDictionaryOf(vbString, vbLong)
    Set nextDictionary = dictionary.Add("Ada", CLng(36))
    AssertFalse "immutable dictionary source", dictionary.ContainsKey("Ada")
    AssertEqual "immutable dictionary result", CLng(36), nextDictionary("Ada")

    Set setValue = ROneCOne.ImmutableHashSetOf(vbString).Add("Ada")
    AssertTrue "immutable hash set", setValue.Contains("Ada")
End Sub

Private Sub TestCompleteSetSurface()
    Dim removed As Long
    Dim values As ROneCOne

    Set values = ROneCOne.HashSetOf(vbLong)
    values.UnionWith Array(CLng(1), CLng(2), CLng(3))
    AssertTrue "set subset", values.IsSubsetOf(Array(1&, 2&, 3&, 4&))
    AssertTrue "set proper subset", values.IsProperSubsetOf(Array(1&, 2&, 3&, 4&))
    AssertTrue "set superset", values.IsSupersetOf(Array(1&, 2&))
    AssertTrue "set proper superset", values.IsProperSupersetOf(Array(1&, 2&))
    values.ExceptWith Array(2&)
    AssertFalse "set ExceptWith", values.Contains(2&)
    values.SymmetricExceptWith Array(3&, 4&)
    AssertTrue "set symmetric add", values.Contains(4&)
    AssertFalse "set symmetric remove", values.Contains(3&)
    removed = values.RemoveWhere( _
        ROneCOne.Func("DelegateProcedures.IsEvenLong"))
    AssertEqual "set RemoveWhere count", CLng(1), removed
End Sub

Private Sub TestConcurrentAndBlockingSurface()
    Dim blocking As ROneCOne
    Dim concurrent As ROneCOne
    Dim removedValue As Long
    Dim valueReference As ROneCOne

    Set concurrent = ROneCOne.ConcurrentDictionaryOf(vbString, vbLong)
    concurrent.Add "Ada", 36&
    AssertTrue "concurrent TryUpdate", concurrent.TryUpdate("Ada", 37&, 36&)
    AssertEqual "concurrent AddOrUpdate existing", 38&, _
        concurrent.AddOrUpdate("Ada", 1&, 38&)
    Set valueReference = ROneCOne.RefLong(removedValue)
    AssertTrue "concurrent TryRemove", _
        concurrent.TryRemove("Ada", valueReference)
    AssertEqual "concurrent TryRemove value", 38&, removedValue

    Set blocking = ROneCOne.BlockingCollectionOf(vbLong, 2&)
    AssertTrue "blocking TryAdd", blocking.TryAddItem(1&)
    blocking.CompleteAdding
    AssertTrue "blocking adding completed", blocking.IsAddingCompleted
    AssertTrue "blocking TryTake", blocking.TryTake(valueReference)
    AssertTrue "blocking completed", blocking.IsCompleted
End Sub

Private Sub TestPriorityAndLinkedSurface()
    Dim linked As ROneCOne
    Dim middle As ROneCOne
    Dim priorityQueue As ROneCOne
    Dim priorityReference As ROneCOne
    Dim value As String
    Dim valueReference As ROneCOne

    Set priorityQueue = ROneCOne.PriorityQueueOf(vbString, vbLong)
    priorityQueue.Enqueue "later", 20&
    priorityQueue.Enqueue "first", 10&
    Set valueReference = ROneCOne.OutOf(vbString)
    Set priorityReference = ROneCOne.OutOf(vbLong)
    AssertTrue "priority TryPeek pair", _
        priorityQueue.TryPeekPair(valueReference, priorityReference)
    AssertEqual "priority peek value", "first", valueReference.Current
    AssertEqual "priority peek priority", 10&, priorityReference.Current
    value = priorityQueue.DequeueEnqueue("replacement", 15&)
    AssertEqual "priority DequeueEnqueue", "first", value
    AssertEqual "priority EnqueueDequeue", "urgent", _
        priorityQueue.EnqueueDequeue("urgent", 5&)
    AssertEqual "priority EnqueueDequeue preserves queue", _
        "replacement", priorityQueue.Peek

    Set linked = ROneCOne.LinkedListOf(vbString)
    Set middle = linked.AddLast("middle")
    linked.AddBefore middle, "first"
    linked.AddAfter middle, "last"
    AssertEqual "linked node previous", "first", middle.Previous.Value
    AssertEqual "linked node next", "last", middle.NextNode.Value
    AssertTrue "linked node owner", middle.List Is linked
    AssertEqual "linked find", "last", linked.FindNode("last").Value
End Sub

Private Sub TestImmutableRangeSurface()
    Dim dictionary As ROneCOne
    Dim list As ROneCOne
    Dim queue As ROneCOne

    Set list = ROneCOne.ImmutableListOf(vbLong).AddRangeImmutable( _
        Array(1&, 2&, 3&))
    Set list = list.SetItem(1&, 20&).RemoveRange(Array(1&, 3&))
    AssertEqual "immutable range count", 1&, list.Count
    AssertEqual "immutable set item", 20&, list.Item(0)

    Set dictionary = ROneCOne.ImmutableDictionaryOf(vbString, vbLong) _
        .Add("Ada", 36&).SetItem("Ada", 37&)
    AssertEqual "immutable dictionary SetItem", 37&, dictionary("Ada")

    Set queue = ROneCOne.ImmutableQueueOf(vbLong).Enqueue(1&).Enqueue(2&)
    Set queue = queue.DequeueImmutable
    AssertEqual "immutable queue dequeue", 2&, queue.Peek
End Sub

Private Sub ResetResults()
    mPassed = 0
    mFailed = 0
    mNextRow = 6
    With ThisWorkbook.Worksheets("Advanced Collection Results")
        .Range("A2:C400").ClearContents
        .Range("A2").Value2 = "Passed"
        .Range("A3").Value2 = "Failed"
        .Range("A4").Value2 = "Status"
        .Range("A5").Value2 = "Error"
    End With
End Sub

Private Sub AssertTrue(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, True, condition
End Sub

Private Sub AssertFalse(ByVal testName As String, ByVal condition As Boolean)
    AssertEqual testName, False, condition
End Sub

Private Sub AssertEqual( _
    ByVal testName As String, _
    ByVal expected As Variant, _
    ByVal actual As Variant _
)
    Dim passed As Boolean

    If IsObject(expected) Or IsObject(actual) Then
        passed = (TypeName(expected) = TypeName(actual))
    Else
        passed = (expected = actual)
    End If

    With ThisWorkbook.Worksheets("Advanced Collection Results")
        .Cells(mNextRow, 1).Value2 = testName
        .Cells(mNextRow, 2).Value2 = IIf(passed, "PASS", "FAIL")
        If Not passed Then
            .Cells(mNextRow, 3).Value2 = _
                "Expected " & CStr(expected) & ", got " & CStr(actual)
        End If
    End With
    mNextRow = mNextRow + 1
    If passed Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
    End If
End Sub

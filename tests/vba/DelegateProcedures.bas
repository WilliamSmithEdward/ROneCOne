Attribute VB_Name = "DelegateProcedures"
Option Explicit

#If Win64 Then
Private Declare PtrSafe Sub CopyPointer Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef destination As LongPtr, _
    ByRef source As LongPtr, _
    ByVal byteCount As LongPtr _
)
#End If

Private mTrace As String
Private mTotal As Long

Private Const OTHER_ERROR As Long = vbObjectError + 5202

Public Function AddValues(ByVal leftValue As Variant, ByVal rightValue As Variant) As Variant
    AddValues = leftValue + rightValue
End Function

Public Sub IncrementLong(ByRef value As Long)
    value = value + 1
End Sub

Public Sub ResetTrace()
    mTrace = vbNullString
End Sub

Public Sub RecordFirst(ByVal value As Variant)
    mTrace = mTrace & "first:" & CStr(value) & "|"
End Sub

Public Sub RecordSecond(ByVal value As Variant)
    mTrace = mTrace & "second:" & CStr(value) & "|"
End Sub

Public Function CurrentTrace() As String
    CurrentTrace = mTrace
End Function

Public Sub AccumulateLong(ByVal value As Variant)
    mTotal = mTotal + CLng(value)
End Sub

Public Sub ResetTotal()
    mTotal = 0
End Sub

Public Function CurrentTotal() As Long
    CurrentTotal = mTotal
End Function

Public Function IsEvenLong(ByVal value As Variant) As Variant
    IsEvenLong = (CLng(value) Mod 2 = 0)
End Function

Public Sub RecordWork()
    mTrace = mTrace & "work|"
End Sub

Public Sub HandleExpected(ByVal errorInfo As Variant)
    mTrace = mTrace & "caught:" & errorInfo.Message & "|"
End Sub

Public Sub RecordFinally()
    mTrace = mTrace & "finally|"
End Sub

Public Function OtherErrorNumber() As Long
    OtherErrorNumber = OTHER_ERROR
End Function

#If Win64 Then
Public Function NativeAddLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    NativeAddLong = leftValue + rightValue
End Function

Public Function NativeAddLongAddress() As LongPtr
    Dim procedureAddress As LongPtr

    CopyPointer procedureAddress, AddressOf NativeAddLong, LenB(procedureAddress)
    NativeAddLongAddress = procedureAddress
End Function

Public Sub NativeIncrementLong(ByRef value As Long)
    value = value + 1
End Sub

Public Function NativeIncrementLongAddress() As LongPtr
    Dim procedureAddress As LongPtr

    CopyPointer procedureAddress, AddressOf NativeIncrementLong, _
        LenB(procedureAddress)
    NativeIncrementLongAddress = procedureAddress
End Function
#End If

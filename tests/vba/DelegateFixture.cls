VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DelegateFixture"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mLastValue As Variant

Public Function DoubleValue(ByVal value As Variant) As Variant
    DoubleValue = value * 2
End Function

Public Function Run(ByVal value As Variant) As Variant
    Run = value + 100
End Function

Public Sub RecordValue(ByVal value As Variant)
    mLastValue = value
End Sub

Public Property Get LastValue() As Variant
    LastValue = mLastValue
End Property

Public Function EchoSelf() As DelegateFixture
    Set EchoSelf = Me
End Function

Public Function CustomerAtLeast40(ByVal customer As GenericCustomer) As Boolean
    If customer Is Nothing Then Exit Function
    CustomerAtLeast40 = (customer.Age >= 40)
End Function

Public Function GetCustomerName(ByVal customer As GenericCustomer) As Variant
    If customer Is Nothing Then
        GetCustomerName = vbNullString
    Else
        GetCustomerName = customer.CustomerName
    End If
End Function

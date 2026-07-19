VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoCustomer"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' A deliberately ordinary user-defined class used to prove that ROneCOne
' collections and LINQ operators preserve concrete VBA object types.

Private mAge As Long
Private mCity As String
Private mCustomerName As String

Public Property Get CustomerName() As String
    CustomerName = mCustomerName
End Property

Public Property Let CustomerName(ByVal value As String)
    mCustomerName = value
End Property

Public Property Get Age() As Long
    Age = mAge
End Property

Public Property Let Age(ByVal value As Long)
    mAge = value
End Property

Public Property Get City() As String
    City = mCity
End Property

Public Property Let City(ByVal value As String)
    mCity = value
End Property

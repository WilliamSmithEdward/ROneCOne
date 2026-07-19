VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoCustomerQuery"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Application methods become strongly named LINQ predicates and selectors
' through ROneCOne.FromMethod. No string expressions or runtime code generation
' are required; the only string is the normal VBA method name binding.

Private mMinimumAge As Long
Private mRequiredCity As String

Public Property Get MinimumAge() As Long
    MinimumAge = mMinimumAge
End Property

Public Property Let MinimumAge(ByVal value As Long)
    mMinimumAge = value
End Property

Public Property Get RequiredCity() As String
    RequiredCity = mRequiredCity
End Property

Public Property Let RequiredCity(ByVal value As String)
    mRequiredCity = value
End Property

Public Function MeetsMinimumAge(ByVal customer As DemoCustomer) As Boolean
    If customer Is Nothing Then Exit Function
    MeetsMinimumAge = (customer.Age >= mMinimumAge)
End Function

Public Function IsInRequiredCity(ByVal customer As DemoCustomer) As Boolean
    If customer Is Nothing Then Exit Function
    IsInRequiredCity = (StrComp(customer.City, mRequiredCity, vbTextCompare) = 0)
End Function

Public Function SelectName(ByVal customer As DemoCustomer) As Variant
    If customer Is Nothing Then Exit Function
    SelectName = customer.CustomerName
End Function

Public Function SelectAge(ByVal customer As DemoCustomer) As Variant
    If customer Is Nothing Then Exit Function
    SelectAge = customer.Age
End Function

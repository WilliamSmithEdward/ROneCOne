VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericCustomer"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mName As String
Private mAge As Long
Private mActive As Boolean
Private mCity As String
Private mManager As GenericCustomer
Private mReports As ROneCOne

Public Property Get CustomerName() As String
    CustomerName = mName
End Property

Public Property Let CustomerName(ByVal value As String)
    mName = value
End Property

Public Property Get Age() As Long
    Age = mAge
End Property

Public Property Let Age(ByVal value As Long)
    mAge = value
End Property

Public Property Get Active() As Boolean
    Active = mActive
End Property

Public Property Let Active(ByVal value As Boolean)
    mActive = value
End Property

Public Property Get City() As String
    City = mCity
End Property

Public Property Let City(ByVal value As String)
    mCity = value
End Property

Public Property Get Manager() As GenericCustomer
    Set Manager = mManager
End Property

Public Property Set Manager(ByVal value As GenericCustomer)
    Set mManager = value
End Property

Public Property Get Reports() As ROneCOne
    Set Reports = mReports
End Property

Public Property Set Reports(ByVal value As ROneCOne)
    Set mReports = value
End Property

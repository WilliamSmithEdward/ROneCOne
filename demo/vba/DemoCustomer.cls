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

' ROneCOne 1.10.0, released 2026-09-22
'
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

' ============================================================================
' A plain customer object used by the collections tutorial
' ----------------------------------------------------------------------------
' This is an everyday VBA class with no ROneCOne magic in it: a customer has a
' name, an age, a city, an optional manager (another customer), and a list of
' direct reports. It exists so the collections demo can hold real objects of
' your own making, not just numbers and text.
'
' The point it proves: when you put DemoCustomer objects into a ROneCOne list
' and run queries over them, what comes back out is still a DemoCustomer, with
' all its properties intact. Your objects go in and come out unchanged, so you
' keep working with them exactly as you defined them here.
' ============================================================================

Private mAge As Long
Private mCity As String
Private mCustomerName As String
Private mManager As DemoCustomer
Private mReports As ROneCOne

Public Property Get CustomerName() As String
    CustomerName = mCustomerName
End Property

Public Property Let CustomerName(ByVal incoming As String)
    mCustomerName = incoming
End Property

Public Property Get Age() As Long
    Age = mAge
End Property

Public Property Let Age(ByVal incoming As Long)
    mAge = incoming
End Property

Public Property Get City() As String
    City = mCity
End Property

Public Property Let City(ByVal incoming As String)
    mCity = incoming
End Property

Public Property Get Manager() As DemoCustomer
    Set Manager = mManager
End Property

Public Property Set Manager(ByVal incoming As DemoCustomer)
    Set mManager = incoming
End Property

Public Property Get Reports() As ROneCOne
    Set Reports = mReports
End Property

Public Property Set Reports(ByVal incoming As ROneCOne)
    Set mReports = incoming
End Property

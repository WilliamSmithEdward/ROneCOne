VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SalesRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================================
' A plain sales row used by the Excel Tables tutorial
' ----------------------------------------------------------------------------
' An everyday VBA class with nothing special in it: three properties whose
' names match the column headers of the Sales table on the sheet. That match
' is the whole trick behind mapping in both directions. ToObjects fills these
' properties by name from each table row, and DataTableFromObjects reads them
' back out by name into a new table.
'
' Property names are matched case-insensitively through VBA's own member
' dispatch, so Amount, amount, and AMOUNT all bind to the same property.
' ============================================================================

Private mAmount As Double
Private mRegion As String
Private mRep As String

Public Property Get Region() As String
    Region = mRegion
End Property

Public Property Let Region(ByVal value As String)
    mRegion = value
End Property

Public Property Get Rep() As String
    Rep = mRep
End Property

Public Property Let Rep(ByVal value As String)
    mRep = value
End Property

Public Property Get Amount() As Double
    Amount = mAmount
End Property

Public Property Let Amount(ByVal value As Double)
    mAmount = value
End Property

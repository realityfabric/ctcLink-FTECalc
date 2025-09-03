VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FTECalc")
Option Explicit

Private Type TArrayContainer
    Data() As Variant
    Rows As Long
    Columns As Long
End Type
Private this As TArrayContainer

Public Property Get Data() As Variant() '(ByVal Row As Long, ByVal Column As Long As Variant
    Data = this.Data '(Row, Column)
End Property

Public Sub SetData(ByVal Rows As Long, ByVal Columns As Long, ByVal DataArray As Variant)
    this.Rows = Rows
    this.Columns = Columns
    this.Data = DataArray
End Sub

Public Property Get Rows() As Long
    Rows = this.Rows
End Property

Public Property Get Columns() As Long
    Columns = this.Columns
End Property


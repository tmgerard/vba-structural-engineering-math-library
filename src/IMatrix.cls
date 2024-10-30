VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit

Public Property Get Columns() As Long
End Property

Public Property Get Rows() As Long
End Property

Public Property Get ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long) As Double
End Property

Public Property Let ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long, ByRef value As Double)
End Property

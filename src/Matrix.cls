VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit
Implements IMatrix

Public Enum MatrixErrors
    BadRows = 5000 + vbObjectError
    BadColumns
    MatrixSizeMismatch
End Enum

Private Type TMatrix
    numRows As Long
    numColumns As Long
    Data() As Double    ' dense row-major storage scheme
End Type

Private this As TMatrix

Public Function SetSize(ByRef numRows As Long, ByRef numColumns As Long) As Matrix

    If numRows <= 0 Then
        Err.Raise Number:=MatrixErrors.BadRows, _
                  Source:="Matrix.SetSize", _
                  Description:="Row size cannot be zero or negative."
    End If
    
    If numColumns <= 0 Then
        Err.Raise Number:=MatrixErrors.BadColumns, _
                  Source:="Matrix.SetSize", _
                  Description:="Column size cannot be zero or negative."
    End If

    With this
        .numRows = numRows
        .numColumns = numColumns
        ReDim .Data(numRows * numColumns - 1)   ' storing matrix in a one-dimensional array
    End With
    
    Set SetSize = Me
    
End Function

Public Property Get Rows() As Long
    Rows = this.numRows
End Property

Public Property Get Columns() As Long
    Columns = this.numColumns
End Property

Public Property Get ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long) As Double
    ValueAt = this.Data(rowIndex * this.numColumns + columnIndex)
End Property

Public Property Let ValueAt(ByRef rowIndex As Long, ByRef columnIndex As Long, ByRef value As Double)
    this.Data(rowIndex * this.numColumns + columnIndex) = value
End Property

Public Property Get ValueAtTransposed(ByRef rowIndex As Long, ByRef columnIndex As Long) As Double
    ValueAtTransposed = Me.ValueAt(columnIndex, rowIndex)
End Property

Public Function IsSquare() As Boolean
    IsSquare = (this.numColumns = this.numRows)
End Function

Public Function AddTo(ByRef rowIndex As Long, ByRef columnIndex As Long, ByRef value As Double) As Matrix
    Me.ValueAt(rowIndex, columnIndex) = Me.ValueAt(rowIndex, columnIndex) + value
    Set AddTo = Me
End Function

Public Function SetData(ByRef matrixData() As Double) As Matrix
    ' matrixData array must be stored in row-major format
    If Not ArrayLength(matrixData) = (this.numColumns * this.numRows) Then
        Err.Raise Number:=MatrixErrors.MatrixSizeMismatch, _
                  Source:="Matrix.SetData", _
                  Description:="Cannot set data: Size mismatch"
    End If
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = matrixData(index)
    Next index
    
    Set SetData = Me
End Function

Public Function SetIdentityRow(ByRef rowIndex As Long) As Matrix

    Dim columnIndex As Long
    For columnIndex = 0 To this.numColumns - 1
        If rowIndex = columnIndex Then
            Me.ValueAt(rowIndex, columnIndex) = 1
        Else
            Me.ValueAt(rowIndex, columnIndex) = 0
        End If
    Next columnIndex
    
    Set SetIdentityRow = Me
End Function

Public Function SetIdentityColumn(ByRef columnIndex As Long) As Matrix

    Dim rowIndex As Long
    For rowIndex = 0 To this.numRows - 1
        If rowIndex = columnIndex Then
            Me.ValueAt(rowIndex, columnIndex) = 1
        Else
            Me.ValueAt(rowIndex, columnIndex) = 0
        End If
    Next rowIndex
    
    Set SetIdentityColumn = Me
End Function

Public Function ScaleValues(ByRef factor As Double) As Matrix
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = factor * this.Data(index)
    Next index
    
    Set ScaleValues = Me
End Function

Public Function Equals(ByRef other As Matrix, Optional ByVal tolerance As Double = CompareTolerance) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Matrix Then ' other is not a Matrix
        Equals = False
    ElseIf Not this.numRows = other.Rows Then
        Equals = False
    ElseIf Not this.numColumns = other.Columns Then
        Equals = False
    Else
        Dim rowIndex As Long
        Dim columnIndex As Long
        For rowIndex = 0 To this.numRows - 1
            For columnIndex = 0 To this.numColumns - 1
                If Not Doubles.Equal( _
                        Me.ValueAt(rowIndex, columnIndex), _
                        other.ValueAt(rowIndex, columnIndex), _
                        tolerance) Then
                    Equals = False
                    Exit Function
                End If
            Next columnIndex
        Next rowIndex
        Equals = True
    End If

End Function

' IMatrix Implmentation
Private Property Get IMatrix_Columns() As Long
    IMatrix_Columns = Me.Columns
End Property

Private Property Get IMatrix_Rows() As Long
    IMatrix_Rows = Me.Rows
End Property

Private Property Let IMatrix_ValueAt(rowIndex As Long, columnIndex As Long, value As Double)
    Me.ValueAt(rowIndex, columnIndex) = value
End Property

Private Property Get IMatrix_ValueAt(rowIndex As Long, columnIndex As Long) As Double
    IMatrix_ValueAt = Me.ValueAt(rowIndex, columnIndex)
End Property

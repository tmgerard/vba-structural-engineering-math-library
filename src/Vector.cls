VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit
Implements IMatrix

Public Enum VectorErrors
    LengthMismatch = 6000 + vbObjectError
End Enum

Public Enum VectorOrientation
    RowVector
    ColumnVector
End Enum

Private Type TVector
    Data() As Double
    Length As Long
    Orientation As VectorOrientation
End Type

Private this As TVector

Private Sub Class_Initialize()
    this.Orientation = ColumnVector
End Sub

Public Function SetLength(ByRef vectorLength As Long) As Vector
    With this
        .Length = vectorLength
        ReDim .Data(vectorLength - 1)
    End With
    
    Set SetLength = Me
End Function

Public Property Get Length() As Long
    Length = this.Length
End Property

Public Property Get Orientation() As VectorOrientation
    Orientation = this.Orientation
End Property

Public Property Let Orientation(ByVal value As VectorOrientation)
    this.Orientation = value
End Property

Public Property Get ValueAt(ByRef index As Long) As Double
    ValueAt = this.Data(index)
End Property

Public Property Let ValueAt(ByRef index As Long, ByRef value As Double)
    this.Data(index) = value
End Property

Public Function AddTo(ByRef index As Long, ByRef amount As Double) As Vector
    this.Data(index) = this.Data(index) + amount
    Set AddTo = Me
End Function

Public Function SetData(ByRef vecData() As Double) As Vector
    If Not ArrayLength(this.Data) = ArrayLength(vecData) Then
        Err.Raise Number:=VectorErrors.LengthMismatch, _
                  Source:="Vector.SetData()", _
                  Description:="Vector data array length mismatch"
    End If
    
    Dim index As Long
    For index = LBound(this.Data) To UBound(this.Data)
        this.Data(index) = vecData(index)
    Next index
    
    Set SetData = Me
End Function

Public Function Equals(ByRef other As Vector) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Vector Then ' other is not a Vector
        Equals = False
    Else
        Dim index As Long
        For index = LBound(this.Data) To UBound(this.Data)
            If Not Doubles.Equal(Me.ValueAt(index), other.ValueAt(index)) Then
                Equals = False
                Exit Function
            End If
        Next index
        Equals = True
    End If

End Function

'IMatrix Implementation
Private Property Get IMatrix_Columns() As Long
    If Me.Orientation = ColumnVector Then
        IMatrix_Columns = 1
    Else
        IMatrix_Columns = Me.Length
    End If
End Property

Private Property Get IMatrix_Rows() As Long
    If Me.Orientation = RowVector Then
        IMatrix_Rows = 1
    Else
        IMatrix_Rows = Me.Length
    End If
End Property

Private Property Let IMatrix_ValueAt(rowIndex As Long, columnIndex As Long, value As Double)
    If Me.Orientation = ColumnVector Then
        
        If Not columnIndex = 0 Then
            Err.Raise Number:=9, _
                      Source:="Vector.IMatrix_ValueAt(Let)", _
                      Description:="Column index out of range. Only one column exists for a column vector."
        End If
        
        Me.ValueAt(rowIndex) = value
        
    Else ' row vector
    
        If Not rowIndex = 0 Then
            Err.Raise Number:=9, _
                      Source:="Vector.IMatrix_ValueAt(Let)", _
                      Description:="Row index out of range. Only one column exists for a row vector."
        End If
        
        Me.ValueAt(columnIndex) = value
        
    End If
End Property

Private Property Get IMatrix_ValueAt(rowIndex As Long, columnIndex As Long) As Double
    If Me.Orientation = ColumnVector Then
        
        If Not columnIndex = 0 Then
            Err.Raise Number:=9, _
                      Source:="Vector.IMatrix_ValueAt(Get)", _
                      Description:="Column index out of range. Only one column exists for a column vector."
        End If
        
        IMatrix_ValueAt = Me.ValueAt(rowIndex)
        
    Else ' row vector
    
        If Not rowIndex = 0 Then
            Err.Raise Number:=9, _
                      Source:="Vector.IMatrix_ValueAt(Get)", _
                      Description:="Row index out of range. Only one column exists for a row vector."
        End If
        
        IMatrix_ValueAt = Me.ValueAt(columnIndex)
        
    End If
End Property

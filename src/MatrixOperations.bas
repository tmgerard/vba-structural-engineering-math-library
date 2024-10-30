Attribute VB_Name = "MatrixOperations"
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit

Public Enum MatrixOperationErrors
    SizeMismatch = vbObjectError + 7000
End Enum

Public Function MatMult(ByRef matA As IMatrix, ByRef matB As IMatrix) As Matrix
    If Not matA.Columns = matB.Rows Then
        Err.Raise Number:=MatrixOperationErrors.SizeMismatch, _
                  Source:="MatrixOperations.MatMult", _
                  Description:="Matrices sizes are not compatible with matrix multiplication."
    End If
    
    Dim result As Matrix
    Set result = New Matrix
    result.SetSize matA.Rows, matB.Columns
    
    Dim sum As Double
    Dim I As Long
    Dim j As Long
    Dim k As Long
    For I = 0 To matA.Rows - 1
        For j = 0 To matB.Columns - 1
            
            sum = 0
            For k = 0 To matA.Columns - 1
                sum = sum + matA.ValueAt(I, k) * matB.ValueAt(k, j)
            Next k
            
            result.ValueAt(I, j) = sum
        Next j
    Next I
    
    Set MatMult = result
    
End Function

Public Function Transpose(ByRef mat As IMatrix) As IMatrix
    If TypeOf mat Is Vector Then
        Set Transpose = TransposeVector(mat)
    ElseIf TypeOf mat Is Matrix Then
        Set Transpose = TransposeMatrix(mat)
    Else
        Err.Raise Number:=GlobalErrors.UnsupportedType, _
                  Source:="MatrixOperations.Transpose", _
                  Description:="Type " & TypeName(mat) & " is not supported by the Transpose function"
    End If
End Function

Private Function TransposeMatrix(ByRef mat As Matrix) As Matrix
    Dim transposed As Matrix
    Set transposed = New Matrix
    transposed.SetSize mat.Columns, mat.Rows
    
    Dim col As Long
    Dim row As Long
    For row = 0 To transposed.Rows - 1
        For col = 0 To transposed.Columns - 1
            transposed.ValueAt(row, col) = mat.ValueAt(col, row)
        Next col
    Next row
    
    Set TransposeMatrix = transposed
End Function

Private Function TransposeVector(ByRef vec As Vector) As Vector
    Dim transposed As Vector
    Set transposed = New Vector
    transposed.SetLength vec.Length
    
    Dim I As Long
    For I = 0 To vec.Length - 1
        transposed.ValueAt(I) = vec.ValueAt(I)
    Next I
    
    If vec.Orientation = ColumnVector Then
        transposed.Orientation = RowVector
    Else
        transposed.Orientation = ColumnVector
    End If
    
    Set TransposeVector = transposed
End Function
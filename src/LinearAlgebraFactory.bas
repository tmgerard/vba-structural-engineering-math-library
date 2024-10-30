Attribute VB_Name = "LinearAlgebraFactory"
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit

Public Function CreateMatrix(ByVal numRows As Long, ByVal numColumns As Long) As Matrix

    Dim mat As Matrix
    Set mat = New Matrix
    mat.SetSize numRows, numColumns
    
    Set CreateMatrix = mat

End Function

Public Function CreateIdentityMatrix(ByVal size As Long) As Matrix

    Dim mat As Matrix
    Set mat = New Matrix
    mat.SetSize size, size
    
    Dim I As Long
    For I = 0 To size
        mat.ValueAt(I, I) = 1
    Next I
    
    Set CreateIdentityMatrix = mat

End Function

Public Function CreateVector(ByVal size As Long) As Vector

    Dim vec As Vector
    Set vec = New Vector
    vec.SetLength size
    
    Set CreateVector = vec

End Function

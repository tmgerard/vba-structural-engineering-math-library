Attribute VB_Name = "VectorOperations"
'@Folder("StructuralMath.LinearAlgebra")
Option Explicit

Public Enum VectorOperationErrors
    SizeMismatch = vbObjectError + 8000
    OrientationMismatch
End Enum

Public Function VecAdd(ByRef vecA As Vector, ByRef vecB As Vector)
    If Not vecA.Length = vecB.Length Then
        Err.Raise Number:=VectorOperationErrors.SizeMismatch, _
                  Source:="VectorOperations.VecAdd", _
                  Description:="Vector sizes are not compatible with vector addition."
    End If
    
    If Not vecA.Orientation = vecB.Orientation Then
        Err.Raise Number:=VectorOperationErrors.OrientationMismatch, _
                  Source:="VectorOperations.VecAdd", _
                  Description:="Vector orientations are not compatible with vector addition."
    End If
    
    Dim vctr As Vector
    Set vctr = CreateVector(vecA.Length)
    vctr.Orientation = vecA.Orientation
    
    Dim i As Long
    For i = 0 To vecA.Length - 1
        vctr.ValueAt(i) = vecA.ValueAt(i) + vecB.ValueAt(i)
    Next i
    
    Set VecAdd = vctr
End Function

Public Function VecSubtract(ByRef vecA As Vector, ByRef vecB As Vector)
    If Not vecA.Length = vecB.Length Then
        Err.Raise Number:=VectorOperationErrors.SizeMismatch, _
                  Source:="VectorOperations.VecSubtract", _
                  Description:="Vector sizes are not compatible with vector subtraction."
    End If
    
    If Not vecA.Orientation = vecB.Orientation Then
        Err.Raise Number:=VectorOperationErrors.OrientationMismatch, _
                  Source:="VectorOperations.VecSubtract", _
                  Description:="Vector orientations are not compatible with vector subtraction."
    End If
    
    Dim vctr As Vector
    Set vctr = CreateVector(vecA.Length)
    vctr.Orientation = vecA.Orientation
    
    Dim i As Long
    For i = 0 To vecA.Length - 1
        vctr.ValueAt(i) = vecA.ValueAt(i) - vecB.ValueAt(i)
    Next i
    
    Set VecSubtract = vctr
End Function

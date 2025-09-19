VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transform2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit

Const MatrixSize As Long = 3

Private Type TTransform2D
    transform As Matrix
End Type

Private this As TTransform2D

Private Sub Class_Initialize()
    Set this.transform = CreateIdentityMatrix(MatrixSize)
End Sub

Private Sub Class_Terminate()
    Set this.transform = Nothing
End Sub

Public Sub Clear()
    Set this.transform = CreateIdentityMatrix(MatrixSize)
End Sub

Public Function RotZ(ByVal angle As Double) As Transform2D
    Dim rotation As Matrix
    Set rotation = CreateMatrix(MatrixSize, MatrixSize)
    With rotation
        .ValueAt(0, 0) = Math.cos(angle)
        .ValueAt(0, 1) = -Math.sin(angle)
        .ValueAt(1, 0) = Math.sin(angle)
        .ValueAt(1, 1) = Math.cos(angle)
        .ValueAt(2, 2) = 1
    End With
    
    Set this.transform = MatMult(this.transform, rotation)
    
    Set RotZ = Me
End Function

Public Function Translate(ByVal dx As Double, ByVal dy As Double) As Transform2D
    Dim trans As Matrix
    Set trans = CreateIdentityMatrix(MatrixSize)
    With trans
        .ValueAt(0, 2) = dx
        .ValueAt(1, 2) = dy
    End With
    
    Set this.transform = MatMult(this.transform, trans)
    
    Set Translate = Me
End Function

Public Function ScaleBy(ByVal sx As Double, ByVal sy As Double) As Transform2D
    Dim scl As Matrix
    Set scl = CreateIdentityMatrix(MatrixSize)
    With scl
        .ValueAt(0, 0) = sx
        .ValueAt(1, 1) = sy
    End With
    
    Set this.transform = MatMult(this.transform, scl)
    
    Set ScaleBy = Me
End Function

Public Function ApplyTo(ByVal pt As Point2D) As Point2D
    ' convert to homogeneous coordinates to allow translation
    Dim pointH As Vector
    Set pointH = CreateVector(MatrixSize)
    With pointH
        .ValueAt(0) = pt.x
        .ValueAt(1) = pt.y
        .ValueAt(2) = 1
    End With
    
    Dim transformed As Matrix
    Set transformed = MatMult(this.transform, pointH)
    
    Set ApplyTo = CreatePoint2D(transformed.ValueAt(0, 0), transformed.ValueAt(1, 0))
End Function


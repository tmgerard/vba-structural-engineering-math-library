VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit

Private Type TLine2D
    Base As Point2D
    direction As Vector2D
End Type

Private this As TLine2D

Private Sub Class_Initialize()
    With this
        Set .Base = New Point2D
        Set .direction = New Vector2D
    End With
End Sub

Private Sub Class_Terminate()
    With this
        Set .Base = Nothing
        Set .direction = Nothing
    End With
End Sub

Public Property Get Base() As Point2D
    Set Base = this.Base
End Property

Public Property Set Base(ByRef point As Point2D)
    Set this.Base = point
End Property

Public Property Get direction() As Vector2D
    Set direction = this.direction
End Property

Public Property Set direction(ByRef Vector As Vector2D)
    Set this.direction = Vector
End Property

Public Function IsParallelTo(ByRef other As Line2D) As Boolean
    IsParallelTo = this.direction.IsParallelTo(other.direction)
End Function

Public Function IsPerpendicularTo(ByRef other As Line2D) As Boolean
    IsPerpendicularTo = this.direction.IsPerpendicularTo(other.direction)
End Function

Public Function PerpendicularThrough(ByRef point As Point2D) As Line2D

    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        .Base = point
        .direction = this.direction.Perpendicular
    End With
    
    Set PerpendicularThrough = newLine
End Function

Public Function ParallelThrough(ByRef point As Point2D) As Line2D

    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        .Base = point
        .direction = this.direction
    End With
    
    Set ParallelThrough = newLine
End Function

Public Function IntersectionWith(ByRef other As Line2D) As Point2D

    If Me.IsParallelTo(other) Then
        Set IntersectionWith = Nothing
        Exit Function
    End If
    
    Dim d1 As Vector2D
    Set d1 = this.direction
    
    Dim d2 As Vector2D
    Set d2 = other.direction
    
    Dim Cross As Double
    Cross = d1.Cross(d2)
    
    Dim delta As Vector2D
    Set delta = CreateVectorBetween(this.Base, other.Base)
    
    Dim t1 As Double
    t1 = (delta.u * d2.v - delta.v * d2.u) / Cross
    
    Set IntersectionWith = this.Base.Displaced(d1, t1)
End Function

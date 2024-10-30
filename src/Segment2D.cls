VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Segment2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit

Public Enum SegmentRatioError
    BadValue = 10000 + vbObjectError
End Enum

Private Const MinRatio As Double = 0#
Private Const MidRatio As Double = 0.5
Private Const MaxRatio As Double = 1#

Private Type TSegment2D
    StartPoint As Point2D
    EndPoint As Point2D
End Type

Private this As TSegment2D

Private Sub Class_Terminate()
    With this
        Set .StartPoint = Nothing
        Set .EndPoint = Nothing
    End With
End Sub

Public Property Get StartPoint() As Point2D
    Set StartPoint = this.StartPoint
End Property

Public Property Set StartPoint(ByVal value As Point2D)
    Set this.StartPoint = value
End Property

Public Property Get EndPoint() As Point2D
    Set EndPoint = this.EndPoint
End Property

Public Property Set EndPoint(ByVal value As Point2D)
    Set this.EndPoint = value
End Property

'@Description "Returns the length of the segment"
Public Property Get Length() As Double
Attribute Length.VB_Description = "Returns the length of the segment"
    Length = this.StartPoint.DistanceTo(this.EndPoint)
End Property

'@Description "Returns Line2D object perpendicular to segment through the point at the middle of the segment."
Public Function Bisector() As Line2D
Attribute Bisector.VB_Description = "Returns Line2D object perpendicular to segment through the point at the middle of the segment."
    Dim newLine As Line2D
    Set newLine = New Line2D
    With newLine
        Set .Base = Me.PointAtMid
        Set .direction = Me.DirectionVector.Perpendicular.Normalized
    End With
    
    Set Bisector = newLine
End Function

'@Description "Returns Point2D object corresponding to the closest point on the line segment to the given Point2D"
Public Function ClosestPointTo(ByVal pointNearSegment As Point2D) As Point2D
Attribute ClosestPointTo.VB_Description = "Returns Point2D object corresponding to the closest point on the line segment to the given Point2D"
    Dim vecToPoint As Vector2D
    Set vecToPoint = pointNearSegment.Subtract(this.StartPoint)

    Dim unitVec As Vector2D
    Set unitVec = Me.DirectionVector.Normalized

    Dim projection As Double
    projection = vecToPoint.ProjectionOver(unitVec)

    Dim newPoint As Point2D
    Set newPoint = New Point2D
    If projection <= 0 Then
        newPoint.x = this.StartPoint.x
        newPoint.y = this.StartPoint.y
    ElseIf projection >= Me.Length Then
        newPoint.x = this.EndPoint.x
        newPoint.y = this.EndPoint.y
    Else
        Set newPoint = this.StartPoint.Displaced(unitVec, projection)
    End If

    Set ClosestPointTo = newPoint
End Function

Public Function DirectionVector() As Vector2D
    Set DirectionVector = this.EndPoint.Subtract(this.StartPoint)
End Function

Public Function IntersectionWith(ByRef other As Segment2D) As Point2D

    Dim d1 As Vector2D
    Set d1 = Me.DirectionVector
    
    Dim d2 As Vector2D
    Set d2 = other.DirectionVector
    
    If d1.IsParallelTo(d2) Then
        Set IntersectionWith = Nothing
        Exit Function
    End If
    
    Dim Cross As Double
    Cross = d1.Cross(d2)
    
    Dim delta As Vector2D
    Set delta = other.StartPoint.Subtract(this.StartPoint)
    
    Dim t1 As Double
    t1 = (delta.u * d2.v - delta.v * d2.u) / Cross
    
    Dim t2 As Double
    t2 = (delta.u * d1.v - delta.v * d1.u) / Cross
    
    If IsValid(t1) And IsValid(t2) Then
        Set IntersectionWith = Me.PointAt(t1)
    Else
        Set IntersectionWith = Nothing
    End If
    
End Function

Public Function PointAt(ByVal lengthRatio As Double) As Point2D

    EnsureValidRatio lengthRatio

    Dim newPoint As Point2D
    Set newPoint = New Point2D

    Select Case lengthRatio
    Case Is <= MinRatio
        newPoint.x = this.StartPoint.x
        newPoint.y = this.StartPoint.y
    Case Is >= MaxRatio
        newPoint.x = this.EndPoint.x
        newPoint.y = this.StartPoint.y
    Case Else
        Set newPoint = this.StartPoint.Displaced(Me.DirectionVector, lengthRatio)
    End Select
    
    Set PointAt = newPoint
End Function

Public Function PointAtMid() As Point2D
    Set PointAtMid = Me.PointAt(MidRatio)
End Function

Private Function MakeValidRatio(ByRef value As Double) As Double
    If value < MinRatio Then
        MakeValidRatio = MinRatio
    ElseIf value > MaxRatio Then
        MakeValidRatio = MaxRatio
    Else
        MakeValidRatio = value
    End If
End Function

Private Sub EnsureValidRatio(ByRef ratio As Double)
    If ratio < MinRatio Or ratio > MaxRatio Then
        Err.Raise Number:=SegmentRatioError.BadValue, _
                  Source:="Segment2DRatio", _
                  Description:="Expected ratio to be in [0, 1] but was " & ratio
    End If
End Sub

Private Function IsValid(ByRef ratio As Double) As Boolean
    IsValid = Not ratio < MinRatio And Not ratio > MaxRatio
End Function

Public Function Equals(ByVal other As Segment2D) As Boolean

    If Me Is other Then
        Equals = True
    ElseIf Not TypeOf other Is Segment2D Then
        Equals = False
    Else
        Equals = this.StartPoint.Equals(other.StartPoint) And this.EndPoint.Equals(other.EndPoint)
    End If

End Function

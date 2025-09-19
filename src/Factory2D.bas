Attribute VB_Name = "Factory2D"
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit

Public Function CreateLine2D(ByRef pntOnLine As Point2D, ByRef dirVec As Vector2D) As Line2D
    Dim line As Line2D
    Set line = New Line2D
    Set line.Base = pntOnLine
    Set line.direction = dirVec
    
    Set CreateLine2D = line
End Function

Public Function CreatePoint2D(ByVal xValue As Double, ByVal yValue As Double) As Point2D
    Dim pnt As Point2D
    Set pnt = New Point2D
    pnt.x = xValue
    pnt.y = yValue
    
    Set CreatePoint2D = pnt
End Function

Public Function CreatePointPolar(ByVal radiusValue As Double, ByVal angleValue As Double) As PointPolar
    Dim pnt As PointPolar
    Set pnt = New PointPolar
    pnt.Radius = radiusValue
    pnt.Theta = angleValue
    
    Set CreatePointPolar = pnt
End Function

Public Function CreateSegment2D(ByRef startPnt As Point2D, ByRef endPnt As Point2D) As Segment2D
    Dim sgmnt As Segment2D
    Set sgmnt = New Segment2D
    With sgmnt
        Set .StartPoint = startPnt
        Set .EndPoint = endPnt
    End With
    
    Set CreateSegment2D = sgmnt
End Function

Public Function CreateVector2D(ByVal uValue As Double, ByVal vValue As Double) As Vector2D
    Dim vctr As Vector2D
    Set vctr = New Vector2D
    vctr.u = uValue
    vctr.v = vValue
    
    Set CreateVector2D = vctr
End Function

Public Function CreateVectorBetween(ByRef point1 As Point2D, ByRef point2 As Point2D) As Vector2D
    Dim vctr As Vector2D
    Set vctr = point2.Subtract(point1)
    
    Set CreateVectorBetween = vctr
End Function

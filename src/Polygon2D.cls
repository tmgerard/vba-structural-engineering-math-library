VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Polygon2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit

Public Enum Polygon2DErrors
    TooFewPoints = vbObjectError + 1000
End Enum

Private Const MinVertices As Integer = 3

Private Type TPolygon2D
    Points As Collection
End Type

Private this As TPolygon2D

Private Sub Class_Initialize()
    Set this.Points = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.Points = Nothing
End Sub

Public Property Get Points() As Collection
    Set Points = this.Points
End Property

Public Property Set Points(ByRef point2DCollection As Collection)

    Dim point As Variant
    For Each point In point2DCollection
        If TypeOf point Is Point2D Then
            this.Points.Add point
        Else    ' raise a type mismatch error
            Err.Raise Number:=13, _
                      Source:="Polygon2D.Points (Set)", _
                      Description:="Type mismatch (" & TypeName(point) & "). Point2D objects required"
        End If
        
    Next point

End Property

Public Sub Add(ByRef point As Point2D)
    this.Points.Add point
End Sub

Public Sub Remove(ByVal index As Long)
    this.Points.Remove index
End Sub

'@Description "Calculates the area of the polygon"
Public Function Area() As Double
Attribute Area.VB_Description = "Calculates the area of the polygon"
    If Not HasEnoughPoints Then
        Err.Raise Number:=Polygon2DErrors.TooFewPoints, _
                  Source:="Polygon2D.Area", _
                  Description:="Too few points defined (Minimum = " & MinVertices & ")."
    End If
    
    Area = Math.Abs(SignedArea)
End Function

'@Description "Returns a Point2D object representing the centroid of the polygon"
Public Function Centroid() As Point2D
Attribute Centroid.VB_Description = "Returns a Point2D object representing the centroid of the polygon"
    If Not HasEnoughPoints Then
        Err.Raise Number:=Polygon2DErrors.TooFewPoints, _
                  Source:="Polygon2D.Area", _
                  Description:="Too few points defined (Minimum = " & MinVertices & ")."
    End If
    
    Dim sumX As Double, sumY As Double
    sumX = 0
    sumY = 0
    
    Dim currentPoint As Point2D
    Dim nextPoint As Point2D
    Dim index As Long
    
    For index = 1 To this.Points.Count - 1
        
        Set currentPoint = this.Points.Item(index)
        Set nextPoint = this.Points.Item(index + 1)
        
        sumX = sumX + (currentPoint.x + nextPoint.x) * _
            (currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y)
        
        sumY = sumY + (currentPoint.y + nextPoint.y) * _
            (currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y)
        
    Next index
    
    ' check if points polygon points close shape
    Set currentPoint = this.Points.Item(this.Points.Count)
    Set nextPoint = this.Points.Item(1)
    If Not currentPoint.Equals(nextPoint) Then
        sumX = sumX + (currentPoint.x + nextPoint.x) * _
            (currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y)
        
        sumY = sumY + (currentPoint.y + nextPoint.y) * _
            (currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y)
    End If
    
    Set Centroid = CreatePoint2D(sumX / (6 * SignedArea), sumY / (6 * SignedArea))
    
End Function

'@Description "Returns True if given point falls within the polygon and False otherwise"
Public Function ContainsPoint(ByRef point As Point2D) As Boolean
Attribute ContainsPoint.VB_Description = "Returns True if given point falls within the polygon and False otherwise"
    
    Dim index As Long
    For index = 1 To this.Points.Count
        If point.Equals(this.Points.Item(index)) Then
            ContainsPoint = True
            Exit Function
        End If
    Next index
    
    Dim pointVectors As Collection
    Set pointVectors = New Collection
    
    Dim vertex As Point2D
    For Each vertex In this.Points
        pointVectors.Add vertex.Subtract(point)
    Next vertex
    
    Dim angleSum As Double
    angleSum = 0
    
    Dim v1 As Vector2D
    Dim v2 As Vector2D
    Dim nextIndex As Long
    For index = 1 To pointVectors.Count
    
        If index < pointVectors.Count Then
            nextIndex = index + 1
        Else
            nextIndex = 1
        End If
        
        Set v1 = pointVectors.Item(index)
        Set v2 = pointVectors.Item(nextIndex)
        
        angleSum = angleSum + v1.AngleTo(v2)
    
    Next index
    
    ContainsPoint = Doubles.Equal(angleSum, 2 * PI)
    
End Function

Public Function Count() As Long
    Count = this.Points.Count
End Function

Public Function Item(ByVal index As Long) As Point2D
    Set Item = this.Points.Item(index)
End Function

Private Function SignedArea() As Double
    Dim result As Double
    result = 0
    
    Dim currentPoint As Point2D
    Dim nextPoint As Point2D
    Dim index As Long
    
    For index = 1 To this.Points.Count - 1
        
        Set currentPoint = this.Points.Item(index)
        Set nextPoint = this.Points.Item(index + 1)
        
        result = result + currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y
        
    Next index
    
    ' check if points polygon points close shape
    Set currentPoint = this.Points.Item(this.Points.Count)
    Set nextPoint = this.Points.Item(1)
    If Not currentPoint.Equals(nextPoint) Then
        result = result + currentPoint.x * nextPoint.y - nextPoint.x * currentPoint.y
    End If
    
    SignedArea = 0.5 * result
    
End Function

Private Function HasEnoughPoints() As Boolean
    HasEnoughPoints = (this.Points.Count >= MinVertices)
End Function

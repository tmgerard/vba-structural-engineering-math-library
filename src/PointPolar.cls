VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointPolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.2D")
Option Explicit
' TODO: add a test module
Private Type TPointPolar
    rad As Double
    rot As Double
    radiusInitialized As Boolean
End Type

Private this As TPointPolar

Public Property Get Radius() As Double
    Radius = this.rad
End Property

Public Property Let Radius(ByVal value As Double)
    
    If value = 0 Then
        this.rot = 0
        this.rad = 0
    Else
        If value < 0 Then
            this.rad = -value
            this.rot = this.rot + PI
        Else
            this.rad = value
        End If
    End If
    this.radiusInitialized = True
End Property

Public Property Get Theta() As Double
    Theta = this.rot
End Property

Public Property Let Theta(ByVal value As Double)
    If this.rad = 0 And this.radiusInitialized Then
        this.rot = 0
    Else
        this.rot = SetRotationWithinRange(value)
    End If
End Property

Private Function SetRotationWithinRange(ByVal rtValue As Double) As Double
    ' rotations will be normalized to a value between negative pi and pi radians
    Dim TwoPi As Double
    TwoPi = 2 * PI
    
    Const IterationLimit As Long = 10000
    Dim loopCount As Long
    loopCount = 0
    While Math.Abs(rtValue) > PI
    
        loopCount = loopCount + 1
        If loopCount > IterationLimit Then ' infinte loop likely
            Err.Raise vbObjectError + 2000, _
                      "PolarPoint.SetRotationWithinRange", _
                      "Inifinite loop detected when aliasing rotation value!"
        End If
        
        rtValue = rtValue + PI
        rtValue = rtValue - Int(rtValue / TwoPi) * TwoPi
        rtValue = rtValue - this.rot - PI
    Wend
    
    SetRotationWithinRange = rtValue
    
End Function

Public Function ToPoint2D() As Point2D
    Dim newPoint As Point2D
    Set newPoint = New Point2D
    With newPoint
        .x = this.rad * Math.cos(this.rot)
        .y = this.rad * Math.sin(this.rot)
    End With
    
    Set ToPoint2D = newPoint
End Function

Public Function Equals(ByVal other As PointPolar, Optional ByVal tolerance As Double = CompareTolerance) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is PointPolar Then
        Equals = False
    Else
        If Math.Abs(Me.Theta) = PI Then ' angle aliasing can return either pi or -pi which are equivalent rotations
            Equals = Doubles.Equal(Me.Radius, other.Radius, tolerance) And _
                     Doubles.Equal(Math.Abs(Me.Theta), Math.Abs(other.Theta), tolerance)
        Else
            Equals = Doubles.Equal(Me.Radius, other.Radius, tolerance) And _
                     Doubles.Equal(Me.Theta, other.Theta, tolerance)
        End If
    End If

End Function

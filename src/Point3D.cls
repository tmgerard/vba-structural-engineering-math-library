VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Point3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.3D")
Option Explicit

Private Type TPoint3D
    x As Double
    y As Double
    z As Double
End Type

Private this As TPoint3D

Public Property Get x() As Double
    x = this.x
End Property

Public Property Let x(ByVal value As Double)
    this.x = value
End Property

Public Property Get y() As Double
    y = this.y
End Property

Public Property Let y(ByVal value As Double)
    this.y = value
End Property

Public Property Get z() As Double
    z = this.z
End Property

Public Property Let z(ByVal value As Double)
    this.z = value
End Property

Public Function DistanceTo(ByVal other As Point3D) As Double
    Dim delta_x As Double
    delta_x = other.x - this.x
    
    Dim delta_y As Double
    delta_y = other.y - this.y
    
    Dim delta_z As Double
    delta_z = other.z - this.z
    
    DistanceTo = Math.Sqr(delta_x ^ 2 + delta_y ^ 2 + delta_z ^ 2)
End Function

Public Function Subtract(ByVal other As Point3D) As Vector3D
    Dim subtractVector As Vector3D
    Set subtractVector = New Vector3D
    With subtractVector
        .u = this.x - other.x
        .v = this.y - other.y
        .w = this.z - other.z
    End With
    
    Set Subtract = subtractVector
End Function

Public Function Equals(ByVal other As Point3D, Optional ByVal tolerance As Double = CompareTolerance) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Point3D Then ' other is not a Point3D
        Equals = False
    Else    ' check that x, y, and z values are the same
        Equals = Doubles.Equal(this.x, other.x, tolerance) And _
                 Doubles.Equal(this.y, other.y, tolerance) And _
                 Doubles.Equal(this.z, other.z, tolerance)
    End If

End Function



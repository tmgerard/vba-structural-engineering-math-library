VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.AnalyticGeometry.3D")
Option Explicit

Private Type TVector3D
    u As Double
    v As Double
    w As Double
End Type

Private this As TVector3D

Public Property Get u() As Double
    u = this.u
End Property

Public Property Let u(ByVal value As Double)
    this.u = value
End Property

Public Property Get v() As Double
    v = this.v
End Property

Public Property Let v(ByVal value As Double)
    this.v = value
End Property

Public Property Get w() As Double
    w = this.w
End Property

Public Property Let w(ByVal value As Double)
    this.w = value
End Property

Public Function Add(ByVal other As Vector3D) As Vector3D
    Dim addVector As Vector3D
    Set addVector = New Vector3D
    With addVector
        .u = this.u + other.u
        .v = this.v + other.v
        .w = this.w + other.w
    End With
    
    Set Add = addVector
End Function

Public Function Subtract(ByVal other As Vector3D) As Vector3D
    Dim subtractVector As Vector3D
    Set subtractVector = New Vector3D
    With subtractVector
        .u = this.u - other.u
        .v = this.v - other.v
        .w = this.w - other.w
    End With
    
    Set Subtract = subtractVector
End Function

Public Function ScaledBy(ByVal factor As Double) As Vector3D
    Dim scaledVector As Vector3D
    Set scaledVector = New Vector3D
    With scaledVector
        .u = factor * this.u
        .v = factor * this.v
        .w = factor * this.w
    End With
    
    Set ScaledBy = scaledVector
End Function

Public Function Norm() As Double
    Norm = Math.Sqr(this.u ^ 2 + this.v ^ 2 + this.w ^ 2)
End Function

Public Function IsNormal() As Boolean
    IsNormal = Doubles.IsEffectivelyOne(Me.Norm)
End Function

Public Function Normalized() As Vector3D
    Set Normalized = Me.ScaledBy(1# / Me.Norm)
End Function

Public Function WithLength(ByVal Length As Double) As Vector3D
    Set WithLength = Me.Normalized.ScaledBy(Length)
End Function

Public Function Dot(ByVal other As Vector3D) As Double
    Dot = (this.u * other.u) + (this.v * other.v) + (this.w * other.w)
End Function

Public Function ProjectionOver(ByVal direction As Vector3D) As Double
    ProjectionOver = Me.Dot(direction.Normalized)
End Function

Public Function Cross(ByVal other As Vector3D) As Vector3D
    Dim crossed As Vector3D
    Set crossed = New Vector3D
    With crossed
        .u = this.v * other.w - this.w * other.v
        .v = -(this.u * other.w - this.w * other.u)
        .w = this.u * other.v - this.v * other.u
    End With
    
    Set Cross = crossed
End Function

Public Function AngleValueTo(ByVal other As Vector3D) As Double
    Dim dotProduct As Double
    dotProduct = Me.Dot(other)
    
    Dim normProduct As Double
    normProduct = Me.Norm * other.Norm

    AngleValueTo = Math2.Acos(dotProduct / normProduct)
End Function

' TODO: Fix 3D vector angle function
'Public Function AngleTo(ByVal other As Vector3D) As Double
'    Dim value As Double
'    value = Me.AngleValueTo(other)
'
'    Dim crossProduct As Double
'    crossProduct = Me.Cross(other)
'
'    AngleTo = Math.Sgn(crossProduct) * value
'End Function

Public Function Equals(ByVal other As Vector3D) As Boolean

    If Me Is other Then ' reference to the exact same object
        Equals = True
    ElseIf Not TypeOf other Is Vector3D Then ' other is not a Vector3D
        Equals = False
    Else    ' check that u, v, and w values are the same
        Equals = Doubles.Equal(this.u, other.u) And _
                 Doubles.Equal(this.v, other.v) And _
                 Doubles.Equal(this.w, other.w)
    End If

End Function

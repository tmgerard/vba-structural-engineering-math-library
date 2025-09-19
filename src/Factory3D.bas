Attribute VB_Name = "Factory3D"
'@Folder("StructuralMath.AnalyticGeometry.3D")
Option Explicit

Public Function CreatePoint3D(ByVal xValue As Double, ByVal yValue As Double, ByVal zValue As Double) As Point3D
    Dim pnt As Point3D
    Set pnt = New Point3D
    pnt.x = xValue
    pnt.y = yValue
    pnt.z = zValue
    
    Set CreatePoint3D = pnt
End Function

Public Function CreateVector3D(ByVal uValue As Double, ByVal vValue As Double, ByVal wValue As Double) As Vector3D
    Dim vctr As Vector3D
    Set vctr = New Vector3D
    vctr.u = uValue
    vctr.v = vValue
    vctr.w = wValue
    
    Set CreateVector3D = vctr
End Function

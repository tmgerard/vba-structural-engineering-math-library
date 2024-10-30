Attribute VB_Name = "Math2"
'@Folder("StructuralMath.Utilities")
Option Explicit

Public Function Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
    If xValue = 0 And yValue = 0 Then
        Atan2 = 0
    ElseIf xValue = 0 And yValue > 0 Then
        Atan2 = WorksheetFunction.Pi / 2#
    ElseIf xValue = 0 And yValue < 0 Then
        Atan2 = -WorksheetFunction.Pi / 2#
    ElseIf xValue > 0 Then
        Atan2 = Math.Atn(yValue / xValue)
    ElseIf xValue < 0 And yValue >= 0 Then
        Atan2 = Math.Atn(yValue / xValue) + WorksheetFunction.Pi
    Else
        Atan2 = Math.Atn(yValue / xValue) - WorksheetFunction.Pi
    End If
End Function

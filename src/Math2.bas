Attribute VB_Name = "Math2"
'@Folder("StructuralMath.Utilities")
Option Explicit

' Define PI to remove worksheet function dependency if used outside of Excel
Public Const PI As Double = 3.14159265358979

' Constant defining default tolerance for floating point value comparisons
Public Const CompareTolerance As Double = 0.00001

Public Function Degrees(ByVal rad As Double) As Double
    Degrees = rad * (180# / PI)
End Function

Public Function Radians(ByVal deg As Double) As Double
    Radians = deg * (PI / 180#)
End Function

Public Function Acos(ByVal value As Double) As Double
    
    If value < -1 Or value > 1 Then
        Err.Raise GlobalErrors.OutsideRange, _
                  "Math2.Acos", _
                  "Argument outside of the range -1 to 1"
    End If
    
    If value = 1 Then
        Acos = 0
    ElseIf value = -1 Then
        Acos = PI
    Else
        Acos = Atn(-value / Sqr(-value * value + 1)) + 2 * Atn(1)
    End If
End Function

Public Function Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
    If xValue = 0 And yValue = 0 Then
        Atan2 = 0
    ElseIf xValue = 0 And yValue > 0 Then
        Atan2 = PI / 2#
    ElseIf xValue = 0 And yValue < 0 Then
        Atan2 = -PI / 2#
    ElseIf xValue > 0 Then
        Atan2 = Math.Atn(yValue / xValue)
    ElseIf xValue < 0 And yValue >= 0 Then
        Atan2 = Math.Atn(yValue / xValue) + PI
    Else
        Atan2 = Math.Atn(yValue / xValue) - PI
    End If
End Function

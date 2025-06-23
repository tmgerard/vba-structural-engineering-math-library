Attribute VB_Name = "Doubles"
'@Folder("StructuralMath.Utilities")
Option Explicit

'@Description "Compares two doubles with a given tolerance for equality."
Public Function Equal(ByVal left As Double, ByVal right As Double, _
    Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute Equal.VB_Description = "Compares two doubles with a given tolerance for equality."
    
    Equal = Math.Abs(right - left) <= tolerance

End Function

'@Description "Checks if given double is within a given tolerance of zero."
Public Function IsEffectivelyZero(ByVal num As Double, Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute IsEffectivelyZero.VB_Description = "Checks if given double is within a given tolerance of zero."
    IsEffectivelyZero = Equal(num, 0#, tolerance)
End Function

'@Description "Checks if a given double is within a given tolerance of one."
Public Function IsEffectivelyOne(ByVal num As Double, Optional ByVal tolerance As Double = CompareTolerance) As Boolean
Attribute IsEffectivelyOne.VB_Description = "Checks if a given double is within a given tolerance of one."
    IsEffectivelyOne = Equal(num, 1#, tolerance)
End Function

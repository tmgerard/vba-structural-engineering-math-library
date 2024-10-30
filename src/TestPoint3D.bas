Attribute VB_Name = "TestPoint3D"
'@IgnoreModule UseMeaningfulName
'@TestModule
'@Folder("Tests.StructuralMath.AnalyticGeometry.3D")


Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private point1 As Point3D
Private point2 As Point3D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set point1 = New Point3D
    With point1
        .x = 0
        .y = 0
        .z = 0
    End With
    
    Set point2 = New Point3D
    With point2
        .x = 2
        .y = 2
        .z = 2
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set point1 = Nothing
    Set point2 = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Method")
Private Sub TestDistanceTo()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 3.46

    'Act:
    Dim actual As Double
    actual = point1.DistanceTo(point2)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual, 0.01)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestEqualsFalse()
    On Error GoTo TestFail
    
    'Assert:
    Assert.IsFalse point1.Equals(point2)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestEqualsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim equalPoint As Point3D
    Set equalPoint = New Point3D
    With equalPoint
        .x = 2
        .y = 2
        .z = 2
    End With
    
    'Act:
    
    'Assert:
    Assert.IsTrue point2.Equals(equalPoint)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Attribute VB_Name = "TestVector3D"
'@TestModule
'@Folder("Tests.StructuralMath.AnalyticGeometry.3D")


Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private vecA As Vector3D
Private vecB As Vector3D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set vecA = New Vector3D
    With vecA
        .u = 3
        .v = -3
        .w = 1
    End With
    
    Set vecB = New Vector3D
    With vecB
        .u = 4
        .v = 9
        .w = 2
    End With
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set vecA = Nothing
    Set vecB = Nothing
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
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector3D
    Set expected = New Vector3D
    With expected
        .u = 7
        .v = 6
        .w = 3
    End With
    
    'Act:
    Dim actual As Vector3D
    Set actual = vecA.Add(vecB)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestSubtract()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector3D
    Set expected = New Vector3D
    With expected
        .u = -1
        .v = -12
        .w = -1
    End With
    
    'Act:
    Dim actual As Vector3D
    Set actual = vecA.Subtract(vecB)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDot()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = -13
    
    'Act:
    Dim actual As Double
    actual = vecA.Dot(vecB)
    
    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestCross()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector3D
    Set expected = New Vector3D
    With expected
        .u = -15
        .v = -2
        .w = 39
    End With
    
    'Act:
    Dim actual As Vector3D
    Set actual = vecA.Cross(vecB)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestNorm()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Double
    expected = 4.35889894354067
    
    'Act:
    Dim actual As Double
    actual = vecA.Norm
    
    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestScale()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector3D
    Set expected = New Vector3D
    With expected
        .u = 9
        .v = -9
        .w = 3
    End With
    
    'Act:
    Dim actual As Vector3D
    Set actual = vecA.ScaledBy(3)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestAngleValueTo()
    On Error GoTo TestFail

    'Arrange:
    Dim expected As Double
    expected = 1.872 ' radians

    'Act:
    Dim actual As Double
    actual = vecA.AngleValueTo(vecB)

    'Assert:
    Assert.IsTrue Doubles.Equal(expected, actual, 0.001), "Expected: " & expected & " Actual: " & actual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

''@TestMethod("Method")
'Private Sub TestAngleToPositive()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim expected As Double
'    expected = Math2.PI / 2 ' radians
'
'    Dim vecX As Vector3D
'    Set vecX = New Vector3D
'    With vecX
'        .u = 1
'        .v = 0
'        .w = 0
'    End With
'
'    Dim vecY As Vector3D
'    Set vecY = New Vector3D
'    With vecY
'        .u = 0
'        .v = 1
'        .w = 0
'    End With
'
'    'Act:
'    Dim actual As Double
'    actual = vecX.AngleTo(vecY)
'
'    'Assert:
'    Assert.IsTrue Doubles.Equal(expected, actual, 0.001), "Expected: " & expected & " Actual: " & actual
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("Method")
'Private Sub TestAngleToNegative()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim expected As Double
'    expected = -Math2.PI / 2  ' radians
'
'    Dim vecX As Vector3D
'    Set vecX = New Vector3D
'    With vecX
'        .u = 1
'        .v = 0
'        .w = 0
'    End With
'
'    Dim vecY As Vector3D
'    Set vecY = New Vector3D
'    With vecY
'        .u = 0
'        .v = -1
'        .w = 0
'    End With
'
'    'Act:
'    Dim actual As Double
'    actual = vecX.AngleTo(vecY)
'
'    'Assert:
'    Assert.IsTrue Doubles.Equal(expected, actual, 0.001), "Expected: " & expected & " Actual: " & actual
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

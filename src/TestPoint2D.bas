Attribute VB_Name = "TestPoint2D"
'@IgnoreModule UseMeaningfulName
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.StructuralMath.AnalyticGeometry.2D")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

Private point1 As Point2D
Private point2 As Point2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set point1 = New Point2D
    With point1
        .x = 1
        .y = 2
    End With
    
    Set point2 = New Point2D
    With point2
        .x = 4
        .y = 6
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
    expected = 5

    'Act:
    Dim actual As Double
    actual = point1.DistanceTo(point2)

    'Assert:
    Assert.AreEqual expected, actual

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 5
        .y = 8
    End With

    'Act:
    Dim actual As Point2D
    Set actual = point1.Add(point2)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestSubtract()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector2D
    Set expected = New Vector2D
    With expected
        .u = -3
        .v = -4
    End With

    'Act:
    Dim actual As Vector2D
    Set actual = point1.Subtract(point2)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDisplacedOneTime()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = point2
    
    Dim displaceVector As Vector2D
    Set displaceVector = New Vector2D
    With displaceVector
        .u = 3
        .v = 4
    End With

    'Act:
    Dim actual As Point2D
    Set actual = point1.Displaced(displaceVector)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDisplacedTwoTimes()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 7
        .y = 10
    End With
    
    Dim displaceVector As Vector2D
    Set displaceVector = New Vector2D
    With displaceVector
        .u = 3
        .v = 4
    End With

    'Act:
    Dim actual As Point2D
    Set actual = point1.Displaced(displaceVector, 2)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDisplacedByHalf()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = New Point2D
    With expected
        .x = 2.5
        .y = 4
    End With
    
    Dim displaceVector As Vector2D
    Set displaceVector = New Vector2D
    With displaceVector
        .u = 3
        .v = 4
    End With

    'Act:
    Dim actual As Point2D
    Set actual = point1.Displaced(displaceVector, 0.5)

    'Assert:
    Assert.IsTrue actual.Equals(expected)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestToPointPolar()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As PointPolar
    Set expected = CreatePointPolar(7#, PI / 6)
    
    Dim pnt As Point2D
    Set pnt = CreatePoint2D(7# * Math.Sqr(3#) / 2#, 7# / 2#)
    
    'Act:
    Dim actual As PointPolar
    Set actual = pnt.ToPointPolar
    
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


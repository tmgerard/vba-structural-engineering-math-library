Attribute VB_Name = "TestPolygon2D"
'@TestModule
'@Folder("Tests.StructuralMath.AnalyticGeometry.2D")


Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

Private plygn As Polygon2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
    #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set plygn = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
    ' https://en.wikipedia.org/wiki/Shoelace_formula#Example
    Set plygn = New Polygon2D
    With plygn
        .Add CreatePoint2D(1, 6)
        .Add CreatePoint2D(3, 1)
        .Add CreatePoint2D(7, 2)
        .Add CreatePoint2D(4, 4)
        .Add CreatePoint2D(8, 5)
    End With
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Method")
Private Sub TestAreaPointsDoNotCloseShape()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 16.5
    
    'Act:
    Dim actual As Double
    actual = plygn.Area
    
    'Assert:
    Assert.AreEqual expected, actual, "Expected = " & expected & " vs. Actual = " & actual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestAreaPointsCloseShape()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 16.5
    plygn.Add CreatePoint2D(1, 6)
    
    'Act:
    Dim actual As Double
    actual = plygn.Area
    
    'Assert:
    Assert.AreEqual expected, actual, "Expected = " & expected & " vs. Actual = " & actual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestCentroidPointsDoNotCloseShape()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(3.88889, 3.66667)
    
    'Act:
    Dim actual As Point2D
    Set actual = plygn.Centroid
    
    'Assert:
    Assert.IsTrue expected.Equals(actual), "Expected =( " & expected.x & ", " & expected.y & _
        ") vs. Actual = (" & actual.x & ", " & actual.y & ")"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestCentroidPointsCloseShape()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(3.88889, 3.66667)
    plygn.Add CreatePoint2D(1, 6)
    
    'Act:
    Dim actual As Point2D
    Set actual = plygn.Centroid
    
    'Assert:
    Assert.IsTrue expected.Equals(actual), "Expected =( " & expected.x & ", " & expected.y & _
        ") vs. Actual = (" & actual.x & ", " & actual.y & ")"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestContainsPointNotCloseShape()
    On Error GoTo TestFail
    
    'Arrange
    Dim pnt As Point2D
    Set pnt = CreatePoint2D(3, 4)
    
    'Assert:
    Assert.IsTrue plygn.ContainsPoint(pnt)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestContainsPointClosedShape()
    On Error GoTo TestFail
    
    'Arrange
    Dim pnt As Point2D
    Set pnt = CreatePoint2D(3, 4)
    plygn.Add CreatePoint2D(1, 6)
    
    'Assert:
    Assert.IsTrue plygn.ContainsPoint(pnt)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDoesNotContainPointNotCloseShape()
    On Error GoTo TestFail
    
    'Arrange
    Dim pnt As Point2D
    Set pnt = CreatePoint2D(0, 0)
    
    'Assert:
    Assert.IsFalse plygn.ContainsPoint(pnt)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestDoesNotContainPointClosedShape()
    On Error GoTo TestFail
    
    'Arrange
    Dim pnt As Point2D
    Set pnt = CreatePoint2D(0, 0)
    plygn.Add CreatePoint2D(1, 6)
    
    'Assert:
    Assert.IsFalse plygn.ContainsPoint(pnt)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

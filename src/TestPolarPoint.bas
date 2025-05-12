Attribute VB_Name = "TestPolarPoint"
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
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Initialization")
Private Sub TestCreateZeroRadiusPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 0 ' Theta Angle
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = CreatePointPolar(0, PI)
    
    'Assert:
    Assert.AreEqual expected, pnt.Theta, "Expected = " & expected & " vs. Actual = " & pnt.Theta

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub TestCreateNonZeroRadiusPoint()
    On Error GoTo TestFail
    
    'Arrange:
    Const ExpectedRadius As Double = 1
    Dim ExpectedAngle As Double
    ExpectedAngle = PI
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = CreatePointPolar(1, PI)
    
    'Assert:
    Assert.AreEqual ExpectedRadius, pnt.Radius, "Expected Radius = " & ExpectedRadius & " vs. Actual Radius = " & pnt.Radius
    Assert.AreEqual ExpectedAngle, pnt.Theta, "Expected Angle = " & ExpectedAngle & " vs. Actual Angle = " & pnt.Theta

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub TestCreateWithLargeAngle()
    On Error GoTo TestFail
    
    'Arrange:
    Const ExpectedRadius As Double = 1
    Dim ExpectedAngle As Double
    ExpectedAngle = -PI / 2
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = CreatePointPolar(1, 7 * PI / 2)
    
    'Assert:
    Assert.IsTrue Doubles.Equal(ExpectedRadius, pnt.Radius), "Expected Radius = " & ExpectedRadius & " vs. Actual Radius = " & pnt.Radius
    Assert.IsTrue Doubles.Equal(ExpectedAngle, pnt.Theta), "Expected Angle = " & ExpectedAngle & " vs. Actual Angle = " & pnt.Theta

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Initialization")
Private Sub TestCreateSetAngleWithZeroRadius()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 0
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = New PointPolar
    pnt.Radius = 0
    pnt.Theta = PI / 2
    
    'Assert:
    Assert.AreEqual expected, pnt.Theta, "Expected Angle = " & expected & " vs. Actual Angle = " & pnt.Theta

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub TestCreateSetRadiusToZero()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 0
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = New PointPolar
    pnt.Theta = PI / 2
    pnt.Radius = 0  ' setting radius to zero after setting theta should force theta to zero
    
    'Assert:
    Assert.AreEqual expected, pnt.Theta, "Expected Angle = " & expected & " vs. Actual Angle = " & pnt.Theta

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialization")
Private Sub TestToPoint2D()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(-0.5, 0)
    Dim pnt As PointPolar
    
    'Act:
    Set pnt = CreatePointPolar(0.5, PI)
    Dim actual As Point2D
    Set actual = pnt.ToPoint2D
    
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

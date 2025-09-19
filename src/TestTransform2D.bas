Attribute VB_Name = "TestTransform2D"
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

Private point1 As Point2D
Private point2 As Point2D

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    
    Set point1 = New Point2D
    With point1
        .x = 0
        .y = 0
    End With
    
    Set point2 = New Point2D
    With point2
        .x = 1
        .y = 0
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
Private Sub TestTranslate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(3, 5)
    
    Dim trans As Transform2D
    Set trans = New Transform2D
    Set trans = trans.Translate(3, 5)
    
    'Act:
    Dim actual As Point2D
    Set actual = trans.ApplyTo(point1)
    
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
Private Sub TestRotZ()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(0, 1)
    
    Dim trans As Transform2D
    Set trans = New Transform2D
    Set trans = trans.RotZ(Radians(90))
    
    'Act:
    Dim actual As Point2D
    Set actual = trans.ApplyTo(point2)
    
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
Private Sub TestScale()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Point2D
    Set expected = CreatePoint2D(5, 0)
    
    Dim trans As Transform2D
    Set trans = New Transform2D
    Set trans = trans.ScaleBy(5, 0)
    
    'Act:
    Dim actual As Point2D
    Set actual = trans.ApplyTo(point2)
    
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

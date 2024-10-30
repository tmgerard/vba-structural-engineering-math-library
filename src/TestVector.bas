Attribute VB_Name = "TestVector"
'@TestModule
'@Folder("Tests.StructuralMath.LinearAlgebra")


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

'@TestMethod("Property")
Private Sub TestLength()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Long = 5
    
    Dim vec As Vector
    Set vec = New Vector
    Set vec = vec.SetLength(expected)

    'Act:

    'Assert:
    Assert.AreEqual expected, vec.Length

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Initialize")
Private Sub TestDefaultValueIsZero()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 0
    
    Dim vec As Vector
    Set vec = New Vector
    Set vec = vec.SetLength(2)

    'Act:

    'Assert:
    Assert.AreEqual expected, vec.ValueAt(0)
    Assert.AreEqual expected, vec.ValueAt(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Property")
Private Sub TestValueAt()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 10#
    
    Dim vec As Vector
    Set vec = New Vector
    Set vec = vec.SetLength(2)
    vec.ValueAt(0) = expected

    'Act:

    'Assert:
    Assert.AreEqual expected, vec.ValueAt(0)
    Assert.AreEqual 0#, vec.ValueAt(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Method")
Private Sub TestAddTo()
    On Error GoTo TestFail
    
    'Arrange:
    Const expected As Double = 10#
    
    Dim vec As Vector
    Set vec = New Vector
    Set vec = vec.SetLength(2).AddTo(0, expected)

    'Act:

    'Assert:
    Assert.AreEqual expected, vec.ValueAt(0)
    Assert.AreEqual 0#, vec.ValueAt(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

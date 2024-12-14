Attribute VB_Name = "TestVectorOperations"
'@TestModule
'@Folder("Tests.StructuralMath.LinearAlgebra")


Option Explicit
Option Private Module

Private vecA As Vector
Private vecB As Vector
Private vecC As Vector
Private vecD As Vector

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
    
    Set vecA = CreateVector(2)
    
    Set vecB = CreateVector(2)
    With vecB
        .ValueAt(0) = 1
        .ValueAt(1) = 1
    End With
    
    Set vecC = CreateVector(2)
    With vecC
        .Orientation = RowVector
        .ValueAt(0) = 1
        .ValueAt(1) = 1
    End With
    
    Set vecD = CreateVector(5)
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    
    Set vecA = Nothing
    Set vecB = Nothing
    Set vecC = Nothing
    Set vecD = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Function")
Private Sub TestAddCompatible()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector
    Set expected = CreateVector(2)
    With expected
        .ValueAt(0) = 1
        .ValueAt(1) = 1
    End With
    
    'Act:
    Dim actual As Vector
    Set actual = VectorOperations.VecAdd(vecA, vecB)
    
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

'@TestMethod("Function")
Private Sub TestSubtractCompatible()
    On Error GoTo TestFail
    
    'Arrange:
    Dim expected As Vector
    Set expected = CreateVector(2)
    With expected
        .ValueAt(0) = -1
        .ValueAt(1) = -1
    End With
    
    'Act:
    Dim actual As Vector
    Set actual = VectorOperations.VecSubtract(vecA, vecB)
    
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

'@TestMethod("Uncategorized")
Private Sub TestAddOrientationMismatch()
    Const ExpectedError As Long = VectorOperationErrors.OrientationMismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Vector
    Set actual = VectorOperations.VecAdd(vecA, vecC)
    
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub TestSubtractOrientationMismatch()
    Const ExpectedError As Long = VectorOperationErrors.OrientationMismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Vector
    Set actual = VectorOperations.VecSubtract(vecA, vecC)
    
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub TestAddSizeMismatch()
    Const ExpectedError As Long = VectorOperationErrors.SizeMismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Vector
    Set actual = VectorOperations.VecAdd(vecA, vecD)
    
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Uncategorized")
Private Sub TestSubtractSizeMismatch()
    Const ExpectedError As Long = VectorOperationErrors.SizeMismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim actual As Vector
    Set actual = VectorOperations.VecSubtract(vecA, vecD)
    
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

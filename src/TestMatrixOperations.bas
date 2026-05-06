Attribute VB_Name = "TestMatrixOperations"
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

'@TestMethod("Function")
Private Sub TestMatMult()
    On Error GoTo TestFail
    
    'Arrange:
    Dim A As Matrix
    Set A = CreateMatrix(2, 2)
    With A
        .ValueAt(0, 0) = 2
        .ValueAt(0, 1) = 1
        .ValueAt(1, 0) = 1
        .ValueAt(1, 1) = 4
    End With
    
    Dim B As Matrix
    Set B = CreateMatrix(2, 3)
    With B
        .ValueAt(0, 0) = 1
        .ValueAt(0, 1) = 2
        .ValueAt(0, 2) = 0
        .ValueAt(1, 0) = 0
        .ValueAt(1, 1) = 1
        .ValueAt(1, 2) = 2
    End With
    
    Dim expected As Matrix
    Set expected = CreateMatrix(2, 3)
    With expected
        .ValueAt(0, 0) = 2
        .ValueAt(0, 1) = 5
        .ValueAt(0, 2) = 2
        .ValueAt(1, 0) = 1
        .ValueAt(1, 1) = 6
        .ValueAt(1, 2) = 8
    End With
    
    'Act:
    Dim actual As Matrix
    Set actual = MatMult(A, B)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Expected Error")
Private Sub TestMatMultMismatchSize()
    Const ExpectedError As Long = MatrixOperationErrors.SizeMismatch
    On Error GoTo TestFail
    
    'Arrange:
    Dim A As Matrix
    Set A = CreateMatrix(2, 2)
    With A
        .ValueAt(0, 0) = 2
        .ValueAt(0, 1) = 1
        .ValueAt(1, 0) = 1
        .ValueAt(1, 1) = 4
    End With
    
    Dim B As Matrix
    Set B = CreateMatrix(2, 3)
    With B
        .ValueAt(0, 0) = 1
        .ValueAt(0, 1) = 2
        .ValueAt(0, 2) = 0
        .ValueAt(1, 0) = 0
        .ValueAt(1, 1) = 1
        .ValueAt(1, 2) = 2
    End With
    
    'Act:
    Dim actual As Matrix
    Set actual = MatMult(B, A)
    
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

'@TestMethod("Function")
Private Sub TestTransposeMatrixObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim B As Matrix
    Set B = CreateMatrix(2, 3)
    With B
        .ValueAt(0, 0) = 1
        .ValueAt(0, 1) = 2
        .ValueAt(0, 2) = 0
        .ValueAt(1, 0) = 0
        .ValueAt(1, 1) = 1
        .ValueAt(1, 2) = 2
    End With
    
    Dim expected As Matrix
    Set expected = CreateMatrix(3, 2)
    With expected
        .ValueAt(0, 0) = 1
        .ValueAt(0, 1) = 0
        .ValueAt(1, 0) = 2
        .ValueAt(1, 1) = 1
        .ValueAt(2, 0) = 0
        .ValueAt(2, 1) = 2
    End With
    
    'Act:
    Dim actual As Matrix
    Set actual = Transpose(B)
    
    'Assert:
    Assert.IsTrue expected.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Function")
Private Sub TestTransposeVectorObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim B As Vector
    Set B = CreateVector(3)
    With B
        .Orientation = ColumnVector
        .ValueAt(0) = 1
        .ValueAt(1) = 2
        .ValueAt(2) = 0
    End With
    
    Dim expected As Vector
    Set expected = CreateVector(3)
    With expected
        .Orientation = RowVector
        .ValueAt(0) = 1
        .ValueAt(1) = 2
        .ValueAt(2) = 0
    End With
    
    'Act:
    Dim actual As IMatrix
    Set actual = Transpose(B)
    
    'Assert:
    If TypeOf actual Is Vector Then
        Dim vActual As Vector
        Set vActual = actual ' cast to Vector object
        Assert.IsTrue expected.Equals(actual)
        Assert.IsTrue vActual.Orientation = RowVector
    Else
        Assert.Fail "Result is not a Vector object"
    End If
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

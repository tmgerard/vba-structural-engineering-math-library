Attribute VB_Name = "TestCholeskySolver"
'@TestModule
'@Folder("Tests.StructuralMath.LinearAlgebra.LinearEqSolver")

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

Private solver As CholeskySolver
Private sysMatrix As Matrix
Private lowMatrix As Matrix
Private sysVector As Vector
Private lowSolution As Vector
Private solution As Vector


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
    
    Dim sysData(15) As Double
    sysData(0) = 4
    sysData(1) = -2
    sysData(2) = 4
    sysData(3) = 2
    sysData(4) = -2
    sysData(5) = 10
    sysData(6) = -2
    sysData(7) = -7
    sysData(8) = 4
    sysData(9) = -2
    sysData(10) = 8
    sysData(11) = 4
    sysData(12) = 2
    sysData(13) = -7
    sysData(14) = 4
    sysData(15) = 7
    
    Set sysMatrix = New Matrix
    sysMatrix.SetSize 4, 4
    sysMatrix.SetData sysData
    
    Dim lowData(15) As Double
    lowData(0) = 2
    lowData(1) = 0
    lowData(2) = 0
    lowData(3) = 0
    lowData(4) = -1
    lowData(5) = 3
    lowData(6) = 0
    lowData(7) = 0
    lowData(8) = 2
    lowData(9) = 0
    lowData(10) = 2
    lowData(11) = 0
    lowData(12) = 1
    lowData(13) = -2
    lowData(14) = 1
    lowData(15) = 1
    
    Set lowMatrix = New Matrix
    lowMatrix.SetSize 4, 4
    lowMatrix.SetData lowData
    
    Set solver = New CholeskySolver
    
    Dim sysVecData(3) As Double
    sysVecData(0) = 20
    sysVecData(1) = -16
    sysVecData(2) = 40
    sysVecData(3) = 28
    
    Set sysVector = New Vector
    sysVector.SetLength 4
    sysVector.SetData sysVecData
    
    Dim lowSolutionData(3) As Double
    lowSolutionData(0) = 10
    lowSolutionData(1) = -2
    lowSolutionData(2) = 10
    lowSolutionData(3) = 4
    
    Set lowSolution = New Vector
    lowSolution.SetLength 4
    lowSolution.SetData lowSolutionData
    
    Dim solutionData(3) As Double
    solutionData(0) = 1
    solutionData(1) = 2
    solutionData(2) = 3
    solutionData(3) = 4
    
    Set solution = New Vector
    solution.SetLength 4
    solution.SetData solutionData
    
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set sysMatrix = Nothing
    Set lowMatrix = Nothing
    Set solver = Nothing
    Set sysVector = Nothing
    Set lowSolution = Nothing
    Set solution = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Algorithm")
Private Sub TestLowerMatrixDecomposition()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Matrix
    Set actual = solver.LowDecomposition(sysMatrix)

    'Assert:
    Assert.IsTrue lowMatrix.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestLowerForwardSubstitution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.ForwardSubstitution(lowMatrix, sysVector)

    'Assert:
    Assert.IsTrue lowSolution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestUpperBackSubstitution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.BackSubstitution(lowMatrix, lowSolution)

    'Assert:
    Assert.IsTrue solution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Algorithm")
Private Sub TestSolution()
    On Error GoTo TestFail
    
    'Arrange:
    

    'Act:
    Dim actual As Vector
    Set actual = solver.Solve(sysMatrix, sysVector)

    'Assert:
    Assert.IsTrue solution.Equals(actual)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Expected Error")
Private Sub TestNonPositiveDefiniteRaisesError()
    On Error GoTo TestFail
    
    Const ExpectedError As Long = CholeskyErrors.Unsolvable
    On Error GoTo TestFail
    Dim badData(3) As Double
    badData(0) = -1: badData(1) = 0: badData(2) = 0: badData(3) = -1
    Dim badMat As Matrix
    Set badMat = New Matrix
    badMat.SetSize 2, 2
    badMat.SetData badData
    solver.LowDecomposition badMat
Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then Resume TestExit Else Resume Assert
End Sub

'@TestMethod("Expected Error")
Private Sub TestNonSquareMatrixRaisesError()
    On Error GoTo TestFail

    Const ExpectedError As Long = CholeskyErrors.NotSquare
    On Error GoTo TestFail
    Dim rectMat As Matrix
    Set rectMat = New Matrix
    rectMat.SetSize 2, 3
    Dim dummyVec As Vector
    Set dummyVec = CreateVector(2)
    solver.Solve rectMat, dummyVec
Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then Resume TestExit Else Resume Assert
End Sub

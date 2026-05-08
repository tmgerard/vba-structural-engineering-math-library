VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CholeskySolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("StructuralMath.LinearAlgebra.LinearEqSolver")
Option Explicit
Implements ILinearEqSolver

Public Enum CholeskyErrors
    BadSize = vbObjectError + 9000
    NotSquare
    Unsolvable
End Enum

Public Function Solve(ByRef mat As IMatrix, ByRef vec As Vector) As Vector

    ' check that matrix and vector work for Cholesky Decomposition
    ValidateSystem mat, vec
    
    Dim lowMat As Matrix
    Set lowMat = LowDecomposition(mat)
    
    Dim lowSolution As Vector
    Set lowSolution = ForwardSubstitution(lowMat, vec)
    
    Set Solve = BackSubstitution(lowMat, lowSolution)
    
End Function

Private Sub ValidateSystem(ByRef mat As IMatrix, ByRef vec As Vector)
    
    If Not mat.Rows = mat.Columns Then ' not a square matrix
        Err.Raise Number:=CholeskyErrors.NotSquare, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="System matrix must be square"
    End If
    
    If Not mat.Columns = vec.Length Then
        Err.Raise Number:=CholeskyErrors.BadSize, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="Size mismatch between matrix and vector."
    End If
       
End Sub

Public Function LowDecomposition(ByRef mat As Matrix) As Matrix

    Dim size As Long
    size = mat.Rows

    Dim decomp As Matrix
    Set decomp = New Matrix
    decomp.SetSize size, size

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim innerSum As Double
    Dim diagonalArg As Double
    
    For i = 0 To size - 1

        ' --- Off-diagonal elements (j < i) ---
        For j = 0 To i - 1
            innerSum = 0
            
            For k = 0 To j - 1
                innerSum = innerSum + decomp.ValueAt(i, k) * decomp.ValueAt(j, k)
            Next k
            
            decomp.ValueAt(i, j) = (mat.ValueAt(i, j) - innerSum) / decomp.ValueAt(j, j)
        Next j

        ' --- Diagonal element (j = i) ---
        innerSum = 0
        For k = 0 To i - 1
            innerSum = innerSum + decomp.ValueAt(i, k) ^ 2
        Next k

        diagonalArg = mat.ValueAt(i, i) - innerSum
        
        If diagonalArg <= 0 Then
            Err.Raise CholeskyErrors.Unsolvable, "Cholesky.LowDecomposition", _
                "Matrix is not positive definite; Cholesky decomposition does not exist."
        End If
        
        decomp.ValueAt(i, i) = Math.Sqr(diagonalArg)

    Next i
    
    Set LowDecomposition = decomp
    
End Function

Public Function ForwardSubstitution(ByRef mat As Matrix, ByRef vec As Vector) As Vector

    Dim size As Long
    size = vec.Length
    
    Dim solution As Vector
    Set solution = New Vector
    Set solution = solution.SetLength(size)
    
    Dim i As Long
    Dim j As Long
    Dim l_ij As Double
    Dim y_j As Double
    Dim b_i As Double
    Dim l_ii As Double
    Dim sum As Double
    Dim solutionValue As Double
    
    For i = 0 To size - 1
    
        sum = 0
        
        For j = 0 To i - 1
        
            l_ij = mat.ValueAt(i, j)
            y_j = solution.ValueAt(j)
            sum = sum + l_ij * y_j
        
        Next j
        
        b_i = vec.ValueAt(i)
        l_ii = mat.ValueAt(i, i)
        solutionValue = (b_i - sum) / l_ii
        solution.ValueAt(i) = solutionValue
        
    Next i
    
    Set ForwardSubstitution = solution

End Function

Public Function BackSubstitution(ByRef mat As Matrix, ByRef vec As Vector) As Vector

    
    Dim size As Long
    size = vec.Length
    
    Dim lastIndex As Long
    lastIndex = size - 1
    
    Dim solution As Vector
    Set solution = New Vector
    Set solution = solution.SetLength(size)
    
    Dim i As Long
    Dim j As Long
    Dim sum As Double
    Dim u_ij As Double
    Dim x_j As Double
    Dim y_i As Double
    Dim u_ii As Double
    Dim solutionValue As Double
    
    For i = lastIndex To 0 Step -1
        
        sum = 0
        
        For j = i + 1 To size - 1
        
            u_ij = mat.ValueAtTransposed(i, j)
            x_j = solution.ValueAt(j)
            sum = sum + u_ij * x_j
        
        Next j
        
        y_i = vec.ValueAt(i)
        u_ii = mat.ValueAtTransposed(i, i)
        solutionValue = (y_i - sum) / u_ii
        solution.ValueAt(i) = solutionValue
        
    Next i
    
    Set BackSubstitution = solution
    
End Function

' ILinearEqSolver Implementation
Private Function ILinearEqSolver_Solve(mat As IMatrix, vec As IMatrix) As Vector
    Set ILinearEqSolver_Solve = Me.Solve(mat, vec)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CholeskySolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StructuralMath.LinearAlgebra.LinearEqSolver")
Option Explicit
Implements ILinearEqSolver

Public Enum CholeskyErrors
    BadSize = vbObjectError + 9000
    NotSquare
    Unsolvable
End Enum

Public Function Solve(ByRef mat As Matrix, ByRef vec As Vector) As Vector

    ' check that matrix and vector work for Cholesky Decomposition
    ValidateSystem mat, vec
    
    Dim lowMat As Matrix
    Set lowMat = LowDecomposition(mat)
    
    Dim lowSolution As Vector
    Set lowSolution = ForwardSubstitution(lowMat, vec)
    
    Set Solve = BackSubstutution(lowMat, lowSolution)
    
End Function

Private Sub ValidateSystem(ByRef mat As Matrix, ByRef vec As Vector)

    If Not mat.Columns = vec.Length Then
        Err.Raise Number:=CholeskyErrors.BadSize, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="Size mismatch between matrix and vector."
    End If
    
    If Not mat.IsSquare Then
        Err.Raise Number:=CholeskyErrors.NotSquare, _
                  Source:="Cholesky.ValidateSystem", _
                  Description:="System matrix must be square"
    End If
End Sub

Public Function LowDecomposition(ByRef mat As Matrix) As Matrix

    On Error GoTo ErrorHandler
    
    Dim size As Long
    size = mat.Rows
    
    Dim decomp As Matrix
    Set decomp = New Matrix
    Set decomp = decomp.SetSize(size, size)
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim sqSum As Double
    Dim m_ij As Double
    Dim diagValue As Double
    Dim nonDiagValue
    Dim nonDiagSum As Double
    Dim l_ik As Double
    Dim l_jk As Double
    Dim l_jj As Double
    
    For i = 0 To size - 1
    
        sqSum = 0
        For j = 0 To i
            
            m_ij = mat.ValueAt(i, j)
            If i = j Then ' value on main diagonal
                diagValue = Math.Sqr(m_ij - sqSum)
                decomp.ValueAt(i, j) = diagValue
            Else ' value under main diagonal
                nonDiagSum = 0
                
                For k = 0 To j - 1
                    
                    l_ik = decomp.ValueAt(i, k)
                    l_jk = decomp.ValueAt(j, k)
                    nonDiagSum = nonDiagSum + l_ik * l_jk
                    
                Next k
                
                l_jj = decomp.ValueAt(j, j)
                nonDiagValue = (m_ij - nonDiagSum) / l_jj
                sqSum = sqSum + nonDiagValue * nonDiagValue
                
                decomp.ValueAt(i, j) = nonDiagValue
   
            End If
            
        Next j
        
    Next i
    
    Set LowDecomposition = decomp
    
    Exit Function
    
ErrorHandler:
    
    Err.Raise CholeskyErrors.Unsolvable, "Cholesky.LowDecomposition", _
        "System of equation is not solvable by Choleski Decomposition."
    
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

Public Function BackSubstutution(ByRef mat As Matrix, ByRef vec As Vector) As Vector

    
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
    
    Set BackSubstutution = solution
    
End Function

' ILinearEqSolver Implementation
Private Function ILinearEqSolver_Solve(mat As Matrix, vec As Vector) As Vector
    Set ILinearEqSolver_Solve = Me.Solve(mat, vec)
End Function

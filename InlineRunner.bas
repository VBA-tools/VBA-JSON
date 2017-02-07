Attribute VB_Name = "InlineRunner"
''
' InlineRunner v1.4.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Runner for outputting results of specs to Immediate window
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Run the given suite
'
' @param {SpecSuite} Specs
' @param {Boolean} [ShowFailureDetails=True] Show failed expectations
' @param {Boolean} [ShowPassed=False] Show passed specs
' @param {Boolean} [ShowSuiteDetails=False] Show details for suite
' --------------------------------------------- '

Public Sub RunSuite(Specs As SpecSuite, Optional ShowFailureDetails As Boolean = True, Optional ShowPassed As Boolean = False, Optional ShowSuiteDetails As Boolean = False)
    Dim SuiteCol As New Collection
    
    SuiteCol.Add Specs
    RunSuites SuiteCol, ShowFailureDetails, ShowPassed, ShowSuiteDetails
End Sub

''
' Run the given collection of spec suites
'
' @param {Collection} of SpecSuite
' @param {Boolean} [ShowFailureDetails=True] Show failed expectations
' @param {Boolean} [ShowPassed=False] Show passed specs
' @param {Boolean} [ShowSuiteDetails=True] Show details for suite
' --------------------------------------------- '

Public Sub RunSuites(SuiteCol As Collection, Optional ShowFailureDetails As Boolean = True, Optional ShowPassed As Boolean = False, Optional ShowSuiteDetails As Boolean = True)
    Dim Suite As SpecSuite
    Dim Spec As SpecDefinition
    Dim TotalCount As Integer
    Dim FailedSpecs As Integer
    Dim PendingSpecs As Integer
    Dim ShowingResults As Boolean
    Dim Indentation As String
    Dim i As Integer
    
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            TotalCount = TotalCount + Suite.SpecsCol.Count
        
            For Each Spec In Suite.SpecsCol
                If Spec.Result = SpecResult.Fail Then
                    FailedSpecs = FailedSpecs + 1
                ElseIf Spec.Result = SpecResult.Pending Then
                    PendingSpecs = PendingSpecs + 1
                End If
            Next Spec
        End If
    Next Suite
    
    Debug.Print vbNewLine & "= " & SummaryMessage(TotalCount, FailedSpecs, PendingSpecs) & " = " & Now & " ========================="
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            If ShowSuiteDetails Then
                Debug.Print SuiteMessage(Suite)
                Indentation = "  "
                ShowingResults = True
            Else
                Indentation = ""
            End If
        
            For Each Spec In Suite.SpecsCol
                If Spec.Result = SpecResult.Fail Then
                    Debug.Print Indentation & FailureMessage(Spec, ShowFailureDetails, Indentation)
                    ShowingResults = True
                ElseIf Spec.Result = SpecResult.Pending Then
                    Debug.Print Indentation & PendingMessage(Spec)
                    ShowingResults = True
                ElseIf ShowPassed Then
                    Debug.Print Indentation & PassingMessage(Spec)
                    ShowingResults = True
                End If
            Next Spec
        End If
    Next Suite
    
    If ShowingResults Then
        Debug.Print "==="
    End If
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Function SummaryMessage(TotalCount As Integer, FailedSpecs As Integer, PendingSpecs As Integer) As String
    If FailedSpecs = 0 Then
        SummaryMessage = "PASS (" & TotalCount - PendingSpecs & " of " & TotalCount & " passed"
    Else
        SummaryMessage = "FAIL (" & FailedSpecs & " of " & TotalCount & " failed"
    End If
    
    If PendingSpecs = 0 Then
        SummaryMessage = SummaryMessage & ")"
    Else
        SummaryMessage = SummaryMessage & ", " & PendingSpecs & " pending)"
    End If
End Function

Private Function FailureMessage(Spec As SpecDefinition, ShowFailureDetails As Boolean, Indentation As String) As String
    Dim FailedExpectation As SpecExpectation
    Dim i As Integer
    
    FailureMessage = ResultMessage(Spec, "X")
    
    If ShowFailureDetails Then
        FailureMessage = FailureMessage & vbNewLine
        
        For Each FailedExpectation In Spec.FailedExpectations
            FailureMessage = FailureMessage & Indentation & "  " & FailedExpectation.FailureMessage
            
            If i + 1 <> Spec.FailedExpectations.Count Then: FailureMessage = FailureMessage & vbNewLine
            i = i + 1
        Next FailedExpectation
    End If
End Function

Private Function PendingMessage(Spec As SpecDefinition) As String
    PendingMessage = ResultMessage(Spec, ".")
End Function

Private Function PassingMessage(Spec As SpecDefinition) As String
    PassingMessage = ResultMessage(Spec, "+")
End Function

Private Function ResultMessage(Spec As SpecDefinition, Symbol As String) As String
    ResultMessage = Symbol & " "
    
    If Spec.Id <> "" Then
        ResultMessage = ResultMessage & Spec.Id & ": "
    End If
    
    ResultMessage = ResultMessage & Spec.Description
End Function

Private Function SuiteMessage(Suite As SpecSuite) As String
    Dim HasFailures As Boolean
    Dim Spec As SpecDefinition
    
    For Each Spec In Suite.SpecsCol
        If Spec.Result = SpecResult.Fail Then
            HasFailures = True
            Exit For
        End If
    Next Spec
    
    If HasFailures Then
        SuiteMessage = "X "
    Else
        SuiteMessage = "+ "
    End If
    
    If Suite.Description <> "" Then
        SuiteMessage = SuiteMessage & Suite.Description
    Else
        SuiteMessage = SuiteMessage & Suite.SpecsCol.Count & " specs"
    End If
End Function

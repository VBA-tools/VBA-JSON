Attribute VB_Name = "DisplayRunner"
''
' DisplayRunner v1.4.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' Runner with sheet output
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Const DefaultSheetName As String = "Tool_Spec_Runner"
Private Const DefaultFilenameRangeName As String = "Filename"
Private Const DefaultOutputStartRow As Integer = 6
Private Const DefaultIdCol As Integer = 1
Private Const DefaultDescCol As Integer = 2
Private Const DefaultResultCol As Integer = 3

Private pFilename As Range
Private pSheet As Worksheet

Private pOutputStartRow As Integer
Private pIdCol As Integer
Private pDescCol As Integer
Private pResultCol As Integer

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Property Get OutputStartRow() As Integer
    If pOutputStartRow <= 0 Then
        pOutputStartRow = DefaultOutputStartRow
    End If
    
    OutputStartRow = pOutputStartRow
End Property

Public Property Let OutputStartRow(value As Integer)
    pOutputStartRow = value
End Property

Public Property Get IdCol() As Integer
    If pIdCol <= 0 Then
        pIdCol = DefaultIdCol
    End If
    
    IdCol = pIdCol
End Property
Public Property Let IdCol(value As Integer)
    pIdCol = value
End Property

Public Property Get DescCol() As Integer
    If pDescCol <= 0 Then
        pDescCol = DefaultDescCol
    End If
    
    DescCol = pDescCol
End Property
Public Property Let DescCol(value As Integer)
    pDescCol = value
End Property

Public Property Get ResultCol() As Integer
    If pResultCol <= 0 Then
        pResultCol = DefaultResultCol
    End If
    
    ResultCol = pResultCol
End Property
Public Property Let ResultCol(value As Integer)
    pResultCol = value
End Property

Public Property Get Filename() As Range
    If pFilename Is Nothing And Not Sheet Is Nothing Then
        Set pFilename = Sheet.Range(DefaultFilenameRangeName)
    End If

    Set Filename = pFilename
End Property
Public Property Set Filename(value As Range)
    Set pFilename = value
End Property

Public Property Get Sheet() As Worksheet
    If pSheet Is Nothing Then
        If SheetExists(DefaultSheetName) Then
            Set pSheet = ThisWorkbook.Sheets(DefaultSheetName)
        Else
            Err.Raise vbObjectError + 1, "DisplayRunner", "Unable to find runner sheet"
        End If
    End If
    Set Sheet = pSheet
End Property

Public Property Set Sheet(value As Worksheet)
    Set pSheet = value
End Property

Public Property Get WBPath() As String
    WBPath = Filename.value
End Property

Public Property Let WBPath(value As String)
    Filename.value = value
End Property

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Run the given suite
'
' @param {SpecSuite} Specs
' @param {Boolean} [Append=False] Append results to existing
' --------------------------------------------- '

Public Sub RunSuite(Specs As SpecSuite, Optional Append As Boolean = False)
    ' Simply add to empty collection and call RunSuites
    Dim SuiteCol As New Collection
    
    SuiteCol.Add Specs
    RunSuites SuiteCol, Append
End Sub

''
' Run the given collection of spec suites
'
' @param {Collection} of SpecSuite
' @param {Boolean} [Append=False] Append results to existing
' --------------------------------------------- '

Public Sub RunSuites(SuiteCol As Collection, Optional Append As Boolean = False)
    Dim Suite As SpecSuite
    Dim Spec As SpecDefinition
    Dim Row As Integer
    Dim Indentation As String
    
    ' 0. Disable screen updating
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ' On Error GoTo Cleanup
    
    ' 1. Clear existing output
    If Not Append Then
        ClearOutput
    End If
    
    ' 2. Loop through Suites and output specs
    Row = NewOutputRow
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            If Suite.Description <> "" Then
                OutputSuiteDetails Suite, Row
                Indentation = "    "
            Else
                Indentation = ""
            End If
        
            For Each Spec In Suite.SpecsCol
                OutputSpec Spec, Row, Indentation
            Next Spec
        End If
    Next Suite
   
Cleanup:

    ' Finally, restore screen updating
    Application.ScreenUpdating = PrevUpdating
    
End Sub

''
' Browse for the workbook to run specs on
' --------------------------------------------- '

Public Sub BrowseForWB()
    Dim BrowseWB As String

    BrowseWB = Application.GetOpenFilename( _
        FileFilter:="Excel Workbooks (*.xls; *.xlsx; *.xlsm), *.xls, *.xlsx, *.xlsm", _
        Title:="Select the Excel Workbook to Test", _
        MultiSelect:=False _
    )

    If BrowseWB <> "" And BrowseWB <> "False" Then
        WBPath = BrowseWB
    End If
End Sub


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Sub OutputSpec(Spec As SpecDefinition, ByRef Row As Integer, Optional Indentation As String = "")
    Sheet.Cells(Row, IdCol) = Spec.Id
    Sheet.Cells(Row, DescCol) = Indentation & Spec.Description
    Sheet.Cells(Row, ResultCol) = Spec.ResultName
    Row = Row + 1
    
    If Spec.FailedExpectations.Count > 0 Then
        Dim Exp As SpecExpectation
        For Each Exp In Spec.FailedExpectations
            Sheet.Cells(Row, DescCol) = Indentation & "X  " & Exp.FailureMessage
            Row = Row + 1
        Next Exp
    End If
End Sub

Private Sub OutputSuiteDetails(Suite As SpecSuite, ByRef Row As Integer)
    Dim HasFailure As Boolean
    Dim Result As String
    Result = "Pass"
    
    For Each Spec In Suite.SpecsCol
        If Spec.Result = SpecResult.Fail Then
            Result = "Fail"
            Exit For
        End If
    Next Spec
    
    Sheet.Cells(Row, DescCol) = Suite.Description
    Sheet.Cells(Row, ResultCol) = Result
    Row = Row + 1
End Sub

Private Sub ClearOutput()
    Dim EndRow As Integer
    
    Dim PrevUpdating As Boolean
    PrevUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    EndRow = NewOutputRow
    If EndRow >= OutputStartRow Then
        Sheet.Range(Cells(OutputStartRow, IdCol), Cells(EndRow, ResultCol)).ClearContents
    End If
    
    Application.ScreenUpdating = PrevUpdating
End Sub

Private Function NewOutputRow() As Integer
    NewOutputRow = OutputStartRow
    
    Do While Sheet.Cells(NewOutputRow, DescCol) <> ""
        NewOutputRow = NewOutputRow + 1
    Loop
End Function

Private Function SheetExists(sheetName As String) As Boolean
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

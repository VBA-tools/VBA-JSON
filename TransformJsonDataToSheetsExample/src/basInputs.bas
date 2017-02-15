Attribute VB_Name = "basInputs"
Option Explicit

Public Sub TransformJsonFile_Click()
On Error GoTo HandleError
    '----------------------------------
    Dim strUrl As String, strJsonObjectWithData As String, strArchiveDirectory As String, strDestinationDirectory As String, strFileNamePrefix As String
    strJsonObjectWithData = mGetNamedRangeInput("Json_Data_Ojbect_Name")
    strArchiveDirectory = mGetNamedRangeInput("JSON_Archive_Directory")
    strDestinationDirectory = mGetNamedRangeInput("Destination_Directory")
    strFileNamePrefix = mGetNamedRangeInput("FileNamePrefix")
    '----------------------------------
    Dim fCloseWorkBook As Boolean, fDelteJsonArchiveFile As Boolean, fAppendDateStampToExcelFilename As Boolean, fNewSheetOnNestedArrayFragment As Boolean
    fCloseWorkBook = mGetNamedRangeInput("chkCloseFileAfterTransform")
    fDelteJsonArchiveFile = mGetNamedRangeInput("chkDeleteJsonFileArchiveDirectory")
    fAppendDateStampToExcelFilename = mGetNamedRangeInput("chkAppendDateStampToExcelFilename")
    fNewSheetOnNestedArrayFragment = mGetNamedRangeInput("chkCreateNewSheetOnNestedFragment")
    '----------------------------------
    If mGetNamedRange("fUseMultipleJsonInput").value Then
        Dim rngMultipleInput As Range
        Dim sht As Worksheet
        Set sht = ThisWorkbook.Worksheets("Multiple JSON Input")
        'Get our useful range
        Dim rngCellsWithValues As Range
        If isSpecialCellValid(sht.UsedRange, xlCellTypeFormulas) Then
            Set rngCellsWithValues = sht.UsedRange.SpecialCells(xlCellTypeFormulas)
            If isSpecialCellValid(sht.UsedRange, xlCellTypeConstants) Then
                Set rngCellsWithValues = Union( _
                    sht.UsedRange.SpecialCells(xlCellTypeConstants), _
                    sht.UsedRange.SpecialCells(xlCellTypeFormulas) _
                )
            End If
        Else
            If isSpecialCellValid(sht.UsedRange, xlCellTypeConstants) Then
                Set rngCellsWithValues = sht.UsedRange.SpecialCells(xlCellTypeConstants)
            End If
        End If
        If isRangeWithCells(rngCellsWithValues) Then
            Set rngMultipleInput = Intersect( _
                sht.Range("A:A"), _
                rngCellsWithValues _
            )
        Else
            Set rngMultipleInput = Intersect( _
                sht.Range("A:A"), _
                sht.UsedRange _
            )
        End If
        Dim rngCell As Range
        
        'Build JSON file for each value in column A
        For Each rngCell In rngMultipleInput.Cells
            strUrl = rngCell.value
            If Len(strUrl) > 0 Then
                ImportJsonFileToWorksheet _
                    strUrl, _
                    strJsonObjectWithData, _
                    strFileNamePrefix, _
                    strArchiveDirectory, _
                    strDestinationDirectory, _
                    fCloseWorkBook, _
                    fDelteJsonArchiveFile, _
                    fAppendDateStampToExcelFilename, _
                    fNewSheetOnNestedArrayFragment
            End If
        Next
    Else
        strUrl = mGetNamedRangeInput("JSON_FileUri")
        ImportJsonFileToWorksheet _
            strUrl, _
            strJsonObjectWithData, _
            strFileNamePrefix, _
            strArchiveDirectory, _
            strDestinationDirectory, _
            fCloseWorkBook, _
            fDelteJsonArchiveFile, _
            fAppendDateStampToExcelFilename, _
            fNewSheetOnNestedArrayFragment
    End If
    
ExitHere:
    Exit Sub
HandleError:
    MsgBox Err.Description, vbCritical, "Transform Json File Error" & Err.Number

    GoTo ExitHere
End Sub

Private Function isSpecialCellValid(rng As Range, varSpecial As XlCellType) As Boolean
On Error Resume Next
    Dim rngTest As Range
    Set rngTest = rng.SpecialCells(varSpecial)
    isSpecialCellValid = Err.Number = 0
    Set rngTest = Nothing
End Function

Private Function isRangeWithCells(ByRef rng As Variant) As Boolean
'only returns true if the range is valid and holds cells
On Error Resume Next
Dim rngTest As Range
    Set rngTest = rng
    Dim lngCellCount As Long
    lngCellCount = rngTest.Cells.Count
    If lngCellCount > 0 Then
        isRangeWithCells = Err.Number = 0
    Else
        isRangeWithCells = False
    End If
    Set rngTest = Nothing
End Function

Private Function mGetNamedRangeInput(strNameIndex As String) As Variant
Dim rngNamed As Range
    Set rngNamed = mGetNamedRange(strNameIndex)
    mGetNamedRangeInput = rngNamed.Offset(0, 1).Cells(1).value
End Function

Public Function mGetNamedRange(strNameIndex As String) As Range
On Error GoTo HandleError
    Dim objNamedRanges As Names
    Set objNamedRanges = ThisWorkbook.Names
    Set mGetNamedRange = objNamedRanges.Item(strNameIndex).RefersToRange
ExitHere:
    Exit Function
HandleError:
    On Error GoTo 0
    Err.Raise 3100, Err.Source, "Named Range Not Found" & " for " + strNameIndex
    GoTo ExitHere
End Function

Sub chkUseMultiple_Click()
    With ActiveSheet.Range("B2").Interior
        If mGetNamedRange("fUseMultipleJsonInput").value Then
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
        Else
            .Pattern = xlNone
        End If
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

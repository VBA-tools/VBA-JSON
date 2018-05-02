Attribute VB_Name = "basInputs"
Option Explicit
'Authored 2015-2018 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
     'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...

Public Sub TransformJsonFile_Click()
On Error GoTo HandleError
    '----------------------------------
    Dim strUrl As String, strArchiveDirectory As String, strDestinationDirectory As String, strFileNamePrefix As String
    strArchiveDirectory = GetNamedRangeInput("JSON_Archive_Directory")
    strDestinationDirectory = GetNamedRangeInput("Destination_Directory")
    strFileNamePrefix = GetNamedRangeInput("FileNamePrefix")
    '----------------------------------
    Dim fCloseWorkBook As Boolean, fDelteJsonArchiveFile As Boolean, fAppendDateStampToExcelFilename As Boolean, fNewSheetOnNestedArrayFragment As Boolean
    fCloseWorkBook = GetNamedRangeInput("chkCloseFileAfterTransform")
    fDelteJsonArchiveFile = GetNamedRangeInput("chkDeleteJsonFileArchiveDirectory")
    fAppendDateStampToExcelFilename = GetNamedRangeInput("chkAppendDateStampToExcelFilename")
    fNewSheetOnNestedArrayFragment = GetNamedRangeInput("chkCreateNewSheetOnNestedFragment")
    '----------------------------------
    If GetNamedRange("fUseMultipleJsonInput").value Then
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
                    GetNamedRangeInput("Json_Data_Ojbect_Name"), _
                    strFileNamePrefix, _
                    strArchiveDirectory, _
                    strDestinationDirectory, _
                    fCloseWorkBook, _
                    fDelteJsonArchiveFile, _
                    fAppendDateStampToExcelFilename, _
                    fNewSheetOnNestedArrayFragment
            End If
        Next
    ElseIf GetNamedRangeInput("JSON_FileUri") = True Then
        'Crawl directory and Sub Directories for all .json files
        Dim fso As Object: Set fso = CreateObject("system.FileSystemObject") 'New FileSystemObject
        Dim objFolder As Object: Set objFolder = fso.GetFolder(GetNamedRange("JsonFileUrl").value)
        CrawlAndProcessFolders _
            objFolder, _
            GetNamedRangeInput("Json_Data_Ojbect_Name"), _
            strFileNamePrefix, _
            strArchiveDirectory, _
            strDestinationDirectory, _
            fCloseWorkBook, _
            fDelteJsonArchiveFile, _
            fAppendDateStampToExcelFilename, _
            fNewSheetOnNestedArrayFragment
    Else
        strUrl = GetNamedRangeInput("JSON_FileUri")
        ImportJsonFileToWorksheet _
            strUrl, _
            GetNamedRangeInput("Json_Data_Ojbect_Name"), _
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

Private Sub CrawlAndProcessFolders(objFolder As Object, _
    Optional ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strFileNamePrefix As String, _
    Optional ByRef strJsonArchiveDirectory As String, _
    Optional ByRef strExcelFileSaveDirectory As String, _
    Optional fCloseWorkBook As Boolean = False, _
    Optional fDelteJsonArchiveFile As Boolean = False, _
    Optional fAppendDateStampToExcelFilename As Boolean = True, _
    Optional fNewSheetOnNestedArrayFragment As Boolean = False _
)
        Dim fso As Object: Set fso = CreateObject("system.FileSystemObject") 'New FileSystemObject
        Dim objFile As Object
        For Each objFile In objFolder.Files
            If LCase(Right(objFile.path, 5)) = ".json" Then
            ImportJsonFileToWorksheet _
                objFile.path, _
                GetNamedRangeInput("Json_Data_Ojbect_Name"), _
                strFileNamePrefix, _
                strArchiveDirectory, _
                strDestinationDirectory, _
                fCloseWorkBook, _
                fDelteJsonArchiveFile, _
                fAppendDateStampToExcelFilename, _
                fNewSheetOnNestedArrayFragment
            End If
        Next
        Dim objSubFolder As Object
        For Each objSubFolder In objFolder.SubFolders
            CrawlAndProcessFolders _
                objSubFolder, _
                GetNamedRangeInput("Json_Data_Ojbect_Name"), _
                strFileNamePrefix, _
                strArchiveDirectory, _
                strDestinationDirectory, _
                fCloseWorkBook, _
                fDelteJsonArchiveFile, _
                fAppendDateStampToExcelFilename, _
                fNewSheetOnNestedArrayFragment
        Next
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

Private Function GetNamedRangeInput(strNameIndex As String) As Variant
Dim rngNamed As Range
    Set rngNamed = GetNamedRange(strNameIndex)
    GetNamedRangeInput = rngNamed.Offset(0, 1).Cells(1).value
End Function

Public Function GetNamedRange(strNameIndex As String) As Range
On Error GoTo HandleError
    Dim objNamedRanges As Names
    Set objNamedRanges = ThisWorkbook.Names
    Set GetNamedRange = objNamedRanges.Item(strNameIndex).RefersToRange
ExitHere:
    Exit Function
HandleError:
    On Error GoTo 0
    Err.Raise 3100, Err.Source, "Named Range Not Found" & " for " + strNameIndex
    GoTo ExitHere
End Function

Sub chkUseMultiple_Click()
    With ActiveSheet.Range("B2").Interior
        If GetNamedRange("fUseMultipleJsonInput").value = True Then
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            GetNamedRange("CheckCrawlDirectoryLink").value = False
            GetNamedRange("CheckCrawlDirectoryLink").Worksheet.Shapes("chkCrawlDirectories").Visible = False
        Else
            .Pattern = xlNone
            GetNamedRange("CheckCrawlDirectoryLink").Worksheet.Shapes("chkCrawlDirectories").Visible = True
            'Trigger Worksheet_Change code
            GetNamedRange("JsonFileUrl").value = GetNamedRange("JsonFileUrl").value
        End If
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

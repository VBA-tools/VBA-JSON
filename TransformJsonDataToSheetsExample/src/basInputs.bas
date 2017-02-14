Attribute VB_Name = "basInputs"
Option Explicit

Public Sub TransformJsonFile_Click()
On Error GoTo HandleError
    '----------------------------------
    Dim strUrl As String, strJsonObjectWithData As String, strArchiveDirectory As String, strDestinationDirectory As String, strFileNamePrefix As String
    strUrl = mGetNamedRangeInput("JSON_FileUri")
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
    
ExitHere:
    Exit Sub
HandleError:
    MsgBox Err.Description, vbCritical, "Transform Json File Error" & Err.Number

    GoTo ExitHere
End Sub

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

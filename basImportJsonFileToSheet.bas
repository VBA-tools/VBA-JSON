Attribute VB_Name = "basImportJsonFileToSheet"
Option Explicit
'Authored 2017 (5 hours) by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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
Private mStrRawData As String
Private mlngCurrentDupSheetCount As Long

Public Sub ImportJsonFileDailyToWorksheet( _
    ByRef strUrl As String, _
    ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strDesinationFileName As String, _
    Optional ByRef strJsonArchiveDirectory As String, _
    Optional ByRef strSolutionFileArchiveDirectory As String _
)
On Error GoTo ExitHere
'Use when the data posted to the web is only updated daily, we check to see if we have data for that day and only proceed after asking
Application.ScreenUpdating = False
Dim strDesinationWorkbookFileName As String
    If Len(strDesinationFileName) > 0 Then
        strDesinationWorkbookFileName = strDesinationFileName
    Else
        strDesinationWorkbookFileName = strURLDecode(Left(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/")), 50 - 6) & Format(Now(), "yymmdd"))
    End If
    strDesinationWorkbookFileName = RemoveForbiddenFilenameCharacters(strDesinationWorkbookFileName)
    Dim strTempDownloadFile As String
    strTempDownloadFile = DownloadUrlFileToTemp(strUrl, "json", strJsonArchiveDirectory)
    mStrRawData = vbNullString
    mlngCurrentDupSheetCount = 1
    ExpandJsonToNewWorkbook strTempDownloadFile, strJSONObjectNameWithData, strDesinationWorkbookFileName, strSolutionFileArchiveDirectory
ExitHere:
    Application.ScreenUpdating = True
    'Delete our temp JSON file if we are done with it
End Sub

Sub ExpandJsonToNewWorkbook( _
    strJsonFilePath As String, _
    strJSONObjectNameWithData As String, _
    Optional strDesinationWorkbookFileName As String, _
    Optional strSolutionDestinationDirectory _
)
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
Dim JsonTS As Object ' TextStream
Dim JsonText As String
Dim Parsed As Dictionary

    ' Read .json file
    Set JsonTS = fso.OpenTextFile(strJsonFilePath, ForReading)
    If JsonTS.AtEndOfStream Then
        JsonText = vbNullString
    Else
        JsonText = JsonTS.ReadAll
    End If
    JsonTS.Close
    Set JsonTS = Nothing
    Set fso = Nothing
    Dim jsonData As Dictionary
    Set jsonData = ParseJson(JsonText)
    Dim wkb As Workbook
    Set wkb = Application.Workbooks.Add()
    Dim wsh As Worksheet
    Set wsh = wkb.Sheets(1)
    wsh.Name = "JSON_Object"
    '----------------------------------------------
    GetAllJsonObjectNestedValues jsonData, wkb, wsh
    '----------------------------------------------
    'Cleanup
    For Each wsh In wkb.Sheets
        If wsh.UsedRange.Cells.Count = 1 Then
            wsh.Activate
            DeleteSheet wsh.Name, wkb
        End If
    Next
    If Len(strSolutionDestinationDirectory) = 0 Then 'and folder exists and we ca write to it...
        strSolutionDestinationDirectory = ThisWorkbook.Path
    Else
        strSolutionDestinationDirectory = GetRelativePathViaParent(strSolutionDestinationDirectory)
    End If
    wkb.SaveAs strSolutionDestinationDirectory & "\" & strDesinationWorkbookFileName, XlFileFormat.xlExcel8
End Sub

Private Function GetAllJsonObjectNestedValues( _
            ByRef dict As Dictionary, _
            ByRef wkb As Workbook, _
            ByRef wsh As Worksheet, _
            Optional strPreviousObjectKey As String) _
As Variant
'This method is overly optimistic that each object will hold data, we will create a sheet even for empty objects, if a
'sheet is found to have no data we delete it, a possibly faster/less memory intensive way would be to run through the json file
'once to determine what objects hold no data first, and ignore them
'-------------------
'additionally there is no direct relationship displayed from the nested object to it's parent using this method, this will
'have to be built out and incorperated to properly import into a relational database if that data is needed
If dict.Count > 0 Then
    Dim aryTemp As Variant
    ReDim aryTemp(0 To 1, 0 To dict.Count - 1) '2 columns, with n rows
    Dim lngDataRow As Variant
    lngDataRow = 1
    Dim lngItem As LongPtr
    If Not wsh Is Nothing Then
        wsh.Activate
    End If
        For lngItem = 0 To dict.Count - 1
            Dim varKey As Variant
            Dim strKeyName As String
            Dim varItem As Variant
            'Dim sheetName As String
            varKey = dict.Keys()(lngItem)
            strKeyName = CStr(varKey)
            If IsObject(dict.Items()(lngItem)) Then
                Dim wshNew As Worksheet
                Dim sheetName As String
                If Len(strPreviousObjectKey) = 0 Then
                    sheetName = Left(varKey, 28)
                Else
                    sheetName = Left(strPreviousObjectKey & "_" & varKey, 28)
                End If
                If SheetExists(sheetName, wkb) Then
                    mlngCurrentDupSheetCount = mlngCurrentDupSheetCount + 1
                    Set wshNew = CreateWorksheet(sheetName & mlngCurrentDupSheetCount, wkb:=wkb)
                Else
                    Set wshNew = CreateWorksheet(sheetName, wkb:=wkb)
                End If
                varItem = "-------------------Object-------------------"
                Dim tmpDictionary As Dictionary
                Dim aryRecursive As Variant
                If TypeName(dict.Items()(lngItem)) = "Dictionary" Then
                    Dim objItemDictionary As Dictionary
                    Set objItemDictionary = dict.Items()(lngItem)
                    aryRecursive = GetAllJsonObjectNestedValues(objItemDictionary, wkb, wshNew, strKeyName)
                Else
                    Dim objItem As Collection
                    Set objItem = dict.Items()(lngItem)
                    Select Case TypeName(objItem)
                        Case "Collection"
                            Dim objItemElement As Variant
                            Dim lngItemElementCounter As Long
                            lngItemElementCounter = 0
                            For Each objItemElement In objItem
                                lngItemElementCounter = lngItemElementCounter + 1
                                If TypeName(objItemElement) = "Dictionary" Then
                                    Set tmpDictionary = objItemElement
                                    aryRecursive = GetAllJsonObjectNestedValues(tmpDictionary, wkb, wshNew, strKeyName)
                                Else
                                    wsh.Activate
                                    If lngItemElementCounter = 1 Then 'Add Column Header
                                        wsh.Range(Cells(lngItemElementCounter, lngItem), Cells(lngItemElementCounter, lngItem)).value = varKey
                                        wsh.Range(Cells(lngItemElementCounter + 1, lngItem), Cells(lngItemElementCounter + 1, lngItem)).value = objItemElement
                                    Else
                                        wsh.Range(Cells(lngItemElementCounter + 1, lngItem), Cells(lngItemElementCounter + 1, lngItem)).value = objItemElement
                                    End If
                                End If
                            Next objItemElement
                            Debug.Print lngItemElementCounter
                        Case "Dictionary"
                            Set tmpDictionary = objItem
                            aryRecursive = GetAllJsonObjectNestedValues(objItem, wkb, wsh) 'Executing this debug prints for testing untill we decide how to export this data to a spreedsheet appropriately
                    End Select
                End If
            Else
                varItem = dict.Items()(lngItem)
                wsh.Activate
                If strKeyName = CStr(wsh.Range(Cells(1, 1), Cells(1, 1)).value) Then
                    lngDataRow = wsh.UsedRange.Rows.Count 'Assuming the same column header value = new row in any key value pair list
                End If
                If lngDataRow = 1 Then
                    wsh.Range(Cells(lngDataRow, lngItem + 1), Cells(lngDataRow, lngItem + 1)).value = varKey
                End If
                wsh.Range(Cells(lngDataRow + 1, lngItem + 1), Cells(lngDataRow + 1, lngItem + 1)).value = varItem
            End If
        Next lngItem
    End If
End Function



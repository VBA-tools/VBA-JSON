Attribute VB_Name = "basImportJsonFileToSheet"
Option Explicit
'Authored 2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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
Private mlngCurrentDupSheetCount As Long

Private Enum uriType
    uriFile = 1
    uriDirectory = 2
    uriHttp = 3
    uriUndefined = 4
End Enum

Public Sub ImportJsonFileToWorksheet( _
    ByRef strUrl As String, _
    Optional ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strFileNamePrefix As String, _
    Optional ByRef strJsonArchiveDirectory As String, _
    Optional ByRef strExcelFileSaveDirectory As String, _
    Optional fCloseWorkBook As Boolean = False, _
    Optional fDelteJsonArchiveFile As Boolean = False, _
    Optional fAppendDateStampToExcelFilename = True _
)
On Error GoTo ExitHere
'Use when the data posted to the web is only updated daily, we check to see if we have data for that day and only proceed after asking
Application.ScreenUpdating = False

'We expect either a Uri or a file, for now check File Exists and then assume it's a Uri if it doesn't
'a more complete method is something like: http://stackoverflow.com/questions/9724779/vba-identifying-whether-a-string-is-a-file-a-Directory-or-a-web-Uri
Dim uriFileType As uriType
uriFileType = mCheckPath(strUrl)
Dim strDesinationWorkbookFileName As String
    Select Case uriFileType
        Case uriFile
            
        Case uriDirectory
            
        Case uriHttp
            If Len(strFileNamePrefix) > 0 Then
                strDesinationWorkbookFileName = Left(strFileNamePrefix, 44)
            Else
                strDesinationWorkbookFileName = strUrlDecode(Left(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/")), 44))
            End If
            If fAppendDateStampToExcelFilename Then
                strDesinationWorkbookFileName = strDesinationWorkbookFileName & Format(Now(), "yymmddhhss") & Right(Timer, 2)
            End If
            strDesinationWorkbookFileName = RemoveForbiddenFilenameCharacters(strDesinationWorkbookFileName)
            Dim strTempDownloadFile As String
            strTempDownloadFile = DownloadUriFileToTemp(strUrl, "json", strJsonArchiveDirectory)
        Case uriUndefined
            MsgBox "", vbOKOnly, "Transform Json File From Web"
            Exit Sub
    End Select
    mlngCurrentDupSheetCount = 1
    '---------------------------------
    ExpandJsonToNewWorkbook _
        strTempDownloadFile, _
        strJSONObjectNameWithData, _
        strDesinationWorkbookFileName, _
        strExcelFileSaveDirectory, _
        fCloseWorkBook
    If fDelteJsonArchiveFile Then
        Kill strTempDownloadFile
    End If
ExitHere:
    Application.ScreenUpdating = True
    'Delete our temp JSON file if we are done with it
End Sub

Sub ExpandJsonToNewWorkbook( _
    strJsonFilePath As String, _
    Optional strJSONObjectNameWithData As String, _
    Optional strDesinationWorkbookFileName As String, _
    Optional strSolutionDestinationDirectory, _
    Optional fCloseWorkBook As Boolean = False _
)
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
Dim JsonTS As Object ' TextStream
Dim JsonText As String
Dim Parsed As Dictionary
    '----------------------------------------------
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
    '----------------------------------------------
    'Parse JSON
    Dim jsonData As Variant: Set jsonData = ParseJson(JsonText)
    Dim wkb As Workbook: Set wkb = Application.Workbooks.Add()
    Dim wsh As Worksheet: Set wsh = wkb.Sheets(1)
    '----------------------------------------------
    'Rename first worksheet
    'The root parsed JSON text will be a Dictionary, Collection or String
    'see https://tools.ietf.org/html/rfc7159#section-2
    Select Case TypeName(jsonData)
        Case "Dictionary"
            wsh.Name = "JSON_object"
        Case "Collection"
            wsh.Name = "JSON_array"
        Case "String"
            wsh.Name = "JSON_name"
    End Select
        
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
    If Len(strSolutionDestinationDirectory) = 0 Then 'and Directory exists and we ca write to it...
        strSolutionDestinationDirectory = ThisWorkbook.path
    Else
        strSolutionDestinationDirectory = GetRelativePathViaParent(strSolutionDestinationDirectory)
    End If
    wkb.SaveAs strSolutionDestinationDirectory & "\" & strDesinationWorkbookFileName, XlFileFormat.xlExcel8
    '----------------------------------------------
    'Final option actions
    If fCloseWorkBook Then
        wkb.Close
    End If
End Sub

Private Function GetAllJsonObjectNestedValues( _
            ByRef objJson As Variant, _
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
'TODO need to read https://tools.ietf.org/html/rfc7159#section-2
'and clarify the object object relation ships found in the spec here

    wsh.Activate
    Select Case TypeName(objJson)
        Case "Dictionary"
            If objJson.Count > 0 Then
                Dim lngDataRow As Variant
                lngDataRow = 1
                Dim lngItem As Long
                For lngItem = 0 To objJson.Count - 1
                    Dim varKey As Variant
                    Dim strKeyName As String
                    Dim varItem As Variant
                    'Dim sheetName As String
                    varKey = objJson.Keys()(lngItem)
                    strKeyName = CStr(varKey)
                    'Obect item is an object...
                    If IsObject(objJson.Items()(lngItem)) Then
                        Dim sheetName As String
                        If Len(strPreviousObjectKey) = 0 Then
                            sheetName = Left(varKey, 28)
                        Else
                            sheetName = Left(strPreviousObjectKey & "_" & varKey, 28)
                        End If
                        varItem = "-------------------Object-------------------"
                        Dim tmpDictionary As Dictionary
                        Dim objRecursive As Variant
                        Select Case TypeName(objJson.Items()(lngItem))
                            Case "Dictionary"
                                Dim objItemDictionary As Dictionary
                                Set objItemDictionary = objJson.Items()(lngItem)
                                objRecursive = GetAllJsonObjectNestedValues(objItemDictionary, wkb, mNewDictionaryWorkSheet(sheetName, wkb), strKeyName)
                            Case "Collection"
                                Dim objItem As Variant
                                Set objItem = objJson.Items()(lngItem)
                                Select Case TypeName(objItem)
                                    Case "Dictionary"
                                        Set tmpDictionary = objItem
                                        objRecursive = GetAllJsonObjectNestedValues(objItem, wkb, mNewDictionaryWorkSheet(sheetName, wkb), strKeyName)   'Executing this debug prints for testing untill we decide how to export this data to a spreedsheet appropriately
                                    Case "Collection"
                                        mWriteCollectionToSheet _
                                            objItem, _
                                            tmpDictionary, _
                                            wkb, _
                                            sheetName, _
                                            strKeyName, _
                                            objRecursive, _
                                            wsh, _
                                            varKey, _
                                            lngItem
                                End Select
                            Case Else
                                varItem = objJson.Items()(lngItem)
                                mWriteElementToTable _
                                    varItem, _
                                    wsh, _
                                    strKeyName, _
                                    lngDataRow, _
                                    lngItem, _
                                    varKey
                        End Select
                    Else 'must be a Number, String, Boolean, or null
                        varItem = objJson.Items()(lngItem)
                        mWriteElementToTable _
                            varItem, _
                            wsh, _
                            strKeyName, _
                            lngDataRow, _
                            lngItem, _
                            varKey
                    End If
                Next lngItem
            End If
        Case "Collection"
            Dim lngCollectionItem As Long
           ' For lngCollectionItem = 0 To objJson.Count - 1
                mWriteCollectionToSheet _
                    objJson, _
                    tmpDictionary, _
                    wkb, _
                    wsh.Name, _
                    strKeyName, _
                    objJson, _
                    wsh, _
                    varKey, _
                    lngCollectionItem
            'Next lngCollectionItem
        Case Else 'must be a Number, String, Boolean, or null
            varItem = objJson.Items()(lngCollectionItem)
            mWriteElementToTable _
                varItem, _
                wsh, _
                strKeyName, _
                lngDataRow, _
                lngItem, _
                varKey
    End Select
End Function

Private Function mNewDictionaryWorkSheet(sheetName As String, wkb As Workbook) As Worksheet
    If SheetExists(sheetName, wkb) Then
        mlngCurrentDupSheetCount = mlngCurrentDupSheetCount + 1
        Set mNewDictionaryWorkSheet = CreateWorksheet(sheetName & mlngCurrentDupSheetCount, wkb:=wkb)
    Else
        Set mNewDictionaryWorkSheet = CreateWorksheet(sheetName, wkb:=wkb)
    End If
    
End Function

Private Sub mWriteCollectionToSheet( _
    objItem As Variant, _
    tmpDictionary As Dictionary, _
    wkb As Workbook, _
    strNewSheetName As String, _
    strKeyName As String, _
    objRecursive As Variant, _
    wsh As Worksheet, _
    varKey As Variant, _
    lngItem As Long)
Dim objItemElement As Variant
Dim lngItemElementCounter As Long
    lngItemElementCounter = 0
    For Each objItemElement In objItem
        lngItemElementCounter = lngItemElementCounter + 1
        If TypeName(objItemElement) = "Dictionary" Then
            Set tmpDictionary = objItemElement
            objRecursive = GetAllJsonObjectNestedValues(tmpDictionary, wkb, wsh, strKeyName)
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
End Sub

Private Sub mWriteElementToTable( _
    varItem As Variant, _
    wsh As Worksheet, _
    strKeyName, _
    lngDataRow, _
    lngItem, _
    varKey _
)
    wsh.Activate
    If strKeyName = CStr(wsh.Range(Cells(1, 1), Cells(1, 1)).value) Then
        lngDataRow = wsh.UsedRange.Rows.Count 'Assuming the same column header value = new row in any key value pair list
    End If
    If lngDataRow = 1 Then
        wsh.Range(Cells(lngDataRow, lngItem + 1), Cells(lngDataRow, lngItem + 1)).value = varKey
    End If
    wsh.Range(Cells(lngDataRow + 1, lngItem + 1), Cells(lngDataRow + 1, lngItem + 1)).value = varItem
End Sub

'********************************************************************
'Next 4 functions Adapted from http://stackoverflow.com/questions/9724779/vba-identifying-whether-a-string-is-a-file-a-Directory-or-a-web-url
'posted March 2016 all Code contributions to stackoverflow after Feb 2016 are under the MIT license:https://opensource.org/licenses/MIT
'********************************************************************
'Some of the fuctions have been Adapted from: http://allenbrowne.com
'On the botom of http://allenbrowne.com/tips.html
'********************************************************************
'Permission
'You may freely use anything (code, forms, algorithms, ...) from these articles and sample databases for any purpose (personal, educational, commercial, resale, ...). All we ask is that you acknowledge this website in your code, with comments such as:
'Source: http://allenbrowne.com
'********************************************************************
Private Function mCheckPath(path) As uriType
    Dim retval
    retval = uriUndefined
    If (retval = uriUndefined) And mHttpExists(path) Then
        retval = uriHttp
    End If
    If (retval = uriUndefined) And mFileExists(path) Then
        retval = uriFile
    End If
    If (retval = uriUndefined) And mDirectoryExists(path) Then
        retval = uriDirectory
    End If
    mCheckPath = retval
End Function

Private Function mFileExists(ByVal strFile As String, Optional bFindDirectories As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindDirectories. If strFile is a Directory, mFileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindDirectories Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include Directories as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the Directory.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
    'If Dir() returns something, the file exists.
    On Error Resume Next
    mFileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Private Function mDirectoryExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    mDirectoryExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Private Function mHttpExists(ByVal sURL As String) As Boolean
    'TODO have not built out how to query that an FTP file is present, for FTP responces see: https://tools.ietf.org/html/rfc959
    On Error GoTo HandleError
    Dim oXHTTP As Object
    On Error Resume Next
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Set oXHTTP = CreateObject("MSXML2.SERVERXMLHTTP")
    End If
    If Not UCase(sURL) Like "HTTP:*" _
        And Not UCase(sURL) Like "HTTPS:*" _
    Then
        sURL = "https://" & sURL
    End If
    oXHTTP.Open "HEAD", sURL, False
    oXHTTP.send
    Debug.Print oXHTTP.Status
    Select Case oXHTTP.Status
        Case 200 To 399, 403, 426 ' maybe 100 and 101?
            'The 2xx (Successful) class of status code indicates that the client's
            'request was successfully received, understood, and accepted. https://tools.ietf.org/html/rfc7231#section-6.3
            '403 status code indicates that the server understood the request but refuses to authorize it
            '426 tells us it's here but you need to upgrade current protocol
            mHttpExists = True
        Case Else '400, 404,410 500's
            mHttpExists = False
    End Select
    Exit Function
HandleError:
    Debug.Print Err.Description
    mHttpExists = False
End Function

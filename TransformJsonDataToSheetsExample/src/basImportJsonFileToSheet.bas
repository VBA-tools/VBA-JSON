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
Private objParsedJson As Variant
Private Enum uriType
    uriFile = 1
    uriDirectory = 2
    uriHttp = 3
    uriUndefined = 4
    uriParentDirectoryExists = 5
End Enum

Public Sub ImportJsonFileToWorksheet( _
    ByRef strUrl As String, _
    Optional ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strFileNamePrefix As String, _
    Optional ByRef strJsonArchiveDirectory As String, _
    Optional ByRef strExcelFileSaveDirectory As String, _
    Optional fCloseWorkBook As Boolean = False, _
    Optional fDelteJsonArchiveFile As Boolean = False, _
    Optional fAppendDateStampToExcelFilename As Boolean = True, _
    Optional fNewSheetOnNestedArrayFragment As Boolean = False _
)
On Error GoTo ExitHere
'Use when the data posted to the web is only updated daily, we check to see if we have data for that day and only proceed after asking
Application.ScreenUpdating = False
'We expect either a http url or a file,
'TODO add support for FTP?
Dim strDesinationWorkbookFileName As String
Dim strJsonSourceFilePath As String
Dim uriFileType As uriType: uriFileType = mCheckPath(strUrl)
    Select Case uriFileType
        Case uriFile
            If Len(strFileNamePrefix) > 0 Then
                strDesinationWorkbookFileName = Left(strFileNamePrefix, 44)
            Else
                strDesinationWorkbookFileName = Left(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "\")), 44)
                strDesinationWorkbookFileName = Left(strDesinationWorkbookFileName, InStrRev(strDesinationWorkbookFileName, ".") - 1) 'remove extension from JSON file for Prefix of excel file
            End If
            If fAppendDateStampToExcelFilename Then
                strDesinationWorkbookFileName = strDesinationWorkbookFileName & Format(Now(), "yymmddhhss") & Right(Timer, 2)
            End If
            strDesinationWorkbookFileName = RemoveForbiddenFilenameCharacters(Replace(strDesinationWorkbookFileName, ".", "_"))
            strJsonSourceFilePath = GetRelativePathViaParent(strUrl)
        Case uriHttp
            If Len(strFileNamePrefix) > 0 Then
                strDesinationWorkbookFileName = Left(strFileNamePrefix, 44)
            Else
                strDesinationWorkbookFileName = strUrlDecode(Left(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/")), 44))
            End If
            If fAppendDateStampToExcelFilename Then
                strDesinationWorkbookFileName = strDesinationWorkbookFileName & Format(Now(), "yymmddhhss") & Right(Timer, 2)
            End If
            strDesinationWorkbookFileName = RemoveForbiddenFilenameCharacters(Replace(strDesinationWorkbookFileName, ".", "_"))
            strJsonSourceFilePath = DownloadUriFileToTemp(strUrl, "json", strJsonArchiveDirectory)
        Case uriDirectory
            MsgBox "Input is a directory, only JSON http(s) url or direct filepath are supported", vbOKOnly, "Transform Json File From Web"
            Exit Sub
        Case uriParentDirectoryExists
            MsgBox "Unable to find file: " & strJSONObjectNameWithData & vbCrLf & " The files Parent Directory does exist", vbOKOnly, "Transform Json File From Web"
            Exit Sub
        Case uriUndefined
            MsgBox "Only JSON http(s) url or filepath are supported", vbOKOnly, "Transform Json File From Web"
            Exit Sub
    End Select
    mlngCurrentDupSheetCount = 1
    '---------------------------------
    ExpandJsonToNewWorkbook _
        strJsonSourceFilePath, _
        strJSONObjectNameWithData, _
        strDesinationWorkbookFileName, _
        strExcelFileSaveDirectory, _
        fCloseWorkBook, _
        fNewSheetOnNestedArrayFragment
    If fDelteJsonArchiveFile Then
        Kill strJsonSourceFilePath
    End If
ExitHere:
    Application.ScreenUpdating = True
    'Delete our temp JSON file if we are done with it
End Sub

Sub ExpandJsonToNewWorkbook( _
    ByRef strJsonFilePath As String, _
    Optional ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strDesinationWorkbookFileName As String, _
    Optional ByRef strSolutionDestinationDirectory As String, _
    Optional fCloseWorkBook As Boolean = False, _
    Optional fNewSheetOnNestedArrayFragment As Boolean = False _
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
    Dim jsonData As Variant
    Set objParsedJson = ParseJson(JsonText)
    Set jsonData = objParsedJson
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
    GetAllJsonObjectNestedValues jsonData, wkb, wsh, wsh.Name, fNewSheetOnNestedArrayFragment
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
    Optional ByRef strPreviousObjectKey As String, _
    Optional ByRef fNewSheetOnNestedArrayFragment As Boolean = False _
) _
As Variant
'This method is overly optimistic that each object will hold data, we will create a sheet even for empty objects, if a
'sheet is found to have no data we delete it, a possibly faster/less memory intensive way would be to run through the json file
'once to determine what objects hold no data first, and ignore them, we may need to do this any way to determine how best to represent the data into tables...
'-------------------
'additionally there is no direct relationship displayed from the nested object to it's parent using this method, this will
'have to be built out and incorperated to properly import into a relational database if that data is needed
'TODO need to read https://tools.ietf.org/html/rfc7159#section-2
'and clarify the object object relation ships found in the spec here
    wsh.Activate
    Select Case TypeName(objJson)
        Case "Dictionary"
            If objJson.Count > 0 Then
                Dim varDataRow As Variant
                varDataRow = 1
                Dim lngItem As Long
                For lngItem = 0 To objJson.Count - 1
                    Dim varKey As Variant
                    Dim strKeyName As String
                    Dim varItem As Variant
                    'Dim sheetName As String
                    varKey = objJson.Keys()(lngItem)
                    strKeyName = CStr(varKey)
                    Dim sheetName As String
                    If Len(strPreviousObjectKey) = 0 Then
                        sheetName = Left(varKey, 28)
                    Else
                        sheetName = Left(strPreviousObjectKey & "_" & varKey, 28)
                    End If
                    'Obect item is an object...
                    If IsObject(objJson.Items()(lngItem)) Then
                        varItem = "-------------------Object-------------------"
                        Dim objRecursive As Variant
                        Dim objItem As Variant
                        Set objItem = objJson.Items()(lngItem)
                        Select Case TypeName(objItem)
                            Case "Dictionary"
                                Dim wshDictionary As Worksheet 'If fNewSheetOnNestedArrayFragment Then
                                Set wshDictionary = mCreateWorkSheet(strPreviousObjectKey, wkb)
                                'End If
                                objRecursive = GetAllJsonObjectNestedValues( _
                                    objItem, _
                                    wkb, _
                                    wshDictionary, _
                                    strKeyName)
                            Case "Collection"
                                Dim wshCollection As Worksheet 'If fNewSheetOnNestedArrayFragment Then
                                Set wshCollection = mCreateWorkSheet(strPreviousObjectKey, wkb)
                                    mWriteCollectionToSheet _
                                        objItem, _
                                        wkb, _
                                        sheetName, _
                                        strKeyName, _
                                        objRecursive, _
                                        wsh, _
                                        varKey, _
                                        lngItem
                            Case Else
                                varItem = objJson.Items()(lngItem)
                                mWriteElementToTable _
                                    varItem, _
                                    wsh, _
                                    strKeyName, _
                                    varDataRow, _
                                    lngItem, _
                                    varKey
                        End Select
                    Else 'must be a Number, String, Boolean, or null
                        varItem = objJson.Items()(lngItem)
                        mWriteElementToTable _
                            varItem, _
                            wsh, _
                            strKeyName, _
                            varDataRow, _
                            lngItem, _
                            varKey
                    End If
                Next lngItem
            End If
        Case "Collection"
                mWriteCollectionToSheet _
                    objJson, _
                    wkb, _
                    "Json_array", _
                    "Json_array", _
                    objJson, _
                    wsh, _
                    "Json_array", _
                    1
        Case Else 'must be a Number, String, Boolean, or null,
        'we can't get here currently as the JSON converter doesn't appear to handle JSON text that does not
        'begin with a dictionary or collection (object or array), this appears to be contrary to the spec and should be corrected
            varItem = objJson.Items()(0)
            mWriteElementToTable _
                varItem, _
                wsh, _
                "Json_Data", _
                varDataRow, _
                lngItem, _
                "Json_name"
    End Select
End Function

Private Function mCreateWorkSheet(ByRef sheetName As String, ByRef wkb As Workbook) As Worksheet
    If SheetExists(sheetName, wkb) Then
        mlngCurrentDupSheetCount = mlngCurrentDupSheetCount + 1
        Set mCreateWorkSheet = CreateWorksheet(sheetName & mlngCurrentDupSheetCount, wkb:=wkb)
    Else
        Set mCreateWorkSheet = CreateWorksheet(sheetName, wkb:=wkb)
    End If
    
End Function

Private Sub mWriteCollectionToSheet( _
    ByRef objItem As Variant, _
    ByRef wkb As Workbook, _
    ByRef strNewSheetName As String, _
    ByRef strKeyName As String, _
    ByRef objRecursive As Variant, _
    ByRef wsh As Worksheet, _
    ByRef varKey As Variant, _
    ByRef lngItem As Long _
)
Dim objItemElement As Variant
Dim lngItemElementCounter As Long
    lngItemElementCounter = 0
    For Each objItemElement In objItem
        lngItemElementCounter = lngItemElementCounter + 1
        If TypeName(objItemElement) = "Dictionary" Then
            objRecursive = GetAllJsonObjectNestedValues(objItemElement, wkb, wsh, strKeyName)
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
    ByRef varItem As Variant, _
    ByRef wsh As Worksheet, _
    ByRef strKeyName As String, _
    ByRef varDataRow As Variant, _
    ByRef lngItem As Long, _
    ByRef varKey As Variant _
)
    wsh.Activate
    If strKeyName = CStr(wsh.Range(Cells(1, 1), Cells(1, 1)).value) Then
        varDataRow = wsh.UsedRange.Rows.Count 'Assuming the same column header value = new row in any key value pair list
    End If
    If varDataRow = 1 Then
        wsh.Range(Cells(varDataRow, lngItem + 1), Cells(varDataRow, lngItem + 1)).value = varKey
    End If
    wsh.Range(Cells(varDataRow + 1, lngItem + 1), Cells(varDataRow + 1, lngItem + 1)).value = varItem
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
Private Function mCheckPath(ByVal path) As uriType
    Dim retval
    Select Case True 'select case only tests one at a time and stops on the first True solution.
        Case mHttpExists(path)
            retval = uriHttp
        Case mFileExists(path)
            retval = uriFile
        Case mFileExists(GetRelativePathViaParent(path))
            retval = uriFile
        Case mFolderExists(path)
            retval = uriDirectory
        Case mFolderExists(GetRelativePathViaParent(path))
            retval = uriParentDirectoryExists
        Case Else
            retval = uriUndefined
    End Select
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
    mFileExists = (Len(dir(strFile, lngAttributes)) > 0)
End Function

Private Function mFolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    mFolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
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

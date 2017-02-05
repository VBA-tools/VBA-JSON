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

Public Sub ImportJsonFileDailyToWorksheet( _
    ByRef strUrl As String, _
    ByRef strJSONObjectNameWithData As String, _
    Optional ByRef strDestinationSheetName As String _
)
'Use when the data posted to the web is only updated daily, we check to see if we have data for that day and only proceed after asking
Dim strSheetName As String
    If Len(strDestinationSheetName) > 0 Then
        strSheetName = strDestinationSheetName
    Else
        strSheetName = strURLDecode(Left(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/")), 30 - 6) & Format(Now(), "yymmdd"))
    End If
    strSheetName = RemoveForbiddenFilenameCharacters(strSheetName)
    If SheetExists(strSheetName) Then
        If ThisWorkbook.Sheets(strSheetName).UsedRange.Cells.Count > 2 Then
            If MsgBox("The site:" & strUrl & vbCrLf & vbCrLf & "Appears to have allready been imported into this workbook today, do you what to reperform this import?", vbYesNo) = vbYes Then
                DeleteSheet strSheetName
            Else
                GoTo ExitHere
            End If
        Else 'import sheet is empty... so delete and resume
            DeleteSheet strSheetName
        End If
    End If
    Dim strTempDownloadFile As String
    strTempDownloadFile = DownloadUrlFileToTemp(strUrl, "json")
    'Worksheet name length limit is 30 characters (-6 for date in yymmdd)
    #If DEBUGENABLED Then
        'Prints results to the imediate window, not structured.
        'First Working Example
        ' ExampleWriteJsonFileToDebugPrintImediateWindow strTempDownloadFile, strSheetName
        
        'Setting up generic debug print values JSON file interpreter
        WriteJsonFileToTheImediateWindow strTempDownloadFile, strJSONObjectNameWithData, strSheetName
    #End If
    
    'Delete our temp JSON file if we are done with it
    
ExitHere:
End Sub



Sub ExampleWriteJsonFileToDebugPrintImediateWindow(strJsonFilePath As String, Optional strSheetName As String)
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
    Dim optionChain As Dictionary 'or collection
    Set optionChain = jsonData("optionChain")
    Dim optionChainResults As Collection
    Set optionChainResults = optionChain("result")
    Dim optionChainResult
    For Each optionChainResult In optionChainResults
        Dim underlyingSymbol As String
        underlyingSymbol = optionChainResult("underlyingSymbol")
        '
        Dim expirationDates As Collection
        Set expirationDates = optionChainResult("expirationDates")
        Dim expirationDate As Variant
        For Each expirationDate In expirationDates
            Debug.Print underlyingSymbol & ":ExpirationDate=" & expirationDate
        Next
        '
        Dim strikes As Collection
        Dim strike As Variant
        Set strikes = optionChainResult("strikes")
        For Each strike In strikes
            Debug.Print underlyingSymbol & ":strike=" & strike
        Next
        '
        Dim hasMiniOptions As Boolean
        hasMiniOptions = optionChainResult("hasMiniOptions")
        Debug.Print underlyingSymbol & ":hasMiniOptions=" & hasMiniOptions
        '
        Dim quotes As Dictionary
        Set quotes = optionChainResult("quote")
        Dim aryQuotesPair As Variant
        aryQuotesPair = GetAllJsonObjectNestedValues(quotes)
        '
        Dim options As Collection
        Set options = optionChainResult("options")
        Dim optionItem As Variant
        For Each optionItem In options
            Dim optionDictionary As Dictionary
            Set optionDictionary = optionItem
            Dim aryOptionPair As Variant
            aryOptionPair = GetAllJsonObjectNestedValues(optionDictionary)
        Next optionItem
    Next optionChainResult
End Sub

Sub WriteJsonFileToTheImediateWindow(strJsonFilePath As String, strJSONObjectNameWithData As String, Optional strSheetName As String)
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
    GetAllJsonObjectNestedValues jsonData
End Sub


Private Function GetAllJsonObjectNestedValues(ByRef dict As Dictionary) As Variant
If dict.Count > 0 Then
Dim aryTemp As Variant
ReDim aryTemp(0 To 1, 0 To dict.Count - 1) '2 columns, with n rows
Dim lngItem As LongPtr
    For lngItem = 0 To dict.Count - 1
        Dim varKey
        Dim varItem
        varKey = dict.Keys()(lngItem)
        If IsObject(dict.Items()(lngItem)) Then
            Dim tmpDictionary As Dictionary
            Dim aryRecursive As Variant
            If TypeName(dict.Items()(lngItem)) = "Dictionary" Then
                Dim objItemDictionary As Dictionary
                Set objItemDictionary = dict.Items()(lngItem)
                aryRecursive = GetAllJsonObjectNestedValues(objItemDictionary)
            Else
                Dim objItem As Collection
                Set objItem = dict.Items()(lngItem)
                Select Case TypeName(objItem)
                    Case "Collection"
                        Dim objItemElement As Variant
                        For Each objItemElement In objItem
                            If TypeName(objItemElement) = "Dictionary" Then
                                Set tmpDictionary = objItemElement
                                aryRecursive = GetAllJsonObjectNestedValues(tmpDictionary)
                            Else
                                Debug.Print objItemElement
                            End If
                        Next objItemElement
                    Case "Dictionary"
                        Set tmpDictionary = objItem
                        aryRecursive = GetAllJsonObjectNestedValues(objItem) 'Executing this debug prints for testing untill we decide how to export this data to a spreedsheet appropriately
                End Select
            End If
            varItem = "Object"
        Else
            varItem = dict.Items()(lngItem)
        End If
        aryTemp(0, lngItem) = varKey
        aryTemp(1, lngItem) = varItem
        Debug.Print aryTemp(0, lngItem), aryTemp(1, lngItem)
    Next lngItem
End If
End Function





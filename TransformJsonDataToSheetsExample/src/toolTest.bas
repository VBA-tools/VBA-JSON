Attribute VB_Name = "toolTest"
Option Explicit

Public Sub ToolTestImportBlsGovJsonFile( _
    Optional strFileTargetDirectory As String, _
    Optional strArchiveDirectory As String, _
    Optional fForceTooManyCallsError As Boolean = False)
Dim strUrl As String
    'Usage:ToolTestImportBlsGovJsonFile "SolutionFiles", "JSON Archive"
    
    'this can be a list to loop through on an import/cover sheet that lists all json files to import.
    'Public requests to v1API are limited to 25 daily
    'From testing a single query called over and over again may not count toward this limit,
    'or the limits have been raised, made 6,600 sequential identical requsts (16mb total) in 10 minutes without error
    'can register at https://data.bls.gov/registrationEngine/ to increase limit to 500 requests/day
    'Series are formating is shown here:https://www.bls.gov/help/hlpforma.htm#
    Dim dblStartTime As Double
    dblStartTime = Timer
    strUrl = "https://api.bls.gov/publicAPI/v1/timeseries/data/LIUR0000SL00019" 'LEU0254555900'LIUR0000SL00019
    ImportJsonFileDailyToWorksheet strUrl, "series", , strArchiveDirectory, strFileTargetDirectory
    If fForceTooManyCallsError Then
        Dim i As Integer
        For i = 1 To 100
            ImportJsonFileDailyToWorksheet strUrl, "series", , strArchiveDirectory, strFileTargetDirectory
        Next i
    End If
    Debug.Print "Completed in: " & Timer - dblStartTime & "Seconds"
End Sub

Public Sub BrokenExampleWriteJsonFileToWorksheet(strJsonFilePath As String, Optional strSheetName As String)
' Advanced example: Read .json file and load into sheet (Windows-only) still working on this...
' (add reference to Microsoft Scripting Runtime)
' {"values":[{"a":1,"b":2,"c": 3},...]}
Dim fso As Object: Set fso = CreateObject("system.FileSystemObject") 'New FileSystemObject
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
    
    ' Parse json to Dictionary
    ' "values" is parsed as Collection
    ' each item in "values" is parsed as Dictionary
    
    Set Parsed = JsonConverter.ParseJson(JsonText)
    
    ' Prepare and write values to sheet
    Dim Values As Variant
    ReDim Values(Parsed("values").Count, 3)
    
    Dim value As Dictionary
    Dim i As Long
    
    i = 0
    For Each value In Parsed("values")
      Values(i, 0) = value("a")
      Values(i, 1) = value("b")
      Values(i, 2) = value("c")
      i = i + 1
    Next value
    If IsMissing(strSheetName) Then
        strSheetName = fso.GetFile(strJsonFilePath).Name
        strSheetName = Left(strSheetName, Len(strSheetName) - InStrRev(strSheetName, "."))
    End If
    ThisWorkbook.Sheets.Add (strSheetName)
    Sheets(strSheetName).Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values

End Sub




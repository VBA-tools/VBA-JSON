Attribute VB_Name = "toolTest"
Option Explicit
'Example JSON string for example parsing
'This example was from http://stackoverflow.com/questions/27723798/parsing-json-us-bls-in-vba-from-ms-access
'using http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.htm
'That project has a BSD lisence
Public Const jsonSource As String = "{" & _
  """status"": ""REQUEST_SUCCEEDED"", " & _
  """responseTime"": 71, " & _
  """message"": [ " & _
  "], " & _
  """Results"": { " & _
    """series"": [ " & _
      "{ " & _
        """seriesID"": ""WPS012"", " & _
        """data"": [ " & _
          "{ " & _
            """year"": ""2014"", " & _
            """period"": ""M11"", " & _
            """periodName"": ""November"", " & _
            """value"": ""153.6"", " & _
            """footnotes"": [ " & _
              "{ " & _
                """code"": ""P"", " & _
                """text"": ""Preliminary. All indexes are subject to revision four months after original publication."" " & _
              "} " & _
            "] " & _
          "} " & _
        "] " & _
      "}]}}"

Sub JsonTest()
    Dim jsonData As Dictionary
    Set jsonData = ParseJson(jsonSource)

    Dim responseTime As String
    responseTime = jsonData("responseTime")

    Dim results As Dictionary
    Set results = jsonData("Results")

    Dim series As Collection
    Set series = results("series")

    Dim seriesItem As Dictionary
    For Each seriesItem In series
        Dim seriesId As String
        seriesId = seriesItem("seriesID")
        Debug.Print seriesId

        Dim data As Collection
        Set data = seriesItem("data")

        Dim dataItem As Dictionary
        For Each dataItem In data
            Dim year As String
            year = dataItem("year")

            Dim period As String
            period = dataItem("period")

            Dim periodName As String
            periodName = dataItem("periodName")

            Dim value As String
            value = dataItem("value")

            Dim footnotes As Collection
            Set footnotes = dataItem("footnotes")

            Dim footnotesItem As Dictionary
            For Each footnotesItem In footnotes
                Dim code As String
                code = footnotesItem("code")

                Dim text As String
                text = footnotesItem("text")

            Next footnotesItem
        Next dataItem
    Next seriesItem
End Sub

'Gerdes, Jeremy
Public Sub TestImportYahooUvxy()
'YAHOO api usage limitations: http://meumobi.github.io/stocks%20apis/2016/03/13/get-realtime-stock-quotes-yahoo-finance-api.html
'Free Yahooo API public limit is 20,000 queries per hour per IP, not to exceed 48,000 per day
'Alternate data sources: https://www.programmableweb.com/news/96-stocks-apis-bloomberg-nasdaq-and-etrade/2013/05/22
    Dim strUrl As String
    'this can be a list to loop through on an import/cover sheet that lists all json files to import.
    strUrl = "https://query1.finance.yahoo.com/v7/finance/options/UVXY"
    ImportJsonFileDailyToWorksheet strUrl, "optionChain"
    
    'Get quote only
    strUrl = "https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20(%22UVXY%22)&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback="
    ImportJsonFileDailyToWorksheet strUrl, "quote", "UVXY_Quote"
    
    'or any other symbol i.e.
    strUrl = "https://query1.finance.yahoo.com/v7/finance/options/IDX"
    ImportJsonFileDailyToWorksheet strUrl, "optionChain"
    
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

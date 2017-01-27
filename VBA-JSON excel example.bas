Option Explicit
public sub WriteJsonFileToWorksheet(strJsonFilePath as string, optional strSheetname as string)
' Advanced example: Read .json file and load into sheet (Windows-only)
' (add reference to Microsoft Scripting Runtime)
' {"values":[{"a":1,"b":2,"c": 3},...]}

Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary

' Read .json file
Set JsonTS = FSO.OpenTextFile(strJsonFilePath, ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = JsonConverter.ParseJson(JsonText)

' Prepare and write values to sheet
Dim Values As Variant
ReDim Values(Parsed("values").Count, 3)

Dim Value As Dictionary
Dim i As Long

i = 0
For Each Value In Parsed("values")
  Values(i, 0) = Value("a")
  Values(i, 1) = Value("b")
  Values(i, 2) = Value("c")
  i = i + 1
Next Value
If ismissing(strSheetname) then
    strSheetname =   FSO.GetFile(strJsonFilePath).Name
    strSheetname = left(strSheetname,len(strSheetname)-instrrev(strSheetname,"."))
End If
ThisWorkbook.Sheets.Add(strSheetname)
Sheets(strSheetname).Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
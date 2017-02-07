Attribute VB_Name = "basSheets"
Option Explicit
'Authored 2014-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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

Public Function CreateWorksheet( _
    strSheetName As String, _
    Optional shtAfter As Object, _
    Optional fDeleteExisting As Boolean = True, _
    Optional wkb As Workbook) _
As Worksheet
    If wkb Is Nothing Then
        Set wkb = ThisWorkbook
    End If
wkb.Activate ' ensure this workbook is active prior to creating sheets

Dim wkshts As Sheets: Set wkshts = wkb.Sheets
Dim sht As Worksheet
    
    If SheetExists(strSheetName) Then
        If fDeleteExisting Then
            wkshts(strSheetName).Delete
        Else
            MsgBox "A sheet by the name " & strSheetName & " allready exists, new sheet not created"
        End If
    End If
    
    If IsNotSheet(shtAfter) Then
        Set shtAfter = wkshts(wkshts.Count)
    End If
    
    Set sht = wkshts.Add(After:=shtAfter)
    sht.Name = strSheetName
    sht.Activate
    
    Set CreateWorksheet = sht
    
End Function

Public Function IsNotSheet(obj As Variant)
On Error Resume Next
    Dim wsh As Worksheet
    Set wsh = obj
    Dim strName As String
    strName = wsh.Name
    IsNotSheet = Err.Number
End Function

Public Function IsNotChart(obj As Variant)
On Error Resume Next
    Dim chrt As Chart
    Set chrt = obj
    Dim strName As String
    strName = chrt.Name
    IsNotChart = Err.Number
End Function

Public Function CreateChart( _
    strSheetName As String, _
    Optional chtAfter As Object, _
    Optional fDeleteExisting As Boolean = True) _
As Chart
ThisWorkbook.Activate ' ensure this workbook is active prior to creating sheets
Dim wkchts As Sheets: Set wkchts = ThisWorkbook.Charts
Dim cht As Chart

    If SheetExists(strSheetName) Then
        If fDeleteExisting Then
            DeleteSheet (strSheetName)
        Else
            MsgBox "A sheet by the name " & strSheetName & " allready exists, new sheet not created"
            Set CreateChart = Nothing
            Exit Function
        End If
    End If
    
    If IsNotChart(chtAfter) Then
        Set chtAfter = wkchts(wkchts.Count)
    End If
    
    'Build the temp chart with very little data and move to chartSheet, for speed plan on setting range latter
    ActiveSheet.Range("A1", "A2").Select
    Dim shpChtNew As Shape
    Dim objChtNew As ChartObject
    
    Set shpChtNew = ActiveSheet.Shapes.AddChart(xlXYScatterLines, 1, 1, 10, 10)
   
    Set objChtNew = shpChtNew.Chart.Parent
    objChtNew.Activate
    Set cht = ActiveChart.Location(Where:=xlLocationAsNewSheet, Name:=strSheetName)
    
    'This method takes too long, can't define Range, and uses defualt chart type what ever that is (can cause errors)
    'Set cht = wkchts.Add(after:=chtAfter)
    'cht.Name = strSheetName
    
    cht.Activate
    
    Set CreateChart = cht

End Function

Public Function SheetExists(strSheetName As String, Optional wkb As Workbook) As Boolean
If wkb Is Nothing Then
    Set wkb = ThisWorkbook
End If
Dim sht As Object
Dim fFoundSheet As Boolean
fFoundSheet = False
     For Each sht In wkb.Sheets
         If LCase(sht.Name) = LCase(strSheetName) Then
             fFoundSheet = True
             Exit For
         End If
     Next
    SheetExists = fFoundSheet
End Function

Public Function DeleteSheet(strSheetName As String, Optional wkb As Workbook) As Boolean 'Returns True if succeeds or sheet never existed
On Error Resume Next
If wkb Is Nothing Then
    Set wkb = ThisWorkbook
End If
Dim sht As Object ' any Sheet Type
    If SheetExists(strSheetName, wkb) Then
        Application.DisplayAlerts = False
        wkb.Sheets(strSheetName).Delete
        Application.DisplayAlerts = True
    End If
    DeleteSheet = Err.Number = 0
End Function


Public Sub RevealAllSheets()
On Error Resume Next
Dim sht As Object
    For Each sht In Sheets
        sht.Visible = True
    Next
    
End Sub

' ToolResetSheetsCodeNamesWill not work unless the Project is trusted in the trust center (Disabled option on Current Configuration)
Private Sub ToolResetSheetsCodeNames(strCodeNamePrefix As String)
Dim oVBComponent As Object 'VBIDE.VBComponent
Dim wkshts As Sheets: Set wkshts = ThisWorkbook.Sheets
Dim sht As Worksheet
Dim intSheetCodeNameCount: intSheetCodeNameCount = 1
    For Each sht In wkshts
          sht.Activate
            For Each oVBComponent In ActiveSheet.Parent.VBProject.VBComponents ' Besides looking at the Name you can also test on Type, etc. (not done here).
            If (oVBComponent.Name = ActiveSheet.CodeName) Then
                oVBComponent.Name = strCodeNamePrefix & intSheetCodeNameCount
            End If
        Next oVBComponent
    Next

End Sub





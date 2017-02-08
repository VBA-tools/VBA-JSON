Attribute VB_Name = "basUrlEncodeDecode"
Option Explicit

Public Function strUrlEncode( _
   strStringToEncode As String, _
   Optional fUsePlusRatherThanHexForSpace As Boolean = False _
) As String

  Dim strCurUri As String
  Dim intCurChr As Integer
  intCurChr = 1

  Do Until intCurChr - 1 = Len(strStringToEncode)
    Select Case Asc(Mid(strStringToEncode, intCurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        strCurUri = strCurUri & Mid(strStringToEncode, intCurChr, 1)
      Case 32
        If fUsePlusRatherThanHexForSpace = True Then
          strCurUri = strCurUri & "+"
        Else
          strCurUri = strCurUri & "%" & Hex(32)
        End If
      Case Else
        strCurUri = strCurUri & "%" & _
          Right("0" & Hex(Asc(Mid(strStringToEncode, _
          intCurChr, 1))), 2)
    End Select
    intCurChr = intCurChr + 1
  Loop

  strUrlEncode = strCurUri

End Function

Public Function strUrlDecode(strEncodedUri As String) As String

Dim intCurChr As Integer
Dim strCurUri As String
Dim strTempUri As String

    If Len(strEncodedUri) > 0 Then
        ' Loop through each char
        For intCurChr = 1 To Len(strEncodedUri)
            strTempUri = Mid(strEncodedUri, intCurChr, 1)
            strTempUri = Replace(strTempUri, "+", " ")
            ' If char is % then get next two chars
            ' and convert from HEX to decimal
            If strTempUri = "%" And Len(strEncodedUri) + 1 > intCurChr + 2 Then
                strTempUri = Mid(strEncodedUri, intCurChr + 1, 2)
                strTempUri = Chr(CDec("&H" & strTempUri))
                ' Increment loop by 2
                intCurChr = intCurChr + 2
            End If
            strCurUri = strCurUri & strTempUri
        Next

        strUrlDecode = strCurUri
    Else
        strUrlDecode = ""
    End If
End Function


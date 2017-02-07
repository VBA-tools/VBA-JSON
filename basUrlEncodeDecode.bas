Attribute VB_Name = "basUrlEncodeDecode"
Option Explicit

Public Function strURLEncode( _
   strStringToEncode As String, _
   Optional fUsePlusRatherThanHexForSpace As Boolean = False _
) As String

  Dim strCurURL As String
  Dim intCurChr As Integer
  intCurChr = 1

  Do Until intCurChr - 1 = Len(strStringToEncode)
    Select Case Asc(Mid(strStringToEncode, intCurChr, 1))
      Case 48 To 57, 65 To 90, 97 To 122
        strCurURL = strCurURL & Mid(strStringToEncode, intCurChr, 1)
      Case 32
        If fUsePlusRatherThanHexForSpace = True Then
          strCurURL = strCurURL & "+"
        Else
          strCurURL = strCurURL & "%" & Hex(32)
        End If
      Case Else
        strCurURL = strCurURL & "%" & _
          Right("0" & Hex(Asc(Mid(strStringToEncode, _
          intCurChr, 1))), 2)
    End Select
    intCurChr = intCurChr + 1
  Loop

  strURLEncode = strCurURL

End Function

Public Function strURLDecode(strEncodedURL As String) As String

Dim intCurChr As Integer
Dim strCurURL As String
Dim strTempURL As String

    If Len(strEncodedURL) > 0 Then
        ' Loop through each char
        For intCurChr = 1 To Len(strEncodedURL)
            strTempURL = Mid(strEncodedURL, intCurChr, 1)
            strTempURL = Replace(strTempURL, "+", " ")
            ' If char is % then get next two chars
            ' and convert from HEX to decimal
            If strTempURL = "%" And Len(strEncodedURL) + 1 > intCurChr + 2 Then
                strTempURL = Mid(strEncodedURL, intCurChr + 1, 2)
                strTempURL = Chr(CDec("&H" & strTempURL))
                ' Increment loop by 2
                intCurChr = intCurChr + 2
            End If
            strCurURL = strCurURL & strTempURL
        Next

        strURLDecode = strCurURL
    Else
        strURLDecode = ""
    End If
End Function


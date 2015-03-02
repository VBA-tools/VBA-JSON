Attribute VB_Name = "JsonConverter"
''
' VBA-JSON v1.0.2
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' === VBA-UTC Headers
#If Mac Then

Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" (ByVal utc_File As Long) As Long

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)

#Else

Private Declare Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)
    
#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal json_String As String, Optional json_ConvertLargeNumbersToString As Boolean = True) As Object
    Dim json_Index As Long
    json_Index = 1
    
    ' Remove vbCr, vbLf, and vbTab from json_String
    json_String = VBA.Replace(VBA.Replace(VBA.Replace(json_String, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(json_String, json_Index, json_ConvertLargeNumbersToString)
    Case "["
        Set ParseJson = json_ParseArray(json_String, json_Index, json_ConvertLargeNumbersToString)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} json_DictionaryCollectionOrArray (Dictionary, Collection, or Array)
' @return {String}
''
Public Function ConvertToJson(ByVal json_DictionaryCollectionOrArray As Variant, Optional json_ConvertLargeNumbersFromString As Boolean = True) As String
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    
    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True

    Select Case VBA.VarType(json_DictionaryCollectionOrArray)
    Case VBA.vbNull, VBA.vbEmpty
        ConvertToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(json_DictionaryCollectionOrArray))
        
        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If json_ConvertLargeNumbersFromString And json_StringIsLargeNumber(json_DictionaryCollectionOrArray) Then
            ConvertToJson = json_DictionaryCollectionOrArray
        Else
            ConvertToJson = """" & json_Encode(json_DictionaryCollectionOrArray) & """"
        End If
    Case VBA.vbBoolean
        If json_DictionaryCollectionOrArray Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        ' Array
        json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
        
        On Error Resume Next
        
        json_LBound = LBound(json_DictionaryCollectionOrArray, 1)
        json_UBound = UBound(json_DictionaryCollectionOrArray, 1)
        json_LBound2D = LBound(json_DictionaryCollectionOrArray, 2)
        json_UBound2D = UBound(json_DictionaryCollectionOrArray, 2)
        
        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
                
                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                        End If
                        
                        json_BufferAppend json_buffer, _
                            ConvertToJson(json_DictionaryCollectionOrArray(json_Index, json_Index2D), _
                                json_ConvertLargeNumbersFromString), _
                            json_BufferPosition, json_BufferLength
                    Next json_Index2D
                    
                    json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    json_BufferAppend json_buffer, _
                        ConvertToJson(json_DictionaryCollectionOrArray(json_Index), _
                            json_ConvertLargeNumbersFromString), _
                        json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If
        
        On Error GoTo 0
        
        json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
        
        ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
    
    ' Dictionary or Collection
    Case VBA.vbObject
        ' Dictionary
        If VBA.TypeName(json_DictionaryCollectionOrArray) = "Dictionary" Then
            json_BufferAppend json_buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In json_DictionaryCollectionOrArray.Keys
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                json_BufferAppend json_buffer, _
                    """" & json_Key & """:" & ConvertToJson(json_DictionaryCollectionOrArray(json_Key), json_ConvertLargeNumbersFromString), _
                    json_BufferPosition, json_BufferLength
            Next json_Key
            json_BufferAppend json_buffer, "}", json_BufferPosition, json_BufferLength
        
        ' Collection
        ElseIf VBA.TypeName(json_DictionaryCollectionOrArray) = "Collection" Then
            json_BufferAppend json_buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In json_DictionaryCollectionOrArray
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_buffer, ",", json_BufferPosition, json_BufferLength
                End If
            
                json_BufferAppend json_buffer, _
                    ConvertToJson(json_Value, json_ConvertLargeNumbersFromString), _
                    json_BufferPosition, json_BufferLength
            Next json_Value
            json_BufferAppend json_buffer, "]", json_BufferPosition, json_BufferLength
        End If
        
        ConvertToJson = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
    Case Else
        ' Number
        On Error Resume Next
        ConvertToJson = VBA.Replace(json_DictionaryCollectionOrArray, ",", ".")
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Dictionary
    Dim json_Key As String
    Dim json_NextChar As String
    
    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1
        
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            
            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Collection
    Set json_ParseArray = New Collection
    
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1
        
        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If
            
            json_ParseArray.Add json_ParseValue(json_String, json_Index, json_ConvertLargeNumbersToString)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index, json_ConvertLargeNumbersToString)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    
    json_SkipSpaces json_String, json_Index
    
    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1
    
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)
        
        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            
            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long, Optional json_ConvertLargeNumbersToString As Boolean = True) As Variant
    Dim json_Char As String
    Dim json_Value As String
    
    json_SkipSpaces json_String, json_Index
    
    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)
        
        If VBA.InStr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15 characters containing only numbers and decimal points -> Number
            If json_ConvertLargeNumbersToString And Len(json_Value) >= 16 Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    json_ParseKey = json_ParseString(json_String, json_Index)
    
    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    
    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)
        
        Select Case json_AscCode
        ' " -> 34 -> \"
        Case 34
            json_Char = "\"""
        ' \ -> 92 -> \\
        Case 92
            json_Char = "\\"
        ' / -> 47 -> \/
        Case 47
            json_Char = "\/"
        ' backspace -> 8 -> \b
        Case 8
            json_Char = "\b"
        ' form feed -> 12 -> \f
        Case 12
            json_Char = "\f"
        ' line feed -> 10 -> \n
        Case 10
            json_Char = "\n"
        ' carriage return -> 13 -> \r
        Case 13
            json_Char = "\r"
        ' tab -> 9 -> \t
        Case 9
            json_Char = "\t"
        ' Non-ascii characters -> convert to 4-digit hex
        Case 0 To 31, 127 To 65535
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select
            
        json_BufferAppend json_buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index
    
    json_Encode = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)
    
    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String
        Dim json_Index As Long
        
        json_StringIsLargeNumber = True
        
        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
    
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
#If Mac Then
    json_buffer = json_buffer & json_Append
#Else
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long
    
    json_AppendLength = VBA.LenB(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition
    
    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim json_TemporaryLength As Long
        
        json_TemporaryLength = json_BufferLength
        Do While json_TemporaryLength < json_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If json_TemporaryLength = 0 Then
                json_TemporaryLength = json_TemporaryLength + 510
            Else
                json_TemporaryLength = json_TemporaryLength + 16384
            End If
        Loop
        
        json_buffer = json_buffer & VBA.Space$((json_TemporaryLength - json_BufferLength) \ 2)
        json_BufferLength = json_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    json_CopyMemory ByVal json_UnsignedAdd(StrPtr(json_buffer), _
                    json_BufferPosition), _
                    ByVal StrPtr(json_Append), _
                    json_AppendLength
    
    json_BufferPosition = json_BufferPosition + json_AppendLength
#End If
End Sub

Private Function json_BufferToString(ByRef json_buffer As String, ByVal json_BufferPosition As Long, ByVal json_BufferLength As Long) As String
#If Mac Then
    json_BufferToString = json_buffer
#Else
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_buffer, json_BufferPosition \ 2)
    End If
#End If
End Function

#If VBA7 Then
Private Function json_UnsignedAdd(json_Start As LongPtr, json_Increment As Long) As LongPtr
#Else
Private Function json_UnsignedAdd(json_Start As Long, json_Increment As Long) As Long
#End If

    If json_Start And &H80000000 Then
        json_UnsignedAdd = json_Start + json_Increment
    ElseIf (json_Start Or &H80000000) < -json_Increment Then
        json_UnsignedAdd = json_Start + json_Increment
    Else
        json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
    End If
End Function

''
' VBA-UTC v1.0.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling
    
#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate
    
    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling
    
#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME
    
    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate
    
    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If
    
    Exit Function
    
utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling
    
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date
    
    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))
    
    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If
            
            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")
                
                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), VBA.CInt(utc_OffsetParts(2)))
                End Select
                
                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If
        
        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), VBA.CInt(utc_TimeParts(2)))
        End Select
        
        If utc_HasOffset Then
            ParseIso = ParseIso + utc_Offset
        Else
            ParseIso = ParseUtc(ParseIso)
        End If
    End If
    
    Exit Function
    
utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling
    
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    
    Exit Function
    
utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    
    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If
    
    utc_Result = utc_ExecuteInShell(utc_ShellCommand)
    
    If utc_Result.utc_Output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_Output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")
        
        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
    Dim utc_File As Long
    Dim utc_Chunk As String
    Dim utc_Read As Long
    
    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")
    
    If utc_File = 0 Then: Exit Function
    
    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File)
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, utc_Read)
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = utc_pclose(utc_File)
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If

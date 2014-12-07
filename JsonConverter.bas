Attribute VB_Name = "JsonConverter"
''
' VBA-JSON v1.0.0-beta.1
' (c) Tim Hall - https://github.com/timhall/VBA-JSONConverter
'
' JSON Converter for VBA
'
' Errors (513-65535 available):
' 10001 - JSON parse error
' 10002 - ISO 8601 date conversion error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
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
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

#If Mac Then

Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" (ByVal utc_command As String, ByVal utc_mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" (ByVal utc_file As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" (ByVal utc_buffer As String, ByVal utc_size As Long, ByVal utc_number As Long, ByVal utc_file As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" (ByVal utc_file As Long) As Long

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#Else

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

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

#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
#Else
Private Declare Sub JSON_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (JSON_MemoryDestination As Any, JSON_MemorySource As Any, ByVal JSON_ByteLength As Long)
#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @param {String} JSON_String
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function ParseJSON(ByVal JSON_String As String, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Object
    Dim JSON_Index As Long
    JSON_Index = 1
    
    ' Remove vbCr, vbLf, and vbTab from JSON_String
    JSON_String = VBA.Replace(VBA.Replace(VBA.Replace(JSON_String, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case VBA.Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set ParseJSON = JSON_ParseObject(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
    Case "["
        Set ParseJSON = JSON_ParseArray(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @param {Variant} JSON_DictionaryCollectionOrArray (Dictionary, Collection, or Array)
' @return {String}
' -------------------------------------- '
Public Function ConvertToJSON(ByVal JSON_DictionaryCollectionOrArray As Variant, Optional JSON_ConvertLargeNumbersFromString As Boolean = True) As String
    Dim json_buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    Dim JSON_Index As Long
    Dim JSON_LBound As Long
    Dim JSON_UBound As Long
    Dim JSON_IsFirstItem As Boolean
    Dim JSON_Index2D As Long
    Dim JSON_LBound2D As Long
    Dim JSON_UBound2D As Long
    Dim JSON_IsFirstItem2D As Boolean
    Dim JSON_Key As Variant
    Dim JSON_Value As Variant
    Dim JSON_DateStr As String
    
    JSON_LBound = -1
    JSON_UBound = -1
    JSON_IsFirstItem = True
    JSON_LBound2D = -1
    JSON_UBound2D = -1
    JSON_IsFirstItem2D = True

    Select Case VBA.VarType(JSON_DictionaryCollectionOrArray)
    Case VBA.vbNull, VBA.vbEmpty
        ConvertToJSON = "null"
    Case VBA.vbDate
        ' Date
        JSON_DateStr = ConvertToIso(VBA.CDate(JSON_DictionaryCollectionOrArray))
        
        ConvertToJSON = """" & JSON_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If JSON_ConvertLargeNumbersFromString And JSON_StringIsLargeNumber(JSON_DictionaryCollectionOrArray) Then
            ConvertToJSON = JSON_DictionaryCollectionOrArray
        Else
            ConvertToJSON = """" & JSON_Encode(JSON_DictionaryCollectionOrArray) & """"
        End If
    Case VBA.vbBoolean
        If JSON_DictionaryCollectionOrArray Then
            ConvertToJSON = "true"
        Else
            ConvertToJSON = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        ' Array
        JSON_BufferAppend json_buffer, "[", JSON_BufferPosition, JSON_BufferLength
        
        On Error Resume Next
        
        JSON_LBound = LBound(JSON_DictionaryCollectionOrArray, 1)
        JSON_UBound = UBound(JSON_DictionaryCollectionOrArray, 1)
        JSON_LBound2D = LBound(JSON_DictionaryCollectionOrArray, 2)
        JSON_UBound2D = UBound(JSON_DictionaryCollectionOrArray, 2)
        
        If JSON_LBound >= 0 And JSON_UBound >= 0 Then
            For JSON_Index = JSON_LBound To JSON_UBound
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend json_buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                If JSON_LBound2D >= 0 And JSON_UBound2D >= 0 Then
                    JSON_BufferAppend json_buffer, "[", JSON_BufferPosition, JSON_BufferLength
                
                    For JSON_Index2D = JSON_LBound2D To JSON_UBound2D
                        If JSON_IsFirstItem2D Then
                            JSON_IsFirstItem2D = False
                        Else
                            JSON_BufferAppend json_buffer, ",", JSON_BufferPosition, JSON_BufferLength
                        End If
                        
                        JSON_BufferAppend json_buffer, _
                            ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index, JSON_Index2D), _
                                JSON_ConvertLargeNumbersFromString), _
                            JSON_BufferPosition, JSON_BufferLength
                    Next JSON_Index2D
                    
                    JSON_BufferAppend json_buffer, "]", JSON_BufferPosition, JSON_BufferLength
                    JSON_IsFirstItem2D = True
                Else
                    JSON_BufferAppend json_buffer, _
                        ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index), _
                            JSON_ConvertLargeNumbersFromString), _
                        JSON_BufferPosition, JSON_BufferLength
                End If
            Next JSON_Index
        End If
        
        On Error GoTo 0
        
        JSON_BufferAppend json_buffer, "]", JSON_BufferPosition, JSON_BufferLength
        
        ConvertToJSON = JSON_BufferToString(json_buffer, JSON_BufferPosition, JSON_BufferLength)
    
    ' Dictionary or Collection
    Case VBA.vbObject
        ' Dictionary
        If VBA.TypeName(JSON_DictionaryCollectionOrArray) = "Dictionary" Then
            JSON_BufferAppend json_buffer, "{", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Key In JSON_DictionaryCollectionOrArray.Keys
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend json_buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend json_buffer, _
                    """" & JSON_Key & """:" & ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Key), JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Key
            JSON_BufferAppend json_buffer, "}", JSON_BufferPosition, JSON_BufferLength
        
        ' Collection
        ElseIf VBA.TypeName(JSON_DictionaryCollectionOrArray) = "Collection" Then
            JSON_BufferAppend json_buffer, "[", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Value In JSON_DictionaryCollectionOrArray
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend json_buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend json_buffer, _
                    ConvertToJSON(JSON_Value, JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Value
            JSON_BufferAppend json_buffer, "]", JSON_BufferPosition, JSON_BufferLength
        End If
        
        ConvertToJSON = JSON_BufferToString(json_buffer, JSON_BufferPosition, JSON_BufferLength)
    Case Else
        ' Number
        On Error Resume Next
        ConvertToJSON = JSON_DictionaryCollectionOrArray
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function JSON_ParseObject(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Dictionary
    Dim JSON_Key As String
    Dim JSON_NextChar As String
    
    Set JSON_ParseObject = New Dictionary
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{'")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If VBA.Mid$(JSON_String, JSON_Index, 1) = "}" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_Key = JSON_ParseKey(JSON_String, JSON_Index)
            JSON_NextChar = JSON_Peek(JSON_String, JSON_Index)
            If JSON_NextChar = "[" Or JSON_NextChar = "{" Then
                Set JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
            Else
                JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
            End If
        Loop
    End If
End Function

Private Function JSON_ParseArray(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Collection
    Set JSON_ParseArray = New Collection
    
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '['")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If VBA.Mid$(JSON_String, JSON_Index, 1) = "]" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_ParseArray.Add JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
        Loop
    End If
End Function

Private Function JSON_ParseValue(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case VBA.Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set JSON_ParseValue = JSON_ParseObject(JSON_String, JSON_Index)
    Case "["
        Set JSON_ParseValue = JSON_ParseArray(JSON_String, JSON_Index)
    Case """", "'"
        JSON_ParseValue = JSON_ParseString(JSON_String, JSON_Index)
    Case Else
        If VBA.Mid$(JSON_String, JSON_Index, 4) = "true" Then
            JSON_ParseValue = True
            JSON_Index = JSON_Index + 4
        ElseIf VBA.Mid$(JSON_String, JSON_Index, 5) = "false" Then
            JSON_ParseValue = False
            JSON_Index = JSON_Index + 5
        ElseIf VBA.Mid$(JSON_String, JSON_Index, 4) = "null" Then
            JSON_ParseValue = Null
            JSON_Index = JSON_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(JSON_String, JSON_Index, 1)) Then
            JSON_ParseValue = JSON_ParseNumber(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
        Else
            Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function JSON_ParseString(JSON_String As String, ByRef JSON_Index As Long) As String
    Dim JSON_Quote As String
    Dim JSON_Char As String
    Dim JSON_Code As String
    Dim json_buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    ' Store opening quote to look for matching closing quote
    JSON_Quote = VBA.Mid$(JSON_String, JSON_Index, 1)
    JSON_Index = JSON_Index + 1
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
        
        Select Case JSON_Char
        Case "\"
            ' Escaped string, \\, or \/
            JSON_Index = JSON_Index + 1
            JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
            
            Select Case JSON_Char
            Case """", "\", "/", "'"
                JSON_BufferAppend json_buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "b"
                JSON_BufferAppend json_buffer, vbBack, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "f"
                JSON_BufferAppend json_buffer, vbFormFeed, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "n"
                JSON_BufferAppend json_buffer, vbCrLf, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "r"
                JSON_BufferAppend json_buffer, vbCr, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "t"
                JSON_BufferAppend json_buffer, vbTab, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                JSON_Index = JSON_Index + 1
                JSON_Code = VBA.Mid$(JSON_String, JSON_Index, 4)
                JSON_BufferAppend json_buffer, VBA.ChrW(VBA.Val("&h" + JSON_Code)), JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 4
            End Select
        Case JSON_Quote
            JSON_ParseString = JSON_BufferToString(json_buffer, JSON_BufferPosition, JSON_BufferLength)
            JSON_Index = JSON_Index + 1
            Exit Function
        Case Else
            JSON_BufferAppend json_buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
            JSON_Index = JSON_Index + 1
        End Select
    Loop
End Function

Private Function JSON_ParseNumber(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    Dim JSON_Char As String
    Dim JSON_Value As String
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
        
        If VBA.InStr("+-0123456789.eE", JSON_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            JSON_Value = JSON_Value & JSON_Char
            JSON_Index = JSON_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15 characters containing only numbers and decimal points -> Number
            If JSON_ConvertLargeNumbersToString And Len(JSON_Value) >= 16 Then
                JSON_ParseNumber = JSON_Value
            Else
                ' Guard for regional settings that use "," for decimal
                ' CStr(0.1) -> "0.1" or "0,1" based on regional settings -> Replace "." with "." or ","
                JSON_Value = VBA.Replace(JSON_Value, ".", VBA.Mid$(VBA.CStr(0.1), 2, 1))
                JSON_ParseNumber = VBA.Val(JSON_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function JSON_ParseKey(JSON_String As String, ByRef JSON_Index As Long) As String
    ' Parse key with single or double quotes
    JSON_ParseKey = JSON_ParseString(JSON_String, JSON_Index)
    
    ' Check for colon and skip if present or throw if not present
    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting ':'")
    Else
        JSON_Index = JSON_Index + 1
    End If
End Function

Private Function JSON_Encode(ByVal JSON_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim JSON_Index As Long
    Dim JSON_Char As String
    Dim JSON_AscCode As Long
    Dim json_buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    
    For JSON_Index = 1 To VBA.Len(JSON_Text)
        JSON_Char = VBA.Mid$(JSON_Text, JSON_Index, 1)
        JSON_AscCode = VBA.AscW(JSON_Char)
        
        Select Case JSON_AscCode
        ' " -> 34 -> \"
        Case 34
            JSON_Char = "\"""
        ' \ -> 92 -> \\
        Case 92
            JSON_Char = "\\"
        ' / -> 47 -> \/
        Case 47
            JSON_Char = "\/"
        ' backspace -> 8 -> \b
        Case 8
            JSON_Char = "\b"
        ' form feed -> 12 -> \f
        Case 12
            JSON_Char = "\f"
        ' line feed -> 10 -> \n
        Case 10
            JSON_Char = "\n"
        ' carriage return -> 13 -> \r
        Case 13
            JSON_Char = "\r"
        ' tab -> 9 -> \t
        Case 9
            JSON_Char = "\t"
        ' Non-ascii characters -> convert to 4-digit hex
        Case 0 To 31, 127 To 65535
            JSON_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(JSON_AscCode), 4)
        End Select
            
        JSON_BufferAppend json_buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
    Next JSON_Index
    
    JSON_Encode = JSON_BufferToString(json_buffer, JSON_BufferPosition, JSON_BufferLength)
End Function

Private Function JSON_Peek(JSON_String As String, ByVal JSON_Index As Long, Optional JSON_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing JSON_Index (ByVal instead of ByRef)
    JSON_SkipSpaces JSON_String, JSON_Index
    JSON_Peek = VBA.Mid$(JSON_String, JSON_Index, JSON_NumberOfCharacters)
End Function

Private Sub JSON_SkipSpaces(JSON_String As String, ByRef JSON_Index As Long)
    ' Increment index to skip over spaces
    Do While JSON_Index > 0 And JSON_Index <= VBA.Len(JSON_String) And VBA.Mid$(JSON_String, JSON_Index, 1) = " "
        JSON_Index = JSON_Index + 1
    Loop
End Sub

Private Function JSON_StringIsLargeNumber(JSON_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See JSON_ParseNumber)
    
    Dim JSON_Length As Long
    Dim JSON_CharIndex As Long
    JSON_Length = VBA.Len(JSON_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If JSON_Length >= 16 And JSON_Length <= 100 Then
        Dim JSON_CharCode As String
        Dim JSON_Index As Long
        
        JSON_StringIsLargeNumber = True
        
        For JSON_CharIndex = 1 To JSON_Length
            JSON_CharCode = VBA.Asc(VBA.Mid$(JSON_String, JSON_CharIndex, 1))
            Select Case JSON_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                JSON_StringIsLargeNumber = False
                Exit Function
            End Select
        Next JSON_CharIndex
    End If
End Function

Private Function JSON_ParseErrorMessage(JSON_String As String, ByRef JSON_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['
    
    Dim JSON_StartIndex As Long
    Dim JSON_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    JSON_StartIndex = JSON_Index - 10
    JSON_StopIndex = JSON_Index + 10
    If JSON_StartIndex <= 0 Then
        JSON_StartIndex = 1
    End If
    If JSON_StopIndex > VBA.Len(JSON_String) Then
        JSON_StopIndex = VBA.Len(JSON_String)
    End If

    JSON_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(JSON_String, JSON_StartIndex, JSON_StopIndex - JSON_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(JSON_Index - JSON_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub JSON_BufferAppend(ByRef json_buffer As String, _
                              ByRef JSON_Append As Variant, _
                              ByRef JSON_BufferPosition As Long, _
                              ByRef JSON_BufferLength As Long)
#If Mac Then
    json_buffer = json_buffer & JSON_Append
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

    Dim JSON_AppendLength As Long
    Dim JSON_LengthPlusPosition As Long
    
    JSON_AppendLength = VBA.LenB(JSON_Append)
    JSON_LengthPlusPosition = JSON_AppendLength + JSON_BufferPosition
    
    If JSON_LengthPlusPosition > JSON_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim JSON_TemporaryLength As Long
        
        JSON_TemporaryLength = JSON_BufferLength
        Do While JSON_TemporaryLength < JSON_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If JSON_TemporaryLength = 0 Then
                JSON_TemporaryLength = JSON_TemporaryLength + 510
            Else
                JSON_TemporaryLength = JSON_TemporaryLength + 16384
            End If
        Loop
        
        json_buffer = json_buffer & VBA.Space$((JSON_TemporaryLength - JSON_BufferLength) \ 2)
        JSON_BufferLength = JSON_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    JSON_CopyMemory ByVal JSON_UnsignedAdd(StrPtr(json_buffer), _
                    JSON_BufferPosition), _
                    ByVal StrPtr(JSON_Append), _
                    JSON_AppendLength
    
    JSON_BufferPosition = JSON_BufferPosition + JSON_AppendLength
#End If
End Sub

Private Function JSON_BufferToString(ByRef json_buffer As String, ByVal JSON_BufferPosition As Long, ByVal JSON_BufferLength As Long) As String
#If Mac Then
    JSON_BufferToString = json_buffer
#Else
    If JSON_BufferPosition > 0 Then
        JSON_BufferToString = VBA.Left$(json_buffer, JSON_BufferPosition \ 2)
    End If
#End If
End Function

#If Win64 Then
Private Function JSON_UnsignedAdd(JSON_Start As LongPtr, JSON_Increment As Long) As LongPtr
#Else
Private Function JSON_UnsignedAdd(JSON_Start As Long, JSON_Increment As Long) As Long
#End If

    If JSON_Start And &H80000000 Then
        JSON_UnsignedAdd = JSON_Start + JSON_Increment
    ElseIf (JSON_Start Or &H80000000) < -JSON_Increment Then
        JSON_UnsignedAdd = JSON_Start + JSON_Increment
    Else
        JSON_UnsignedAdd = (JSON_Start + &H80000000) + (JSON_Increment + &H80000000)
    End If
End Function

''
' VBA-UTC v1.0.0-rc.1
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
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Moved to top)
'#If Mac Then
'
'Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" (ByVal utc_command As String, ByVal utc_mode As String) As Long
'Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" (ByVal utc_file As Long) As Long
'Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" (ByVal utc_buffer As String, ByVal utc_size As Long, ByVal utc_number As Long, ByVal utc_file As Long) As Long
'Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" (ByVal utc_file As Long) As Long
'
'Private Type utc_ShellResult
'    utc_Output As String
'    utc_ExitCode As Long
'End Type
'
'#Else
'
'' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
'' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
'' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
'Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
'    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
'Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
'    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
'Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
'    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long
'
'Private Type utc_SYSTEMTIME
'    utc_wYear As Integer
'    utc_wMonth As Integer
'    utc_wDayOfWeek As Integer
'    utc_wDay As Integer
'    utc_wHour As Integer
'    utc_wMinute As Integer
'    utc_wSecond As Integer
'    utc_wMilliseconds As Integer
'End Type
'
'Private Type utc_TIME_ZONE_INFORMATION
'    utc_Bias As Long
'    utc_StandardName(0 To 31) As Integer
'    utc_StandardDate As utc_SYSTEMTIME
'    utc_StandardBias As Long
'    utc_DaylightName(0 To 31) As Integer
'    utc_DaylightDate As utc_SYSTEMTIME
'    utc_DaylightBias As Long
'End Type
'
'#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @param {Date} utc_UtcDate
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo ErrorHandling
    
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

ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' -------------------------------------- '
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo ErrorHandling
    
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
    
ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @param {Date} utc_IsoString
' @return {Date} Local date
' -------------------------------------- '
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo ErrorHandling
    
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
    
ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' -------------------------------------- '
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo ErrorHandling
    
    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")
    
    Exit Function
    
ErrorHandling:
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
    Dim utc_file As Long
    Dim utc_Chunk As String
    Dim utc_Read As Long
    
    On Error GoTo ErrorHandling
    utc_file = utc_popen(utc_ShellCommand, "r")
    
    If utc_file = 0 Then: Exit Function
    
    Do While utc_feof(utc_file) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_file)
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, utc_Read)
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = utc_pclose(File)
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


Attribute VB_Name = "JSONConverter"
''
' VBA-JSONConverter v0.5.0
' (c) Tim Hall - https://github.com/timhall/VBA-JSONConverter
'
' JSON Converter for VBA
'
' Errors (513-65535 available):
' 10001 - JSON parse error
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
    
    ' Remove vbCrLf, vbCr, vbLf, and vbTab from JSON_String
    JSON_String = Replace(Replace(Replace(Replace(JSON_String, vbCrLf, ""), vbCr, ""), vbLf, ""), vbTab, "")
    
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case Mid$(JSON_String, JSON_Index, 1)
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
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    Dim JSON_Index As Long
    Dim JSON_LBound As Long
    Dim JSON_UBound As Long
    Dim JSON_IsFirstItem As Boolean
    Dim JSON_Key As Variant
    Dim JSON_Value As Variant
    JSON_IsFirstItem = True

    Select Case VarType(JSON_DictionaryCollectionOrArray)
    Case vbNull, vbEmpty
        ConvertToJSON = "null"
    Case vbDate
        ' TODO Verify date formatting
        ConvertToJSON = """" & CStr(JSON_DictionaryCollectionOrArray) & """"
    Case vbString
        ' String (or large number encoded as string)
        If JSON_ConvertLargeNumbersFromString And JSON_StringIsLargeNumber(JSON_DictionaryCollectionOrArray) Then
            ConvertToJSON = JSON_DictionaryCollectionOrArray
        Else
            ConvertToJSON = """" & JSON_Encode(JSON_DictionaryCollectionOrArray) & """"
        End If
    Case vbBoolean
        If JSON_DictionaryCollectionOrArray Then
            ConvertToJSON = "true"
        Else
            ConvertToJSON = "false"
        End If
    
    Case vbVariant, vbArray, vbArray + vbVariant
        ' Array
        JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
        
        On Error Resume Next
        
        JSON_LBound = LBound(JSON_DictionaryCollectionOrArray)
        JSON_UBound = UBound(JSON_DictionaryCollectionOrArray)
        
        If JSON_LBound >= 0 And JSON_UBound >= 0 Then
            For JSON_Index = JSON_LBound To JSON_UBound
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
                
                JSON_BufferAppend JSON_Buffer, _
                    ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index), JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Index
        End If
        
        On Error GoTo 0
        
        JSON_BufferAppend JSON_Buffer, "]", JSON_BufferPosition, JSON_BufferLength
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
    
    ' Dictionary or Collection
    Case vbObject
        ' Dictionary
        If TypeName(JSON_DictionaryCollectionOrArray) = "Dictionary" Then
            JSON_BufferAppend JSON_Buffer, "{", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Key In JSON_DictionaryCollectionOrArray.Keys
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend JSON_Buffer, _
                    """" & JSON_Key & """:" & ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Key), JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Key
            JSON_BufferAppend JSON_Buffer, "}", JSON_BufferPosition, JSON_BufferLength
        
        ' Collection
        ElseIf TypeName(JSON_DictionaryCollectionOrArray) = "Collection" Then
            JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
            For Each JSON_Value In JSON_DictionaryCollectionOrArray
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                End If
            
                JSON_BufferAppend JSON_Buffer, _
                    ConvertToJSON(JSON_Value, JSON_ConvertLargeNumbersFromString), _
                    JSON_BufferPosition, JSON_BufferLength
            Next JSON_Value
            JSON_BufferAppend JSON_Buffer, "]", JSON_BufferPosition, JSON_BufferLength
        End If
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
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
    If Mid$(JSON_String, JSON_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{'")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If "}" = Mid$(JSON_String, JSON_Index, 1) Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf "," = Mid$(JSON_String, JSON_Index, 1) Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_Key = JSON_ParseKey(JSON_String, JSON_Index)
            JSON_NextChar = JSON_Peek(JSON_String, JSON_Index)
            If "{" = JSON_NextChar Or "[" = JSON_NextChar Then
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
    If Mid$(JSON_String, JSON_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '['")
    Else
        JSON_Index = JSON_Index + 1
        
        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If "]" = Mid$(JSON_String, JSON_Index, 1) Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf "," = Mid$(JSON_String, JSON_Index, 1) Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If
            
            JSON_ParseArray.Add JSON_ParseValue(JSON_String, JSON_Index, JSON_ConvertLargeNumbersToString)
        Loop
    End If
End Function

Private Function JSON_ParseValue(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set JSON_ParseValue = JSON_ParseObject(JSON_String, JSON_Index)
    Case "["
        Set JSON_ParseValue = JSON_ParseArray(JSON_String, JSON_Index)
    Case """", "'"
        JSON_ParseValue = JSON_ParseString(JSON_String, JSON_Index)
    Case Else
        If Mid$(JSON_String, JSON_Index, 4) = "true" Then
            JSON_ParseValue = True
            JSON_Index = JSON_Index + 4
        ElseIf Mid$(JSON_String, JSON_Index, 5) = "false" Then
            JSON_ParseValue = False
            JSON_Index = JSON_Index + 5
        ElseIf Mid$(JSON_String, JSON_Index, 4) = "null" Then
            JSON_ParseValue = Null
            JSON_Index = JSON_Index + 4
        ElseIf InStr("0123456789", Mid$(JSON_String, JSON_Index, 1)) Then
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
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    ' Store opening quote to look for matching closing quote
    JSON_Quote = Mid$(JSON_String, JSON_Index, 1)
    JSON_Index = JSON_Index + 1
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = Mid$(JSON_String, JSON_Index, 1)
        
        Select Case JSON_Char
        Case "\"
            ' Escaped string, \\, or \/
            JSON_Index = JSON_Index + 1
            JSON_Char = Mid$(JSON_String, JSON_Index, 1)
            
            Select Case JSON_Char
            Case """", "\", "/", "'"
                JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "b"
                JSON_BufferAppend JSON_Buffer, vbBack, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "f"
                JSON_BufferAppend JSON_Buffer, vbFormFeed, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "n"
                JSON_BufferAppend JSON_Buffer, vbCrLf, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "r"
                JSON_BufferAppend JSON_Buffer, vbCr, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "t"
                JSON_BufferAppend JSON_Buffer, vbTab, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                JSON_Index = JSON_Index + 1
                JSON_Code = Mid$(JSON_String, JSON_Index, 4)
                JSON_BufferAppend JSON_Buffer, ChrW(Val("&h" + JSON_Code)), JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 4
            End Select
        Case JSON_Quote
            JSON_ParseString = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition, JSON_BufferLength)
            JSON_Index = JSON_Index + 1
            Exit Function
        Case Else
            JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
            JSON_Index = JSON_Index + 1
        End Select
    Loop
End Function

Private Function JSON_ParseNumber(JSON_String As String, ByRef JSON_Index As Long, Optional JSON_ConvertLargeNumbersToString As Boolean = True) As Variant
    Dim JSON_Char As String
    Dim JSON_Value As String
    
    JSON_SkipSpaces JSON_String, JSON_Index
    
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = Mid$(JSON_String, JSON_Index, 1)
        
        If InStr("+-0123456789.eE", JSON_Char) Then
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
                JSON_ParseNumber = Val(JSON_Value)
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
    If Mid$(JSON_String, JSON_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting ':'")
    Else
        JSON_Index = JSON_Index + 1
    End If
End Function

Private Function JSON_Encode(ByVal JSON_Text As Variant) As String
    ' TODO
    JSON_Encode = JSON_Text
End Function

Private Function JSON_Peek(JSON_String As String, ByVal JSON_Index As Long, Optional JSON_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing JSON_Index (ByVal instead of ByRef)
    JSON_SkipSpaces JSON_String, JSON_Index
    JSON_Peek = Mid$(JSON_String, JSON_Index, JSON_NumberOfCharacters)
End Function

Private Sub JSON_SkipSpaces(JSON_String As String, ByRef JSON_Index As Long)
    ' Increment index to skip over spaces
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String) And Mid$(JSON_String, JSON_Index, 1) = " "
        JSON_Index = JSON_Index + 1
    Loop
End Sub

Private Function JSON_StringIsLargeNumber(JSON_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See JSON_ParseNumber)
    
    Dim JSON_Length As Long
    JSON_Length = Len(JSON_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If JSON_Length >= 16 And JSON_Length <= 100 Then
        Dim JSON_CharCode As String
        Dim JSON_Index As Long
        
        JSON_StringIsLargeNumber = True
        
        For i = 1 To JSON_Length
            JSON_CharCode = Asc(Mid$(JSON_String, i, 1))
            Select Case JSON_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                JSON_StringIsLargeNumber = False
                Exit Function
            End Select
        Next i
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
    If JSON_StopIndex > Len(JSON_String) Then
        JSON_StopIndex = Len(JSON_String)
    End If

    JSON_ParseErrorMessage = "Error parsing JSON:" & vbNewLine & _
                             Mid$(JSON_String, JSON_StartIndex, JSON_StopIndex - JSON_StartIndex + 1) & vbNewLine & _
                             VBA.Space$(JSON_Index - JSON_StartIndex) & "^" & vbNewLine & _
                             ErrorMessage
End Function

Private Sub JSON_BufferAppend(ByRef JSON_Buffer As String, _
                              ByRef JSON_Append As String, _
                              ByRef JSON_BufferPosition As Long, _
                              ByRef JSON_BufferLength As Long)
#If Mac Then
    JSON_Buffer = JSON_Buffer & JSON_Append
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
    
    JSON_AppendLength = LenB(JSON_Append)
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
        
        JSON_Buffer = JSON_Buffer & Space$((JSON_TemporaryLength - JSON_BufferLength) \ 2)
        JSON_BufferLength = JSON_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    JSON_CopyMemory ByVal JSON_UnsignedAdd(StrPtr(JSON_Buffer), _
                    JSON_BufferPosition), _
                    ByVal StrPtr(JSON_Append), _
                    JSON_AppendLength
    
    JSON_BufferPosition = JSON_BufferPosition + JSON_AppendLength
#End If
End Sub

Private Function JSON_BufferToString(ByRef JSON_Buffer As String, ByVal JSON_BufferPosition As Long, ByVal JSON_BufferLength As Long) As String
#If Mac Then
    JSON_BufferToString = JSON_Buffer
#Else
    If JSON_BufferPosition > 0 Then
        JSON_BufferToString = Left$(JSON_Buffer, JSON_BufferPosition \ 2)
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

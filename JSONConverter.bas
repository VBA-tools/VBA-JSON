Attribute VB_Name = "JSONConverter"
''
' VBA-JSONConverter v0.0.0
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

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @param {String} JSON_String
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function ParseJSON(ByVal JSON_String As String) As Object
    Dim JSON_Index As Long
    JSON_Index = 1
    
    ' Remove vbCrLf, vbCr, vbLf, and vbTab from JSON_String
    JSON_String = Replace(Replace(Replace(Replace(JSON_String, vbCrLf, ""), vbCr, ""), vbLf, ""), vbTab, "")
    
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case Mid$(JSON_String, JSON_Index, 1)
    Case "{"
        Set ParseJSON = JSON_ParseObject(JSON_String, JSON_Index)
    Case "["
        Set ParseJSON = JSON_ParseArray(JSON_String, JSON_Index)
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
Public Function ConvertToJSON(ByVal JSON_DictionaryCollectionOrArray As Variant) As String
    Dim JSON_Buffer As String
    Dim JSON_Index As Long
    Dim JSON_LBound As Long
    Dim JSON_UBound As Long
    Dim JSON_IsFirstItem As Boolean
    Dim JSON_Key As Variant
    Dim JSON_Value As Variant
    JSON_IsFirstItem = True

    Select Case VarType(JSON_DictionaryCollectionOrArray)
    Case vbNull
        ConvertToJSON = "null"
    Case vbEmpty
        ConvertToJSON = "null"
    Case vbDate
        ConvertToJSON = """" & CStr(JSON_DictionaryCollectionOrArray) & """"
    Case vbString
        ConvertToJSON = """" & JSON_Encode(JSON_DictionaryCollectionOrArray) & """"
    Case vbBoolean
        If JSON_DictionaryCollectionOrArray Then
            ConvertToJSON = "true"
        Else
            ConvertToJSON = "false"
        End If
    Case vbVariant, vbArray, vbArray + vbVariant
        On Error Resume Next
        JSON_LBound = LBound(JSON_DictionaryCollectionOrArray)
        JSON_UBound = UBound(JSON_DictionaryCollectionOrArray)
        On Error GoTo 0
        
        JSON_BufferAppend JSON_Buffer, "["
        If JSON_LBound >= 0 And JSON_UBound >= 0 Then
            For JSON_Index = JSON_LBound To JSON_UBound
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ","
                End If
                
                JSON_BufferAppend JSON_Buffer, ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Index))
            Next JSON_Index
        End If
        JSON_BufferAppend JSON_Buffer, "]"
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer)
    Case vbObject
        If TypeName(JSON_DictionaryCollectionOrArray) = "Dictionary" Then
            JSON_BufferAppend JSON_Buffer, "{"
            For Each JSON_Key In JSON_DictionaryCollectionOrArray.Keys
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ","
                End If
            
                JSON_BufferAppend JSON_Buffer, """" & JSON_Key & """:" & ConvertToJSON(JSON_DictionaryCollectionOrArray(JSON_Key))
            Next JSON_Key
            JSON_BufferAppend JSON_Buffer, "}"
        ElseIf TypeName(JSON_DictionaryCollectionOrArray) = "Collection" Then
            JSON_BufferAppend JSON_Buffer, "["
            For Each JSON_Value In JSON_DictionaryCollectionOrArray
                If JSON_IsFirstItem Then
                    JSON_IsFirstItem = False
                Else
                    JSON_BufferAppend JSON_Buffer, ","
                End If
            
                JSON_BufferAppend JSON_Buffer, ConvertToJSON(JSON_Value)
            Next JSON_Value
            JSON_BufferAppend JSON_Buffer, "]"
        End If
        
        ConvertToJSON = JSON_BufferToString(JSON_Buffer)
    Case Else
        ' Number
        ' Note: For cases where commas are used in place of decimals, replace with decimals
        ConvertToJSON = Replace(JSON_DictionaryCollectionOrArray, ",", ".")
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function JSON_ParseObject(JSON_String As String, ByRef JSON_Index As Long) As Dictionary
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
                Set JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index)
            Else
                JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index)
            End If
        Loop
    End If
End Function

Private Function JSON_ParseArray(JSON_String As String, ByRef JSON_Index As Long) As Collection
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
            
            JSON_ParseArray.Add JSON_ParseValue(JSON_String, JSON_Index)
        Loop
    End If
End Function

Private Function JSON_ParseValue(JSON_String As String, ByRef JSON_Index As Long) As Variant
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
            JSON_ParseValue = JSON_ParseNumber(JSON_String, JSON_Index)
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
            Case """", "\", "/"
                JSON_BufferAppend JSON_Buffer, JSON_Char
                JSON_Index = JSON_Index + 1
            Case "b"
                JSON_BufferAppend JSON_Buffer, vbBack
                JSON_Index = JSON_Index + 1
            Case "f"
                JSON_BufferAppend JSON_Buffer, vbFormFeed
                JSON_Index = JSON_Index + 1
            Case "n"
                JSON_BufferAppend JSON_Buffer, vbCrLf
                JSON_Index = JSON_Index + 1
            Case "r"
                JSON_BufferAppend JSON_Buffer, vbCr
                JSON_Index = JSON_Index + 1
            Case "t"
                JSON_BufferAppend JSON_Buffer, vbTab
                JSON_Index = JSON_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                JSON_Index = JSON_Index + 1
                JSON_Code = Mid$(JSON_String, JSON_Index, 4)
                JSON_BufferAppend JSON_Buffer, ChrW(Val("&h" + JSON_Code))
                JSON_Index = JSON_Index + 4
            End Select
        Case JSON_Quote
            JSON_ParseString = JSON_BufferToString(JSON_Buffer)
            JSON_Index = JSON_Index + 1
            Exit Function
        Case Else
            JSON_BufferAppend JSON_Buffer, JSON_Char
            JSON_Index = JSON_Index + 1
        End Select
    Loop
End Function

Private Function JSON_ParseNumber(JSON_String As String, ByRef JSON_Index As Long) As Variant
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
            ' TODO Look into more detailed number parsing
            JSON_ParseNumber = Val(JSON_Value)
            Exit Function
        End If
    Loop
End Function

Private Function JSON_ParseKey(JSON_String As String, ByRef JSON_Index As Long) As String
    JSON_ParseKey = JSON_ParseString(JSON_String, JSON_Index)
    
    JSON_SkipSpaces JSON_String, JSON_Index
    If Mid$(JSON_String, JSON_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting ':'")
    Else
        JSON_Index = JSON_Index + 1
    End If
End Function

Private Function JSON_Encode(ByVal JSON_Text As Variant) As String
    JSON_Encode = JSON_Text
End Function

Private Function JSON_Peek(JSON_String As String, ByVal JSON_Index As Long, Optional JSON_NumberOfCharacters As Long = 1) As String
    JSON_SkipSpaces JSON_String, JSON_Index
    JSON_Peek = Mid$(JSON_String, JSON_Index, JSON_NumberOfCharacters)
End Function

Private Sub JSON_SkipSpaces(JSON_String As String, ByRef JSON_Index As Long)
    ' Increment index to skip over spaces
    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String) And Mid$(JSON_String, JSON_Index, 1) = " "
        JSON_Index = JSON_Index + 1
    Loop
End Sub

Private Sub JSON_BufferAppend(ByRef JSON_Buffer As String, ByRef JSON_Append As String)
    ' TODO Convert to direct memory buffer
    JSON_Buffer = JSON_Buffer & JSON_Append
End Sub

Private Function JSON_BufferToString(ByRef JSON_Buffer As String) As String
    ' TODO Convert to direct memory buffer
    JSON_BufferToString = JSON_Buffer
End Function

Private Function JSON_ParseErrorMessage(JSON_String As String, ByRef JSON_Index As Long, ErrorMessage As String)
    Dim JSON_StartIndex As Long
    Dim JSON_StopIndex As Long
    
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

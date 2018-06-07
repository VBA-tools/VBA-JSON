Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-JSON"
    
    On Error Resume Next
    
    Dim JsonString As String
    Dim JsonObject As Object
    Dim NestedObject As Object
    Dim EmptyVariant As Variant
    Dim NothingObject As Object
    
    Dim MultiDimensionalArray(1, 1) As Variant
    
    ' ============================================= '
    ' ParseJson
    ' ============================================= '
    
    With Specs.It("should parse object string")
        JsonString = "{""a"":1,""b"":3.14,""c"":""abc"",""d"":false,""e"":[1,3.14,""abc"",false,[1,2,3],{""a"":1}],""f"":{""a"":1},""g"":null}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject("a")).ToEqual 1
        .Expect(JsonObject("b")).ToEqual 3.14
        .Expect(JsonObject("c")).ToEqual "abc"
        .Expect(JsonObject("d")).ToEqual False
        
        .Expect(JsonObject("e")).ToNotBeUndefined
        .Expect(JsonObject("e")(1)).ToEqual 1
        .Expect(JsonObject("e")(2)).ToEqual 3.14
        .Expect(JsonObject("e")(3)).ToEqual "abc"
        .Expect(JsonObject("e")(4)).ToEqual False
        .Expect(JsonObject("e")(5)).ToNotBeUndefined
        .Expect(JsonObject("e")(5)(1)).ToEqual 1
        .Expect(JsonObject("e")(5)(2)).ToEqual 2
        .Expect(JsonObject("e")(5)(3)).ToEqual 3
        .Expect(JsonObject("e")(6)).ToNotBeUndefined
        .Expect(JsonObject("e")(6)("a")).ToEqual 1
        
        .Expect(JsonObject("f")).ToNotBeUndefined
        .Expect(JsonObject("f")("a")).ToEqual 1
        
        .Expect(JsonObject("g")).ToBeNull
    End With
    
    With Specs.It("should parse array string")
        JsonString = "[1,3.14,""abc"",false,[1,2,3],{""a"":1},null]"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject(1)).ToEqual 1
        .Expect(JsonObject(2)).ToEqual 3.14
        .Expect(JsonObject(3)).ToEqual "abc"
        .Expect(JsonObject(4)).ToEqual False
        .Expect(JsonObject(5)).ToNotBeUndefined
        .Expect(JsonObject(5)(1)).ToEqual 1
        .Expect(JsonObject(5)(2)).ToEqual 2
        .Expect(JsonObject(5)(3)).ToEqual 3
        .Expect(JsonObject(6)).ToNotBeUndefined
        .Expect(JsonObject(6)("a")).ToEqual 1
        .Expect(JsonObject(7)).ToBeNull
    End With
    
    With Specs.It("should parse nested array string")
        JsonString = "[[[1,2,3],4],5]"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject(1)).ToNotBeUndefined
        .Expect(JsonObject(1)(1)).ToNotBeUndefined
        .Expect(JsonObject(1)(1)(1)).ToEqual 1
        .Expect(JsonObject(1)(1)(2)).ToEqual 2
        .Expect(JsonObject(1)(1)(3)).ToEqual 3
        .Expect(JsonObject(1)(2)).ToEqual 4
        .Expect(JsonObject(2)).ToEqual 5
    End With
    
    With Specs.It("should parse escaped single quote in key and value")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=2
        JsonString = "{'a\'b':'c\'d'}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject.Exists("a'b")).ToEqual True
        .Expect(JsonObject("a'b")).ToEqual "c'd"
    End With
    
    With Specs.It("should parse nested objects and arrays")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=7
        JsonString = "{""total_rows"":36778,""offset"":26220,""rows"":[" & vbNewLine & _
            "{""id"":""6b80c0b76"",""key"":""a@bbb.net"",""value"":{""entryid"":""81151F241C2500"",""subject"":""test subject"",""senton"":""2009-7-09 22:03:43""}}," & vbNewLine & _
            "{""id"":""b10ed9bee"",""key"":""b@bbb.net"",""value"":{""entryid"":""A7C3CF74EA95C9F"",""subject"":""test subject2"",""senton"":""2009-4-21 10:18:26""}}" & vbNewLine & _
        "]}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject("offset")).ToEqual 26220
        .Expect(JsonObject("rows")(2)("key")).ToEqual "b@bbb.net"
    End With
    
    With Specs.It("should handle very long numbers as strings (e.g. BIGINT)")
        JsonString = "[123456789012345678901234567890, 1.123456789012345678901234567890, 123456789012345, 1.23456789012345]"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject(1)).ToEqual "123456789012345678901234567890"
        .Expect(JsonObject(2)).ToEqual "1.123456789012345678901234567890"
        .Expect(JsonObject(3)).ToEqual 123456789012345#
        .Expect(JsonObject(4)).ToEqual 1.23456789012345
        
        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
        JsonString = "[123456789012345678901234567890]"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject(1)).ToEqual 1.23456789012346E+29
        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = False
    End With
    
    With Specs.It("should parse double-backslash as backslash")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=11
        JsonString = "[""C:\\folder\\picture.jpg""]"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject(1)).ToEqual "C:\folder\picture.jpg"
    End With
    
    With Specs.It("should allow keys and values with colons")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=14
        JsonString = "{""a:b"":""c:d""}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject.Exists("a:b")).ToEqual True
        .Expect(JsonObject("a:b")).ToEqual "c:d"
    End With
    
    With Specs.It("should allow spaces in keys")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=19
        JsonString = "{""a b  c"":""d e  f""}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject.Exists("a b  c")).ToEqual True
        .Expect(JsonObject("a b  c")).ToEqual "d e  f"
    End With
    
    With Specs.It("should allow unquoted keys with option")
        JsonConverter.JsonOptions.AllowUnquotedKeys = True
        JsonString = "{a:""a"",b     :""b""}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)

        .Expect(JsonObject).ToNotBeUndefined
        .Expect(JsonObject.Exists("a")).ToEqual True
        .Expect(JsonObject("a")).ToEqual "a"
        .Expect(JsonObject.Exists("b")).ToEqual True
        .Expect(JsonObject("b")).ToEqual "b"
        JsonConverter.JsonOptions.AllowUnquotedKeys = False
    End With
    
    ' ============================================= '
    ' ConvertToJson
    ' ============================================= '
    
    With Specs.It("should convert object to string")
        Set JsonObject = New Dictionary
        JsonObject.Add "a", 1
        JsonObject.Add "b", 3.14
        JsonObject.Add "c", "abc"
        JsonObject.Add "d", False
        JsonObject.Add "e", New Collection
        JsonObject("e").Add 1
        JsonObject("e").Add 3.14
        JsonObject("e").Add "abc"
        JsonObject("e").Add False
        JsonObject("e").Add Array(1, 2, 3)
        JsonObject("e").Add New Dictionary
        JsonObject("e")(6).Add "a", 1
        JsonObject.Add "f", New Dictionary
        JsonObject("f").Add "a", 1
        JsonObject.Add "g", Null
        
        JsonString = JsonConverter.ConvertToJson(JsonObject)
        .Expect(JsonString).ToEqual "{""a"":1,""b"":3.14,""c"":""abc"",""d"":false,""e"":[1,3.14,""abc"",false,[1,2,3],{""a"":1}],""f"":{""a"":1},""g"":null}"
    End With
    
    With Specs.It("should convert collection to string")
        Set JsonObject = New Collection
        JsonObject.Add 1
        JsonObject.Add 3.14
        JsonObject.Add "abc"
        JsonObject.Add False
        JsonObject.Add Array(1, 2, 3)
        JsonObject.Add New Dictionary
        JsonObject(6).Add "a", 1
        JsonObject.Add Null
    
        JsonString = JsonConverter.ConvertToJson(JsonObject)
        .Expect(JsonString).ToEqual "[1,3.14,""abc"",false,[1,2,3],{""a"":1},null]"
    End With
    
    With Specs.It("should convert array to string")
        JsonString = JsonConverter.ConvertToJson(Array(1, 3.14, "abc", False, Array(1, 2, 3)))
        .Expect(JsonString).ToEqual "[1,3.14,""abc"",false,[1,2,3]]"
    End With
    
    With Specs.It("should convert very long numbers as strings (e.g. BIGINT)")
        JsonString = JsonConverter.ConvertToJson(Array("123456789012345678901234567890", "1.123456789012345678901234567890", "1234567890123456F"))
        .Expect(JsonString).ToEqual "[123456789012345678901234567890,1.123456789012345678901234567890,""1234567890123456F""]"
        
        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
        JsonString = JsonConverter.ConvertToJson(Array("123456789012345678901234567890"))
        .Expect(JsonString).ToEqual "[""123456789012345678901234567890""]"
        JsonConverter.JsonOptions.UseDoubleForLargeNumbers = False
    End With
    
    With Specs.It("should convert dates to ISO 8601")
        JsonString = JsonConverter.ConvertToJson(DateSerial(2003, 1, 15) + TimeSerial(12, 5, 6))
        
        ' Due to UTC conversion, test shape and year, month, and seconds
        .Expect(JsonString).ToMatch "2003-01-"
        .Expect(JsonString).ToMatch "T"
        .Expect(JsonString).ToMatch ":06.000Z"
    End With
    
    With Specs.It("should convert 2D arrays")
        ' Checks https://code.google.com/p/vba-json/issues/detail?id=8
        MultiDimensionalArray(0, 0) = 1
        MultiDimensionalArray(0, 1) = 2
        MultiDimensionalArray(1, 0) = 3
        MultiDimensionalArray(1, 1) = 4
        JsonString = JsonConverter.ConvertToJson(MultiDimensionalArray)
        .Expect(JsonString).ToEqual "[[1,2],[3,4]]"
    End With
    
    With Specs.It("should handle strongly typed arrays")
        Dim LongArray(3) As Long
        LongArray(0) = 1
        LongArray(1) = 2
        LongArray(2) = 3
        LongArray(3) = 4
        
        JsonString = JsonConverter.ConvertToJson(LongArray)
        .Expect(JsonString).ToEqual "[1,2,3,4]"
        
        Dim StringArray(3) As String
        StringArray(0) = "A"
        StringArray(1) = "B"
        StringArray(2) = "C"
        StringArray(3) = "D"
        
        JsonString = JsonConverter.ConvertToJson(StringArray)
        .Expect(JsonString).ToEqual "[""A"",""B"",""C"",""D""]"
    End With
    
    With Specs.It("should json-encode strings")
        Dim Strings As Variant
        Strings = Array("""\" & vbCr & vbLf & vbTab & vbBack & vbFormFeed, ChrW(128) & ChrW(32767), "#$%&{|}~")
        
        JsonString = JsonConverter.ConvertToJson(Strings)
        .Expect(JsonString).ToEqual "[""\""\\\r\n\t\b\f"",""\u0080\u7FFF"",""#$%&{|}~""]"
    End With
    
    With Specs.It("should escape solidus with option")
        Strings = Array("a/b")
        
        JsonString = JsonConverter.ConvertToJson(Strings)
        .Expect(JsonString).ToEqual "[""a/b""]"
        
        JsonConverter.JsonOptions.EscapeSolidus = True
        JsonString = JsonConverter.ConvertToJson(Strings)
        .Expect(JsonString).ToEqual "[""a\/b""]"
        JsonConverter.JsonOptions.EscapeSolidus = False
    End With
    
    With Specs.It("should handle Empty and Nothing in arrays as null")
        JsonString = JsonConverter.ConvertToJson(Array("a", EmptyVariant, NothingObject, Empty, Nothing, "z"))
        .Expect(JsonString).ToEqual "[""a"",null,null,null,null,""z""]"

        Set JsonObject = New Collection
        JsonObject.Add "a"
        JsonObject.Add EmptyVariant
        JsonObject.Add NothingObject
        JsonObject.Add Empty
        JsonObject.Add Nothing
        JsonObject.Add ""
        JsonObject.Add "z"
    
        JsonString = JsonConverter.ConvertToJson(JsonObject)
        .Expect(JsonString).ToEqual "[""a"",null,null,null,null,"""",""z""]"
    End With
    
    With Specs.It("should handle Empty and Nothing in objects as undefined")
        Set JsonObject = New Dictionary
        JsonObject.Add "a", "a"
        JsonObject.Add "b", EmptyVariant
        JsonObject.Add "c", NothingObject
        JsonObject.Add "d", Empty
        JsonObject.Add "e", Nothing
        JsonObject.Add "f", ""
        JsonObject.Add "z", "z"
        
        JsonString = JsonConverter.ConvertToJson(JsonObject)
        .Expect(JsonString).ToEqual "{""a"":""a"",""f"":"""",""z"":""z""}"
    End With
    
    With Specs.It("should use whitespace number/string")
        ' Nested, plain array + 2
        JsonString = JsonConverter.ConvertToJson(Array(1, Array(2, Array(3))), 2)
        .Expect(JsonString).ToEqual _
            "[" & vbNewLine & _
            "  1," & vbNewLine & _
            "  [" & vbNewLine & _
            "    2," & vbNewLine & _
            "    [" & vbNewLine & _
            "      3" & vbNewLine & _
            "    ]" & vbNewLine & _
            "  ]" & vbNewLine & _
            "]"
        
        ' Nested Dictionary + Tab
        Set JsonObject = New Dictionary
        JsonObject.Add "a", Array(1, 2, 3)
        JsonObject.Add "b", "c"
        Set NestedObject = New Dictionary
        NestedObject.Add "d", "e"
        JsonObject.Add "nested", NestedObject
        
        JsonString = JsonConverter.ConvertToJson(JsonObject, VBA.vbTab)
        .Expect(JsonString).ToEqual _
            "{" & vbNewLine & _
            vbTab & """a"": [" & vbNewLine & _
            vbTab & vbTab & "1," & vbNewLine & _
            vbTab & vbTab & "2," & vbNewLine & _
            vbTab & vbTab & "3" & vbNewLine & _
            vbTab & "]," & vbNewLine & _
            vbTab & """b"": ""c""," & vbNewLine & _
            vbTab & """nested"": {" & vbNewLine & _
            vbTab & vbTab & """d"": ""e""" & vbNewLine & _
            vbTab & "}" & vbNewLine & _
            "}"
            
        ' Multi-dimensional array + 4
        MultiDimensionalArray(0, 0) = 1
        MultiDimensionalArray(0, 1) = 2
        MultiDimensionalArray(1, 0) = Array(1, 2, 3)
        MultiDimensionalArray(1, 1) = 4
        JsonString = JsonConverter.ConvertToJson(MultiDimensionalArray, 4)
        .Expect(JsonString).ToEqual _
            "[" & vbNewLine & _
            "    [" & vbNewLine & _
            "        1," & vbNewLine & _
            "        2" & vbNewLine & _
            "    ]," & vbNewLine & _
            "    [" & vbNewLine & _
            "        [" & vbNewLine & _
            "            1," & vbNewLine & _
            "            2," & vbNewLine & _
            "            3" & vbNewLine & _
            "        ]," & vbNewLine & _
            "        4" & vbNewLine & _
            "    ]" & vbNewLine & _
            "]"
        
        ' Collection + "-"
        Set JsonObject = New Collection
        JsonObject.Add Array(1, 2, 3)
        
        JsonString = JsonConverter.ConvertToJson(JsonObject, "-")
        .Expect(JsonString).ToEqual _
            "[" & vbNewLine & _
            "-[" & vbNewLine & _
            "--1," & vbNewLine & _
            "--2," & vbNewLine & _
            "--3" & vbNewLine & _
            "-]" & vbNewLine & _
            "]"
    End With
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    
    With Specs.It("should have descriptive parsing errors")
        Err.Clear
        JsonString = "Howdy!"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", _
            "Howdy!", "^", "Expecting '{' or '['"
        
        Err.Clear
        JsonString = "{""abc""}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", _
            "{""abc""}", "      ^", "Expecting ':'"
        
        Err.Clear
        JsonString = "{""abc"":True}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", _
            "{""abc"":True}", "       ^", "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"
        
        Err.Clear
        JsonString = "{""abc"":undefined}"
        Set JsonObject = JsonConverter.ParseJson(JsonString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", _
            "{""abc"":undefined}", "       ^", "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"
    End With
    
    InlineRunner.RunSuite Specs
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite Specs
End Sub

Public Function ToMatchParseError(Actual As Variant, Args As Variant) As Variant
    Dim Partial As String
    Dim Arrow As String
    Dim Message As String
    Dim Description As String
    
    If UBound(Args) < 2 Then
        ToMatchParseError = "Need to pass expected partial, arrow, and message"""
    ElseIf Err.Number = 10001 Then
        Partial = Args(0)
        Arrow = Args(1)
        Message = Args(2)
        Description = "Error parsing JSON:" & vbNewLine & Partial & vbNewLine & Arrow & vbNewLine & Message
        
        Dim Parts As Variant
        Parts = Split(Err.Description, vbNewLine)
        
        If Parts(1) <> Partial Then
            ToMatchParseError = "Expected " & Parts(1) & " to equal " & Partial
        ElseIf Parts(2) <> Arrow Then
            ToMatchParseError = "Expected " & Parts(2) & " to equal " & Arrow
        ElseIf Parts(3) <> Message Then
            ToMatchParseError = "Expected " & Parts(3) & " to equal " & Message
        ElseIf Err.Description <> Description Then
            ToMatchParseError = "Expected " & Err.Description & " to equal " & Description
        Else
            ToMatchParseError = True
        End If
    Else
        ToMatchParseError = "Expected error number " & Err.Number & " to be 10001"
    End If
End Function

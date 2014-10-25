Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-JSONConverter"
    
    On Error Resume Next
    
    Dim JSONString As String
    Dim JSONObject As Object
    
    ' ParseJSON
    With Specs.It("should parse object string")
        JSONString = "{""a"":1,""b"":3.14,""c"":""abc"",""d"":false,""e"":[1,3.14,""abc"",false,[1,2,3],{""a"":1}],""f"":{""a"":1},""g"":null}"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect(JSONObject).ToNotBeUndefined
        .Expect(JSONObject("a")).ToEqual 1
        .Expect(JSONObject("b")).ToEqual 3.14
        .Expect(JSONObject("c")).ToEqual "abc"
        .Expect(JSONObject("d")).ToEqual False
        
        .Expect(JSONObject("e")).ToNotBeUndefined
        .Expect(JSONObject("e")(1)).ToEqual 1
        .Expect(JSONObject("e")(2)).ToEqual 3.14
        .Expect(JSONObject("e")(3)).ToEqual "abc"
        .Expect(JSONObject("e")(4)).ToEqual False
        .Expect(JSONObject("e")(5)).ToNotBeUndefined
        .Expect(JSONObject("e")(5)(1)).ToEqual 1
        .Expect(JSONObject("e")(5)(2)).ToEqual 2
        .Expect(JSONObject("e")(5)(3)).ToEqual 3
        .Expect(JSONObject("e")(6)).ToNotBeUndefined
        .Expect(JSONObject("e")(6)("a")).ToEqual 1
        
        .Expect(JSONObject("f")).ToNotBeUndefined
        .Expect(JSONObject("f")("a")).ToEqual 1
        
        .Expect(JSONObject("g")).ToBeNull
    End With
    
    With Specs.It("should parse array string")
        JSONString = "[1,3.14,""abc"",false,[1,2,3],{""a"":1},null]"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect(JSONObject).ToNotBeUndefined
        .Expect(JSONObject(1)).ToEqual 1
        .Expect(JSONObject(2)).ToEqual 3.14
        .Expect(JSONObject(3)).ToEqual "abc"
        .Expect(JSONObject(4)).ToEqual False
        .Expect(JSONObject(5)).ToNotBeUndefined
        .Expect(JSONObject(5)(1)).ToEqual 1
        .Expect(JSONObject(5)(2)).ToEqual 2
        .Expect(JSONObject(5)(3)).ToEqual 3
        .Expect(JSONObject(6)).ToNotBeUndefined
        .Expect(JSONObject(6)("a")).ToEqual 1
        .Expect(JSONObject(7)).ToBeNull
    End With
    
    With Specs.It("should parse nested array string")
        JSONString = "[[[1,2,3],4],5]"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect(JSONObject).ToNotBeUndefined
        .Expect(JSONObject(1)).ToNotBeUndefined
        .Expect(JSONObject(1)(1)).ToNotBeUndefined
        .Expect(JSONObject(1)(1)(1)).ToEqual 1
        .Expect(JSONObject(1)(1)(2)).ToEqual 2
        .Expect(JSONObject(1)(1)(3)).ToEqual 3
        .Expect(JSONObject(1)(2)).ToEqual 4
        .Expect(JSONObject(2)).ToEqual 5
    End With
    
    ' ConvertToJSON
    With Specs.It("should convert object to string")
        Set JSONObject = New Dictionary
        JSONObject.Add "a", 1
        JSONObject.Add "b", 3.14
        JSONObject.Add "c", "abc"
        JSONObject.Add "d", False
        JSONObject.Add "e", New Collection
        JSONObject("e").Add 1
        JSONObject("e").Add 3.14
        JSONObject("e").Add "abc"
        JSONObject("e").Add False
        JSONObject("e").Add Array(1, 2, 3)
        JSONObject("e").Add New Dictionary
        JSONObject("e")(6).Add "a", 1
        JSONObject.Add "f", New Dictionary
        JSONObject("f").Add "a", 1
        JSONObject.Add "g", Null
        
        JSONString = JSONConverter.ConvertToJSON(JSONObject)
        .Expect(JSONString).ToEqual "{""a"":1,""b"":3.14,""c"":""abc"",""d"":false,""e"":[1,3.14,""abc"",false,[1,2,3],{""a"":1}],""f"":{""a"":1},""g"":null}"
    End With
    
    With Specs.It("should convert collection to string")
        Set JSONObject = New Collection
        JSONObject.Add 1
        JSONObject.Add 3.14
        JSONObject.Add "abc"
        JSONObject.Add False
        JSONObject.Add Array(1, 2, 3)
        JSONObject.Add New Dictionary
        JSONObject(6).Add "a", 1
        JSONObject.Add Null
    
        JSONString = JSONConverter.ConvertToJSON(JSONObject)
        .Expect(JSONString).ToEqual "[1,3.14,""abc"",false,[1,2,3],{""a"":1},null]"
    End With
    
    With Specs.It("should convert array to string")
        JSONString = JSONConverter.ConvertToJSON(Array(1, 3.14, "abc", False, Array(1, 2, 3)))
        .Expect(JSONString).ToEqual "[1,3.14,""abc"",false,[1,2,3]]"
    End With
    
    ' Errors
    With Specs.It("should have descriptive parsing errors")
        Err.Clear
        JSONString = "Howdy!"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", "Howdy!", "^", "Expecting '{' or '['"
        
        Err.Clear
        JSONString = "{""abc""}"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", "{""abc""}", "      ^", "Expecting ':'"
        
        Err.Clear
        JSONString = "{""abc"":True}"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", "{""abc"":True}", "       ^", "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"
        
        Err.Clear
        JSONString = "{""abc"":undefined}"
        Set JSONObject = JSONConverter.ParseJSON(JSONString)
        
        .Expect.RunMatcher "ToMatchParseError", "to match parse error", "{""abc"":undefined}", "       ^", "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['"
    End With
    
    InlineRunner.RunSuite Specs
End Function

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

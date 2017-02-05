Attribute VB_Name = "Specs"
#If False Then 'using JsonSpec instead for this example , it may be identical, haven't taken the time to check...

Private pForDisplay As Boolean
Private pUseNative As Boolean

Public Sub SpeedTest()
    #If Mac Then
        ' Mac
        ExecuteSpeedTest CompareToNative:=False
    #Else
        ' Windows
        ExecuteSpeedTest CompareToNative:=True
    #End If
End Sub

Sub ToggleNative(Optional Enabled As Boolean = True)
    Dim code As CodeModule
    Dim Lines As Variant
    Dim i As Integer
    
    Set code = ThisWorkbook.VBProject.VBComponents("Dictionary").CodeModule
    Lines = Split(code.Lines(1, 50), vbNewLine)
    
    For i = 0 To UBound(Lines)
        If InStr(1, Lines(i), "#Const UseScriptingDictionaryIfAvailable") Then
            code.ReplaceLine i + 1, "#Const UseScriptingDictionaryIfAvailable = " & Enabled
            Exit Sub
        End If
    Next i
End Sub

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    pForDisplay = True
    DisplayRunner.RunSuite Specs()
    pForDisplay = False
End Sub

Public Function Specs() As SpecSuite
    Dim UseNative As Boolean

#If Mac Then
    UseNative = False
#Else
    If pUseNative Then
        UseNative = True
        pUseNative = False
    Else
        If Not pForDisplay Then
            ' Run native specs first
            pUseNative = True
            Specs
        End If
        
        UseNative = False
    End If
#End If

    Set Specs = New SpecSuite
    If UseNative Then
        Specs.Description = "Scripting.Dictionary"
    Else
        Specs.Description = "VBA-Dictionary"
    End If
    
    Dim dict As Object
    Dim Items As Variant
    Dim Keys As Variant
    Dim Key As Variant
    Dim Item As Variant
    Dim A As New Collection
    Dim B As New Dictionary
    
    ' Properties
    ' ------------------------- '
    With Specs.It("should get count of items")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        .Expect(dict.Count).ToEqual 3
        
        dict.Remove "C"
        .Expect(dict.Count).ToEqual 2
    End With
    
    With Specs.It("should get item by key")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        
        .Expect(dict.Item("B")).ToEqual 3.14
        .Expect(dict.Item("D")).ToBeEmpty
        .Expect(dict("B")).ToEqual 3.14
        .Expect(dict("D")).ToBeEmpty
    End With
    
    With Specs.It("should let item by key")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        
        ' Let + New
        dict("D") = True
        
        ' Let + Replace
        dict("A") = 456
        dict("B") = 3.14159
        
        ' Should have correct values
        .Expect(dict("A")).ToEqual 456
        .Expect(dict("B")).ToEqual 3.14159
        .Expect(dict("C")).ToEqual "ABC"
        .Expect(dict("D")).ToEqual True
        
        ' Should have correct order
        .Expect(dict.Keys()(0)).ToEqual "A"
        .Expect(dict.Keys()(1)).ToEqual "B"
        .Expect(dict.Keys()(2)).ToEqual "C"
        .Expect(dict.Keys()(3)).ToEqual "D"
    End With
    
    With Specs.It("should set item by key")
        Set dict = CreateDictionary(UseNative)

        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"

        ' Set + New
        Set dict("D") = CreateDictionary(UseNative)
        dict("D").Add "key", "D"

        ' Set + Replace
        Set dict("A") = CreateDictionary(UseNative)
        dict("A").Add "key", "A"
        Set dict("B") = CreateDictionary(UseNative)
        dict("B").Add "key", "B"

        ' Should have correct values
        .Expect(dict.Item("A")("key")).ToEqual "A"
        .Expect(dict.Item("B")("key")).ToEqual "B"
        .Expect(dict.Item("C")).ToEqual "ABC"
        .Expect(dict.Item("D")("key")).ToEqual "D"

        ' Should have correct order
        .Expect(dict.Keys()(0)).ToEqual "A"
        .Expect(dict.Keys()(1)).ToEqual "B"
        .Expect(dict.Keys()(2)).ToEqual "C"
        .Expect(dict.Keys()(3)).ToEqual "D"
    End With
    
    With Specs.It("should change key")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        
        dict.Key("B") = "PI"
        .Expect(dict("PI")).ToEqual 3.14
    End With
    
    With Specs.It("should use CompareMode")
        Set dict = CreateDictionary(UseNative)
        dict.CompareMode = 0
        
        dict.Add "A", 123
        dict("a") = 456
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        
        .Expect(dict("A")).ToEqual 123
        .Expect(dict("a")).ToEqual 456
        
        Set dict = CreateDictionary(UseNative)
        dict.CompareMode = 1
        
        dict.Add "A", 123
        dict("a") = 456
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        
        .Expect(dict("A")).ToEqual 456
        .Expect(dict("a")).ToEqual 456
    End With
    
    With Specs.It("should allow Variant for key")
        Set dict = CreateDictionary(UseNative)
        
        Key = "A"
        dict(Key) = 123
        .Expect(dict(Key)).ToEqual 123
        
        Key = "B"
        Set dict(Key) = CreateDictionary(UseNative)
        .Expect(dict(Key).Count).ToEqual 0
    End With
    
    With Specs.It("should handle numeric keys")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add 3, 1
        dict.Add 2, 2
        dict.Add 1, 3
        dict.Add "3", 4
        dict.Add "2", 5
        dict.Add "1", 6

        .Expect(dict(3)).ToEqual 1
        .Expect(dict(2)).ToEqual 2
        .Expect(dict(1)).ToEqual 3
        .Expect(dict("3")).ToEqual 4
        .Expect(dict("2")).ToEqual 5
        .Expect(dict("1")).ToEqual 6
        
        .Expect(dict.Keys()(0)).ToEqual 3
        .Expect(dict.Keys()(1)).ToEqual 2
        .Expect(dict.Keys()(2)).ToEqual 1
        .Expect(TypeName(dict.Keys()(0))).ToEqual "Integer"
        .Expect(TypeName(dict.Keys()(1))).ToEqual "Integer"
        .Expect(TypeName(dict.Keys()(2))).ToEqual "Integer"
    End With
    
    With Specs.It("should handle boolean keys")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add True, 1
        dict.Add False, 2
        
        .Expect(dict(True)).ToEqual 1
        .Expect(dict(False)).ToEqual 2
        
        .Expect(dict.Keys()(0)).ToEqual True
        .Expect(dict.Keys()(1)).ToEqual False
        .Expect(TypeName(dict.Keys()(0))).ToEqual "Boolean"
        .Expect(TypeName(dict.Keys()(1))).ToEqual "Boolean"
    End With
    
    With Specs.It("should handle object keys")
        Set dict = CreateDictionary(UseNative)
        
        Set A = New Collection
        Set B = New Dictionary
        
        A.Add 123
        B.Add "a", 456
        
        dict.Add A, "123"
        dict.Add B, "456"
        
        .Expect(dict(A)).ToEqual "123"
        .Expect(dict(B)).ToEqual "456"
        
        dict.Remove B
        dict.Key(A) = B
        
        .Expect(dict(B)).ToEqual "123"
    End With
    
    ' Methods
    ' ------------------------- '
    With Specs.It("should add an item")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        dict.Add "E", Array(1, 2, 3)
        dict.Add "F", dict
        
        .Expect(dict("A")).ToEqual 123
        .Expect(dict("B")).ToEqual 3.14
        .Expect(dict("C")).ToEqual "ABC"
        .Expect(dict("D")).ToEqual True
        .Expect(dict("E")(1)).ToEqual 2
        .Expect(dict("F")("C")).ToEqual "ABC"
    End With
    
    With Specs.It("should check if an item exists")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "Exists", 123
        .Expect(dict.Exists("Exists")).ToEqual True
        .Expect(dict.Exists("Doesn't Exist")).ToEqual False
    End With
    
    With Specs.It("should get an array of all items")
        Set dict = CreateDictionary(UseNative)
        
        .Expect(dict.Items).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        Items = dict.Items
        .Expect(UBound(Items)).ToEqual 3
        .Expect(Items(0)).ToEqual 123
        .Expect(Items(3)).ToEqual True
        
        dict.Remove "A"
        dict.Remove "B"
        dict.Remove "C"
        dict.Remove "D"
        .Expect(dict.Items).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
    End With
    
    With Specs.It("should get an array of all keys")
        Set dict = CreateDictionary(UseNative)
        
        .Expect(dict.Keys).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        Keys = dict.Keys
        .Expect(UBound(Keys)).ToEqual 3
        .Expect(Keys(0)).ToEqual "A"
        .Expect(Keys(3)).ToEqual "D"
        
        dict.RemoveAll
        .Expect(dict.Keys).RunMatcher "Specs.ToBeAnEmptyArray", "to be an empty array"
    End With
    
    With Specs.It("should remove item")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        .Expect(dict.Count).ToEqual 4
        
        dict.Remove "C"
                
        .Expect(dict.Count).ToEqual 3
    End With
    
    With Specs.It("should remove all items")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        .Expect(dict.Count).ToEqual 4
        
        dict.RemoveAll
        
        .Expect(dict.Count).ToEqual 0
    End With
    
    ' Other
    ' ------------------------- '
    With Specs.It("should For Each over keys")
        Set dict = CreateDictionary(UseNative)
        
        Set Keys = New Collection
        For Each Key In dict.Keys
            Keys.Add Key
        Next Key
        
        .Expect(Keys.Count).ToEqual 0
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        Set Keys = New Collection
        For Each Key In dict.Keys
            Keys.Add Key
        Next Key
        
        .Expect(Keys.Count).ToEqual 4
        .Expect(Keys(1)).ToEqual "A"
        .Expect(Keys(4)).ToEqual "D"
    End With
    
    With Specs.It("should For Each over items")
        Set dict = CreateDictionary(UseNative)
        
        Set Items = New Collection
        For Each Item In dict.Items
            Items.Add Item
        Next Item
        
        .Expect(Items.Count).ToEqual 0
        
        dict.Add "A", 123
        dict.Add "B", 3.14
        dict.Add "C", "ABC"
        dict.Add "D", True
        
        Set Items = New Collection
        For Each Item In dict.Items
            Items.Add Item
        Next Item
        
        .Expect(Items.Count).ToEqual 4
        .Expect(Items(1)).ToEqual 123
        .Expect(Items(4)).ToEqual True
    End With
    
    With Specs.It("should have UBound of -1 for empty Keys and Items")
        Set dict = CreateDictionary(UseNative)
        
        .Expect(UBound(dict.Keys)).ToEqual -1
        .Expect(UBound(dict.Items)).ToEqual -1
        .Expect(Err.Number).ToEqual 0
    End With
    
    ' Errors
    ' ------------------------- '
    Err.Clear
    On Error Resume Next
    
    With Specs.It("should throw 5 when changing CompareMode with items in Dictionary")
        Set dict = CreateDictionary(UseNative)
        dict.Add "A", 123
        
        dict.CompareMode = vbTextCompare
        
        .Expect(Err.Number).ToEqual 5
        Err.Clear
    End With
    
    With Specs.It("should throw 457 on Add if key exists")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add "A", 123
        dict.Add "A", 456
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
        
        dict.RemoveAll
        dict.Add "A", 123
        dict.Add "a", 456
        
        .Expect(Err.Number).ToEqual 0
        Err.Clear
        
        dict.RemoveAll
        dict.CompareMode = vbTextCompare
        dict.Add "A", 123
        dict.Add "a", 456
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
    End With
    
    With Specs.It("should throw 32811 on Remove if key doesn't exist")
        Set dict = CreateDictionary(UseNative)
        
        dict.Remove "A"
        
        .Expect(Err.Number).ToEqual 32811
        Err.Clear
    End With
    
    With Specs.It("should throw 457 for Boolean key quirks")
        Set dict = CreateDictionary(UseNative)
        
        dict.Add True, "abc"
        dict.Add -1, "def"
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
        
        dict.Add False, "abc"
        dict.Add 0, "def"
        
        .Expect(Err.Number).ToEqual 457
        Err.Clear
    End With
    
    On Error GoTo 0
    InlineRunner.RunSuite Specs
End Function

Public Sub ExecuteSpeedTest(Optional CompareToNative As Boolean = False)
    Dim Counts As Variant
    Counts = Array(5000, 5000, 5000, 5000, 7500, 7500, 7500, 7500)
    
    Dim Baseline As Collection
    If CompareToNative Then
        Set Baseline = RunSpeedTest(Counts, True)
    End If
    
    Dim results As Collection
    Set results = RunSpeedTest(Counts, False)
    
    Debug.Print vbNewLine & "SpeedTest Results:" & vbNewLine
    PrintResults "Add", Baseline, results, 0
    PrintResults "Iterate", Baseline, results, 1
End Sub

Public Sub PrintResults(Test As String, Baseline As Collection, results As Collection, Index As Integer)
    Dim BaselineAvg As Single
    Dim ResultsAvg As Single
    Dim i As Integer
    
    If Not Baseline Is Nothing Then
        For i = 1 To Baseline.Count
            BaselineAvg = BaselineAvg + Baseline(i)(Index)
        Next i
        BaselineAvg = BaselineAvg / Baseline.Count
    End If
    
    For i = 1 To results.Count
        ResultsAvg = ResultsAvg + results(i)(Index)
    Next i
    ResultsAvg = ResultsAvg / results.Count
    
    Dim Result As String
    Result = Test & ": " & Format(Round(ResultsAvg, 0), "#,##0") & " ops./s"
    
    If Not Baseline Is Nothing Then
        Result = Result & " vs. " & Format(Round(BaselineAvg, 0), "#,##0") & " ops./s "
    
        If ResultsAvg < BaselineAvg Then
            Result = Result & Format(Round(BaselineAvg / ResultsAvg, 0), "#,##0") & "x slower"
        ElseIf BaselineAvg > ResultsAvg Then
            Result = Result & Format(Round(ResultsAvg / BaselineAvg, 0), "#,##0") & "x faster"
        End If
    End If
    Result = Result
    
    If results.Count > 1 Then
        Result = Result & vbNewLine
        For i = 1 To results.Count
            Result = Result & "  " & Format(Round(results(i)(Index), 0), "#,##0")
            
            If Not Baseline Is Nothing Then
                Result = Result & " vs. " & Format(Round(Baseline(i)(Index), 0), "#,##0")
            End If
            
            Result = Result & vbNewLine
        Next i
    End If
    
    Debug.Print Result
End Sub

Public Function RunSpeedTest(Counts As Variant, Optional UseNative As Boolean = False) As Collection
    Dim results As New Collection
    Dim CountIndex As Integer
    Dim dict As Object
    Dim i As Long
    Dim AddResult As Double
    Dim Key As Variant
    Dim value As Variant
    Dim IterateResult As Double
    Dim Timer As New PreciseTimer
    
    For CountIndex = LBound(Counts) To UBound(Counts)
        Timer.StartTimer
    
        Set dict = CreateDictionary(UseNative)
        For i = 1 To Counts(CountIndex)
            dict.Add CStr(i), i
        Next i
        
        ' Convert to seconds
        AddResult = Timer.TimeElapsed / 1000#
        
        ' Convert to ops./s
        If AddResult > 0 Then
            AddResult = Counts(CountIndex) / AddResult
        Else
            ' Due to single precision, timer resolution is 0.01 ms, set to 0.005 ms
            AddResult = Counts(CountIndex) / 0.005
        End If
        
        Timer.StartTimer
        
        For Each Key In dict.Keys
            value = dict.Item(Key)
        Next Key
        
        ' Convert to seconds
        IterateResult = Timer.TimeElapsed / 1000#
        
        ' Convert to ops./s
        If IterateResult > 0 Then
            IterateResult = Counts(CountIndex) / IterateResult
        Else
            ' Due to single precision, timer resolution is 0.01 ms, set to 0.005 ms
            IterateResult = Counts(CountIndex) / 0.005
        End If
        
        results.Add Array(AddResult, IterateResult)
    Next CountIndex
    
    Set RunSpeedTest = results
End Function

Public Function CreateDictionary(Optional UseNative As Boolean = False) As Object
    If UseNative Then
        Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Else
        Set CreateDictionary = New Dictionary
    End If
End Function

Public Function ToBeAnEmptyArray(Actual As Variant) As Variant
    Dim UpperBound As Long

    Err.Clear
    On Error Resume Next
    
    ' First, make sure it's an array
    If IsArray(Actual) = False Then
        ' we weren't passed an array, return True
        ToBeAnEmptyArray = True
    Else
        ' Attempt to get the UBound of the array. If the array is
        ' unallocated, an error will occur.
        UpperBound = UBound(Actual, 1)
        If (Err.Number <> 0) Then
            ToBeAnEmptyArray = True
        Else
            ' Check for case of -1 UpperBound (Scripting.Dictionary.Keys/Items)
            Err.Clear
            If LBound(Actual) > UpperBound Then
                ToBeAnEmptyArray = True
            Else
                ToBeAnEmptyArray = False
            End If
        End If
    End If
End Function

#End If


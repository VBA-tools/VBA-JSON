# VBA-JSON

JSON conversion and parsing for VBA (Windows and Mac Excel, Access, and other Office applications).
It grew out of the excellent project [vba-json](https://code.google.com/p/vba-json/),
with additions and improvements made to resolve bugs and improve performance (as part of [VBA-Web](https://github.com/VBA-tools/VBA-Web)).

Tested in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+.

- For Windows-only support, include a reference to "Microsoft Scripting Runtime"
- For Mac and Windows support, include [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

<a href="https://www.patreon.com/timhall">
  <img src="https://timhall.github.io/assets/donate-patreon@2x.png" width="217" alt="Donate">
</a>

# Examples

```vb
Dim Json As Object
Set Json = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")

' Json("a") -> 123
' Json("b")(2) -> 2
' Json("c")("d") -> 456
Json("c")("e") = 789

Debug.Print JsonConverter.ConvertToJson(Json)
' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"

Debug.Print JsonConverter.ConvertToJson(Json, Whitespace:=2)
' -> "{
'       "a": 123,
'       "b": [
'         1,
'         2,
'         3,
'         4
'       ],
'       "c": {
'         "d": 456,
'         "e": 789  
'       }
'     }"
```

```vb
' Advanced example: Read .json file and load into sheet (Windows-only)
' (add reference to Microsoft Scripting Runtime)
' {"values":[{"a":1,"b":2,"c": 3},...]}

Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary

' Read .json file
Set JsonTS = FSO.OpenTextFile("example.json", ForReading)
JsonText = JsonTS.ReadAll
JsonTS.Close

' Parse json to Dictionary
' "values" is parsed as Collection
' each item in "values" is parsed as Dictionary
Set Parsed = JsonConverter.ParseJson(JsonText)

' Prepare and write values to sheet
Dim Values As Variant
ReDim Values(Parsed("values").Count, 3)

Dim Value As Dictionary
Dim i As Long

i = 0
For Each Value In Parsed("values")
  Values(i, 0) = Value("a")
  Values(i, 1) = Value("b")
  Values(i, 2) = Value("c")
  i = i + 1
Next Value

Sheets("example").Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
```

## Options

VBA-JSON includes a few options for customizing parsing/conversion if needed:

- __UseDoubleForLargeNumbers__ (Default = `False`) VBA only stores 15 significant digits, so any numbers larger than that are truncated.
  This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits.
  By default, VBA-JSON will use `String` for numbers longer than 15 characters that contain only digits, use this option to use `Double` instead.
- __AllowUnquotedKeys__ (Default = `False`) The JSON standard requires object keys to be quoted (`"` or `'`), use this option to allow unquoted keys.
- __EscapeSolidus__ (Default = `False`) The solidus (`/`) is not required to be escaped, use this option to escape them as `\/` in `ConvertToJson`.

```VB.net
JsonConverter.JsonOptions.EscapeSolidus = True
```

## Installation

1. Download the [latest release](https://github.com/VBA-tools/VBA-JSON/releases)
2. Import `JsonConverter.bas` into your project (Open VBA Editor, `Alt + F11`; File > Import File)
3. Add `Dictionary` reference/class
   - For Windows-only, include a reference to "Microsoft Scripting Runtime"
   - For Windows and Mac, include [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

## Resources

- [Tutorial Video (Red Stapler)](https://youtu.be/CFFLRmHsEAs)

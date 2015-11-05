# VBA-JSON

JSON conversion and parsing for VBA (Windows and Mac Excel, Access, and other Office applications). 
It grew out of the excellent project [vba-json](https://code.google.com/p/vba-json/), 
with additions and improvements made to resolve bugs and improve performance (as part of [VBA-Web](https://github.com/VBA-tools/VBA-Web)).

Tested in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+. 

- For Windows-only support, include a reference to "Microsoft Scripting Runtime"
- For Mac and Windows support, include [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary).

# Example

```VB.net
Dim Json As Object
Set Json = JsonConverter.ParseJSON("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")

' Json("a") -> 123
' Json("b")(2) -> 2
' Json("c")("d") -> 456
Json("c")("e") = 789

Debug.Print JsonConverter.ConvertToJson(Json) 
' -> "{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456,""e"":789}}"
```

## Options

VBA-JSON includes a few options for customizing parsing/conversion if needed:

- __UseDoubleForLargeNumbers__ (Default = `False`) VBA only stores 15 significant digits, so any numbers larger than that are truncated.
  This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits.
  By default, VBA-JSON will use `String` for numbers longer than 15 characters that contain only digits, use this option to use `Double` instead.
- __AllowUnquotedKeys__ (Default = `False`) The JSON standard requires object keys to be quoted (`"` or `'`), use this option to allow unquoted keys.
- __EscapeSolidus__ (Default = `False`) The solidus (/) is not required to be escaped, use this option to escape them as `\/` in `ConvertToJson`.

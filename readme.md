# VBA-JSON

JSON conversion and parsing for VBA (Windows and Mac Excel, Access, and other Office applications). 
It grew out of the excellent project [vba-json](https://code.google.com/p/vba-json/), 
with additions and improvements made to resolve bugs and improve performance (as part of [Excel-REST](https://github.com/timhall/Excel-REST)).

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

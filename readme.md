# VBA-JSONConverter

JSON conversion and parsing for VBA (Excel, Access, and other Office applications). It grew out of the excellent project [vba-json](https://code.google.com/p/vba-json/), with additions and improvements made to resolve bugs and improve performance (as part of [Excel-REST](https://github.com/timhall/Excel-REST)).

Tested in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+. 

- For Windows-only support, include a reference to "Microsoft Scripting Runtime"
- For Mac support or to skip adding a reference, include [VBA-Dictionary](https://github.com/timhall/VBA-Dictionary).

# Example

```VB
Dim JSON As Object
Set JSON = JSONConverter.JSONParse("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")

' JSON("a") -> 123
' JSON("b")(2) -> 3
' JSON("c")("d") -> 456

Debug.Print JSONConverter.JSONToString(JSON) 
' -> "{""a"":123,""b"":[[1,2],[3,4]],""c"":{""d"":456}}"
```

# VBA-JSONConverter

JSON conversion and parsing for VBA (Excel, Access, and other Office applications). It grew out of the excellent project [vba-json](https://code.google.com/p/vba-json/), with additions and improvements made to resolve bugs and improve performance (as part of [Excel-REST](https://github.com/timhall/Excel-REST)).

Tested in Windows Excel 2013, but should apply to 2007+. For Mac support, include VBA-Dictionary (coming soon!)

# Example

```
Dim JSON As Object
Set JSON = JSONConverter.Parse("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")

' JSON("a") -> 123
' JSON("b")(2) -> 3
' JSON("c")("d") -> 456

Debug.Print JSONConverter.ToString(JSON) 
' or JSONConverter.Stringify(JSON)
' -> "{""a"":123,""b"":[[1,2],[3,4]],""c"":{""d"":456}}"
```
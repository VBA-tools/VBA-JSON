Attribute VB_Name = "JSONConverter"
''
' Convert JSON string to object (Dictionary/Collection)
'
' @param {String} JSON
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function Parse(JSON As String) As Object
    
End Function

''
' Convert JSON object (Dictionary/Collection/Array) to string
'
' @param {Object} JSON (Dictionary, Collection, or Array)
' @return {String}
' -------------------------------------- '
Public Function ToString(JSON As Object) As String
    
End Function
Public Function Stringify(JSON As Object) As String
    Stringify = ToString(JSON)
End Function

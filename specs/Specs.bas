Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-JSONConverter"
    
    InlineRunner.RunSuite Specs
End Function

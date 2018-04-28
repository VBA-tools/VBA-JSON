Attribute VB_Name = "modTest2"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestRoundTrip
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : For a super simple example check that Dictionary > JSON String > Dictionary gets back to where we started...
' -----------------------------------------------------------------------------------------------------------------------
Sub TestRoundTrip()

    Dim DCTIn As New Dictionary
    Dim DCTOut As Dictionary
    Dim JsonString As String
    Dim JsonString2 As String

    DCTIn.Add "Number", 100
    DCTIn.Add "String", "Hello"
    DCTIn.Add "Array", Array(1, 2, 3, 4, 5)
    JsonString = ConvertToJson(DCTIn)
    
    Set DCTOut = ParseJson(JsonString)
    JsonString2 = ConvertToJson(DCTOut)
    
    Debug.Print JsonString = JsonString2

End Sub

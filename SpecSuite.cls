VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' SpecSuite v1.4.0
' (c) Tim Hall - https://github.com/timhall/Excel-TDD
'
' A collection of specs with the workbook that they act on
'
' @dependencies
' @author tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private pSpecsCol As Collection

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Description As String

Public BeforeEachCallback As String
Public BeforeEachCallbackArgs As Variant

Public Property Get SpecsCol() As Collection
    If pSpecsCol Is Nothing Then: Set pSpecsCol = New Collection
    Set SpecsCol = pSpecsCol
End Property
Public Property Let SpecsCol(value As Collection)
    Set pSpecsCol = value
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Create a new spec definition with description
'
' @param {String} Description
' @param {String} [SpecId] Useful for identifying specific specs
' @returns {SpecDefinition} Initialized Spec Definition
' --------------------------------------------- '

Public Function It(Description As String, Optional SpecId As String = "") As SpecDefinition
    Dim Spec As New SpecDefinition
    
    ' Call BeforeEach if defined
    ExecuteBeforeEach
    
    ' Initialize spec
    Spec.Description = Description
    Spec.Id = SpecId
    Me.SpecsCol.Add Spec
    Set It = Spec
End Function

''
' Setup a callback to run before each spec
'
' @param {String} Callback
' @param {...} CallbackArgs any additional arguments to pass as array to callback each time
'
' Example:
' BeforeEach "Cleanup", 100, 200
'
' ' Cleanup is called before each spec with Args = [100, 200]
' Sub Cleanup(Args As Variant)
'   ThisWorkbook.Sheets("Main").Cells(1, 1) = Args(0) ' (100)
'   ThisWorkbook.Sheets("Main").Cells(2, 1) = Args(1) ' (200)
' End Sub
' --------------------------------------------- '

Public Sub BeforeEach(Callback As String, ParamArray CallbackArgs() As Variant)
    Me.BeforeEachCallback = Callback
    Me.BeforeEachCallbackArgs = CallbackArgs
End Sub


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Sub ExecuteBeforeEach()
    If Me.BeforeEachCallback <> "" Then
        Dim HasArguments As Boolean
        If VarType(Me.BeforeEachCallbackArgs) = vbObject Then
            If Not Me.BeforeEachCallbackArgs Is Nothing Then
                HasArguments = True
            End If
        ElseIf IsArray(Me.BeforeEachCallbackArgs) Then
            If UBound(Me.BeforeEachCallbackArgs) >= 0 Then
                HasArguments = True
            End If
        End If
    
        If HasArguments Then
            Application.Run Me.BeforeEachCallback, Me.BeforeEachCallbackArgs
        Else
            Application.Run Me.BeforeEachCallback
        End If
    End If
End Sub

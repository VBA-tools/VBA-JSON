VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringAppend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : clsStringAppend
' Author    : Philip Swannell
' Date      : 26-Jan-2018
' Purpose   : Class for constructing strings in a loop, avoiding "Shlemiel the painter" performance
'---------------------------------------------------------------------------------------
Option Base 1
Dim m_TheString As String
Dim m_NumCharsWritten As Long
Dim m_NumCharsStored As Long

Public Function Report()
1         On Error GoTo ErrHandler
2         Report = VBA.Left$(m_TheString, m_NumCharsWritten)
3         Exit Function
ErrHandler:
4         Err.Raise vbObjectError + 1, , "#clsStringAppend.Report (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function Max(x As Long, y As Long)
1         If x > y Then
2             Max = x
3         Else
4             Max = y
5         End If
End Function

Public Sub Append(TheString As String)
          Dim L As Long
          Dim NumCharsToAdd As Long
1         On Error GoTo ErrHandler
2         L = VBA.Len(TheString)
          
3         If L + m_NumCharsWritten > m_NumCharsStored Then
4             NumCharsToAdd = Max(L, m_NumCharsStored)
5             m_TheString = m_TheString + VBA.Space$(NumCharsToAdd)
6             m_NumCharsStored = m_NumCharsStored + NumCharsToAdd
7         End If

8         Mid$(m_TheString, m_NumCharsWritten + 1, L) = TheString
9         m_NumCharsWritten = m_NumCharsWritten + L

10        Exit Sub
ErrHandler:
11        Err.Raise vbObjectError + 1, , "#clsStringAppend.Append (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Attribute VB_Name = "modTest1"
Option Explicit

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

#If Mac Then
#ElseIf VBA7 Then

Private Declare PtrSafe Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)

#Else

Private Declare Sub json_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (json_MemoryDestination As Any, json_MemorySource As Any, ByVal json_ByteLength As Long)

#End If

'---------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Returns time in seconds since system start up. High resolution.
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
'---------------------------------------------------------------------------------------
Function sElapsedTime() As Double
          Dim a As Currency, b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         QueryPerformanceFrequency b
4         sElapsedTime = a / b
5         Exit Function
ErrHandler:
6         Err.Raise vbObjectError + 1, , "#sElapsedTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CompareTwoMethods
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Test harness to compare execution speed of existing json_BufferAppend versus clsStringAppend
'              For N from 1000 to 1000000 I get clsAppend approx 5 times faster than json_BufferAppend
'              In addition, clsAppend does not use Windows API calls and thus should work on Mac. I presume (but
'              haven't tested) that the code as is exhibits ""Shlemiel the painter" performance on Mac since
'              method json_BufferAppend just does naive string append on Mac.
' -----------------------------------------------------------------------------------------------------------------------
Sub CompareTwoMethods()
          Dim Result1 As String, Result2 As String
          Dim AppendThis As String
          Dim i As Long, N As Long
          Dim t1 As Double, t2 As Double, t3 As Double
          Dim json_buffer As String
          Dim json_BufferPosition As Long
          Dim json_BufferLength As Long
          Dim cSA As New clsStringAppend
          
1         On Error GoTo ErrHandler
2         AppendThis = "xyz"
3         N = 100000

4         t1 = sElapsedTime()

5         For i = 1 To N
6             json_BufferAppend json_buffer, AppendThis, json_BufferPosition, json_BufferLength
7         Next i
8         Result1 = json_BufferToString(json_buffer, json_BufferPosition, json_BufferLength)

9         t2 = sElapsedTime()

10        For i = 1 To N
11            cSA.Append AppendThis
12        Next i
13        Result2 = cSA.Report

14        t3 = sElapsedTime()

15        Debug.Print String(50, "-")
16        Debug.Print "N = " & Format(N, "###,###") & " Len(AppendThis) = " & Len(AppendThis)
17        Debug.Print "Results Agree?", Result1 = Result2
18        Debug.Print "Time json_BufferToString", Format((t2 - t1) * 1000, "0.000") & " milliseconds"
19        Debug.Print "Time clsStringAppend", Format((t3 - t2) * 1000, "0.000") & " milliseconds"
20        Debug.Print "Ratio:               ", (t2 - t1) / (t3 - t2)


21        Exit Sub
ErrHandler:
22        MsgBox "#CompareTwoMethods (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub

Private Sub json_BufferAppend(ByRef json_buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
#If Mac Then
    json_buffer = json_buffer & json_Append
#Else
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.LenB(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim json_TemporaryLength As Long

        json_TemporaryLength = json_BufferLength
        Do While json_TemporaryLength < json_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If json_TemporaryLength = 0 Then
                json_TemporaryLength = json_TemporaryLength + 510
            Else
                json_TemporaryLength = json_TemporaryLength + 16384
            End If
        Loop

        json_buffer = json_buffer & VBA.Space$((json_TemporaryLength - json_BufferLength) \ 2)
        json_BufferLength = json_TemporaryLength
    End If

    ' Copy memory from append to buffer at buffer position
    json_CopyMemory ByVal json_UnsignedAdd(StrPtr(json_buffer), _
                    json_BufferPosition), _
                    ByVal StrPtr(json_Append), _
                    json_AppendLength

    json_BufferPosition = json_BufferPosition + json_AppendLength
#End If
End Sub

Private Function json_BufferToString(ByRef json_buffer As String, ByVal json_BufferPosition As Long, ByVal json_BufferLength As Long) As String
#If Mac Then
    json_BufferToString = json_buffer
#Else
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_buffer, json_BufferPosition \ 2)
    End If
#End If
End Function

#If VBA7 Then
Private Function json_UnsignedAdd(json_Start As LongPtr, json_Increment As Long) As LongPtr
#Else
Private Function json_UnsignedAdd(json_Start As Long, json_Increment As Long) As Long
#End If

    If json_Start And &H80000000 Then
        json_UnsignedAdd = json_Start + json_Increment
    ElseIf (json_Start Or &H80000000) < -json_Increment Then
        json_UnsignedAdd = json_Start + json_Increment
    Else
        json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
    End If
End Function





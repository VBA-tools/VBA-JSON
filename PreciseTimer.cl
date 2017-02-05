VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreciseTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' PreciseTimer
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Very accurate timer (Windows) with fallback to VBA Timer for Mac
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

#If Mac Then
Private pStartTime As Single
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long


' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const TWO_32 = 4294967296# ' = 256# * 256# * 256# * 256#
Private pCounterStart As LARGE_INTEGER
Private pCounterEnd As LARGE_INTEGER
Private pFrequency As Double

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
#End If

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Get ms that have elapsed
' --------------------------------------------- '
Property Get TimeElapsed() As Double
#If Mac Then
    ' Convert Single seconds to Double ms
    TimeElapsed = 1000# * CDbl(VBA.Timer - pStartTime)
#Else
    Dim crStart As Double
    Dim crStop As Double
    
    QueryPerformanceCounter pCounterEnd
    
    crStart = LargeIntToDouble(pCounterStart)
    crStop = LargeIntToDouble(pCounterEnd)
    TimeElapsed = 1000# * (crStop - crStart) / pFrequency
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Start timer
' --------------------------------------------- '
Public Sub StartTimer()
#If Mac Then
    pStartTime = VBA.Timer
#Else
    QueryPerformanceCounter pCounterStart
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then
#Else
Private Function LargeIntToDouble(LargeInt As LARGE_INTEGER) As Double
    Dim Low As Double
    Low = LargeInt.lowpart
    If Low < 0 Then
        Low = Low + TWO_32
    End If
    LargeIntToDouble = LargeInt.highpart * TWO_32 + Low
End Function

Private Sub Class_Initialize()
    Dim PerfFrequency As LARGE_INTEGER
    QueryPerformanceFrequency PerfFrequency
    pFrequency = LargeIntToDouble(PerfFrequency)
End Sub
#End If


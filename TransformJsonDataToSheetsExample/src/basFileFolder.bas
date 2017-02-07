Attribute VB_Name = "basFileFolder"
Option Explicit
'Authored 2015-2017 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
     'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...

Public Const ForReading = 1
Public Const ForWriting = 2
Public Const ForAppending = 8

Private Declare Function URLDownloadToFileA Lib "urlmon" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, _
    ByVal lpfnCB As Long) _
As Long

Public Function RemoveForbiddenFilenameCharacters(strFilename As String) As String
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247(v=vs.85).aspx
'< (less than)
'> (greater than)
': (colon)
'" (double quote)
'/ (forward slash)
'\ (backslash)
'| (vertical bar or pipe)
'? (question mark)
'* (asterisk)
Dim strForbidden As Variant
    For Each strForbidden In Array("/", "\", "|", ":", "*", "?", "<", ">", """")
        strFilename = Replace(strFilename, strForbidden, "_")
    Next
    RemoveForbiddenFilenameCharacters = strFilename
End Function

Public Function DownloadUrlFileToTemp( _
    ByVal strUrl As String, _
    Optional ByVal strDestinationExtension As String = "txt", _
    Optional strJsonArchiveDirectory As String) _
As String
    Dim lngRetVal As Long
    Dim strTempFilePath As String
    Dim strTargetDirectory As String
    strTempFilePath = Left(RemoveForbiddenFilenameCharacters(Right(strUrl, Len(strUrl) - InStrRev(strUrl, "/"))), 30)
    If Len(strJsonArchiveDirectory) > 0 Then 'should be validating that the dir exists and we can write to it.
        strTargetDirectory = GetRelativePathViaParent(strJsonArchiveDirectory)
    Else
        strTargetDirectory = Environ$("TEMP")
    End If
    strTempFilePath = strTargetDirectory & "\" & strTempFilePath & Format(Now(), "yymmddhhss") & Right(Timer, 2) & "." & strDestinationExtension
    lngRetVal = URLDownloadToFileA(0, strUrl, strTempFilePath, 0, 0)
    If lngRetVal Then
        Err.Raise Err.LastDllError, , "Download failed."
    End If
    DownloadUrlFileToTemp = strTempFilePath
    Debug.Print strTempFilePath
End Function

Public Function DeleteFile(strPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object ' As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If fso.FileExists(strPath) Then
        fso.DeleteFile strPath
    End If
    DeleteFile = Err.Number = 0
    Set fso = Nothing
End Function

Public Function BuildDir(strPath) As Boolean
    On Error Resume Next
    Dim fso As Object ' As Scripting.FileSystemObject
    Dim arryPaths As Variant
    Dim strBuiltPath As String, intDir As Integer, fRestore As Boolean: fRestore = False
    If Left(strPath, 2) = "\\" Then
        strPath = Right(strPath, Len(strPath) - 2)
        fRestore = True
    End If
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    arryPaths = Split(strPath, "\")
    'Restore Server file path
    If fRestore Then
        arryPaths(0) = "\\" & arryPaths(0)
    End If
    For intDir = 0 To UBound(arryPaths)
        strBuiltPath = strBuiltPath & arryPaths(intDir)
        If Not fso.FolderExists(strBuiltPath) Then
            fso.CreateFolder strBuiltPath
        End If
        strBuiltPath = strBuiltPath & "\"
    Next
    BuildDir = (Err.Number = 0) 'True if no errors
    Set fso = Nothing
End Function

Public Function GetRelativePathViaParentAlternateRoot(ByVal strRootPath As String, ByVal strDestination As String, Optional ByRef intParentCount As Integer)
    If Left(strDestination, 3) = "..\" Then
        intParentCount = intParentCount + 1
        strRootPath = Left(strRootPath, InStrRev(strRootPath, "\") - 1)
        strDestination = Right(strDestination, Len(strDestination) - 3)
        GetRelativePathViaParentAlternateRoot = GetRelativePathViaParentAlternateRoot(strRootPath, strDestination, intParentCount)
    ElseIf Left(strDestination, 1) = "\" And Not (Left(strDestination, 2) = "\\") Then
        strDestination = Right(strDestination, Len(strDestination) - 1)
    ElseIf Right(strDestination, 1) = "\" Then
        strDestination = Left(strDestination, Len(strDestination) - 1)
    End If
    If intParentCount <> -1 Then
        GetRelativePathViaParentAlternateRoot = StripTrailingBackSlash(strRootPath) & "\" & strDestination
    End If
    intParentCount = -1
End Function

Public Function GetRelativePathViaParent(Optional ByVal strPath As String)
'Usage for up 2 dirs is GetRelativePathViaParent("..\..\Destination")
    Dim strVal As String
    If Left(strPath, 2) = "\\" Or Mid(strPath, 2, 1) = ":" Then
        strVal = strPath
    Else
        Dim strCurrentPath As String
        Dim oThisApplication As Object:    Set oThisApplication = Application
        Select Case True
            Case InStrRev(oThisApplication.Name, "Excel") > 0
                strCurrentPath = oThisApplication.ThisWorkbook.Path
            Case InStrRev(oThisApplication.Name, "Access") > 0
                strCurrentPath = oThisApplication.CurrentProject.Path
        End Select
        Dim fIsServerPath As Boolean: fIsServerPath = False
         If Left(strCurrentPath, 2) = "\\" Then
             strCurrentPath = Right(strCurrentPath, Len(strCurrentPath) - 2)
             fIsServerPath = True
        End If
        Dim aryCurrentFolder As Variant
        aryCurrentFolder = Split(strCurrentPath, "\")
        Dim aryParentPath As Variant
        aryParentPath = Split(strPath, "..\")
        If fIsServerPath Then
            aryCurrentFolder(0) = "\\" & aryCurrentFolder(0)
        End If
        Dim intDir As Integer
        For intDir = 0 To UBound(aryCurrentFolder) - UBound(aryParentPath)
            strVal = strVal & aryCurrentFolder(intDir) & "\"
        Next
        strVal = StripTrailingBackSlash(strVal)
        If IsArrayAllocated(aryParentPath) Then
            strVal = strVal & "\" & aryParentPath(UBound(aryParentPath))
        End If
    End If
    If BuildDir(strVal) Then
        GetRelativePathViaParent = strVal
    End If
End Function

Public Sub SaveStringToFile(ByRef strFilePath As String, ByRef strString As String)
On Error GoTo HandleError

Dim intFileNumber As Long
Dim abyteByteArray() As Byte

    ' Delete existing file if needed
    If LenB(Dir(strFilePath)) <> 0 Then _
        Kill strFilePath

    ' Get free file number
    intFileNumber = FreeFile
    ' Open file for binary write
    Open strFilePath For Binary Access Write As intFileNumber
    ' Convert string to byte array
    ' Note: Must save string as byte array or Put function
    ' will convert string from unicode to ANSI.
    ' Empty string will NOT cause error.
    abyteByteArray() = strString
    ' Save data to file
    ' Note: Unallocated array will NOT cause error.
    Put intFileNumber, 1, abyteByteArray()
    ' Close file
    Close intFileNumber

ExitHere:
    Exit Sub

HandleError:
    ' Close file if needed
    ' Note: Below line of code will not raise an error even if no file is open
    Close intFileNumber
    Select Case Err.Number
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select

End Sub

Private Function IsArrayAllocated(ByRef avarArray As Variant) As Boolean
    On Error Resume Next
    ' Normally we only need to check LBound to determine if an array has been allocated.
    ' Some function such as Split will set LBound and UBound even if array is not allocated.
    ' See http://www.cpearson.com/excel/isarrayallocated.aspx for more details.
    IsArrayAllocated = IsArray(avarArray) And _
        Not IsError(LBound(avarArray, 1)) And _
        LBound(avarArray, 1) <= UBound(avarArray, 1)
End Function


Public Function StripTrailingBackSlash(ByRef strPath As String)
        If Right(strPath, 1) = "\" Then
            StripTrailingBackSlash = Left(strPath, Len(strPath) - 1)
        Else
            StripTrailingBackSlash = strPath
        End If
End Function

Public Sub OpenFileWithExplorer(ByRef strFilePath As String, Optional ByRef fReadOnly As Boolean = True)

    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    Set wshShell = Nothing

End Sub

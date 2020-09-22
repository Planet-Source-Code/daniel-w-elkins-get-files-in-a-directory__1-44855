Attribute VB_Name = "ModFunc"
Option Explicit

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function WinDir() As String
Dim sWin As String * 255
GetWindowsDirectory sWin, 255
WinDir = Trim$(sWin)
sWin = Empty
End Function

Public Sub AddDirFiles(sDir As String, objLV As ListView)
Dim A As Long
Dim sTmp As String, sMap As String, sDir2 As String
Dim sFileData As String
If Right(sDir, 1) = "\" Then
sDir2 = sDir
Else
sDir2 = sDir & "\"
End If
With objLV
.ListItems.Clear
sTmp = Dir(sDir2)
If Len(sTmp) > 0 Then
Do
sMap = sTmp
.ListItems.Add , , sMap, , "FileName"
.ListItems(.ListItems.Count).ListSubItems.Add , , FileLen(sDir2 & sTmp) & " (Bytes)", "FileSize"
.ListItems(.ListItems.Count).ListSubItems.Add , , sDir2 & sTmp, "FilePath"
DoEvents
sTmp = Dir$
Loop Until Len(sTmp) = 0
End If
End With
End Sub

Function GetFileName(sPath As String) As String
On Error Resume Next
Dim sBuff() As String: sBuff() = Split(sPath, "\")
GetFileName = sBuff(UBound(sBuff))
End Function

Function DirFileNum(sDir As String) As Long
Dim sTmp As String, sMapName As String
If Right(sDir, 1) <> "\" Then
sTmp = Dir(sDir & "\")
Else
sTmp = Dir(sDir)
End If
If Len(sTmp) > 0 Then
Do
sMapName = sTmp
DirFileNum = DirFileNum + 1
sTmp = Dir
DoEvents
Loop Until Len(sTmp) = 0
End If
End Function


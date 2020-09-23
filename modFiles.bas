Attribute VB_Name = "modFiles"
Option Explicit
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Public Function FilePath(ByVal Pth As String) As String
 FilePath = Left$(Pth, InStrRev(Pth, "\") - 1)
End Function
Public Function FileTitle(ByVal Pth As String) As String
 FileTitle = Mid$(Pth, InStrRev(Pth, "\") + 1)
End Function
Public Function TrimNull(StrZ As String) As String
'much faster than instr, left$
 TrimNull = Left$(StrZ, lstrlenW(StrPtr(StrZ)))
End Function
Public Function QualifyPath(ByVal sPath As String) As String
 If Right$(sPath, 1) <> "\" Then
  QualifyPath = sPath & "\"
 Else: QualifyPath = sPath
 End If
End Function
Public Function UnQualifyPath(ByVal sPath As String) As String
 If Right$(sPath, 1) = "\" Then
  UnQualifyPath = Left$(sPath, Len(sPath) - 1)
 Else: UnQualifyPath = sPath
 End If
End Function

Public Function ReadFileBinary(ByVal sFile As String) As String
 Dim hFile As Long
 hFile = FreeFile
 Open sFile For Binary As #hFile
 ReadFileBinary = String$(LOF(hFile), Chr$(0))
 Get #hFile, , ReadFileBinary
 Close #hFile
End Function
Public Function FileExists(ByVal sFile As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sFile)
 FileExists = (Err.Number = 0) And ((eAttr And vbDirectory) = 0)
 On Error GoTo 0
End Function
'===============================
Public Function FolderExists(ByVal sPath As String) As Boolean
 Dim eAttr As Long
 On Error Resume Next
 eAttr = GetAttr(sPath)
 FolderExists = (Err.Number = 0) And ((eAttr And vbDirectory) = vbDirectory)
 On Error GoTo 0
End Function

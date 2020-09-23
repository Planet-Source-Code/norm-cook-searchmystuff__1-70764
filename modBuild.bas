Attribute VB_Name = "modBuild"
Option Explicit
Private Const ARR_MAX              As Long = &H3FFF&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH             As Long = 260
Private Const VBDOT                As Long = 46
Private Const SDS                  As String = "*.*"

Public Type WIN32_FIND_DATA
 dwFileAttributes As Long
 ftCreationTime As Currency
 ftLastAccessTime As Currency
 ftLastWriteTime As Currency
 nFileSizeHigh As Long
 nFileSizeLow As Long
 dwReserved0 As Long
 dwReserved1 As Long
 cFileName As String * MAX_PATH
 cAlternate As String * 14
End Type

Private Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecA" (ByVal pszFile As String, ByVal pszSpec As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'use with Spec = "*.frm; *.bas" etc
Public Sub EnumFilesStringArrayWildCard(ByVal StartPath As String, _
    ByRef Arr() As String, ByRef Count As Long, _
    ByVal Spec As String, _
    Optional ByVal Recursive As Boolean = True)
 Count = 0 'just in case
 ReDim Arr(1 To ARR_MAX) 'vice redim preserve on each file
 EFIStrArrWC QualifyPath(StartPath), Arr, Count, Spec, Recursive
 If Count Then
  ReDim Preserve Arr(1 To Count)
 End If
End Sub

Private Sub EFIStrArrWC(ByVal StartPath As String, _
          ByRef Arr() As String, _
          ByRef Count As Long, _
          ByVal Spec As String, _
          ByVal Recursive As Boolean)

 Dim Valid As Boolean
 Dim FoundFile As String
 Dim hFile As Long
 Dim mFD As WIN32_FIND_DATA
 'this can be removed for other uses
 With frmBuild.lblPath
  .Caption = StartPath
  .Refresh
 End With
 hFile = FindFirstFile(StartPath & SDS, mFD)
 Valid = (hFile <> INVALID_HANDLE_VALUE)
 Do While Valid
  FoundFile = TrimNull(mFD.cFileName)
  If (mFD.dwFileAttributes And vbDirectory) Then
   If AscW(FoundFile) <> VBDOT Then
    If Recursive Then
     EFIStrArrWC StartPath & FoundFile & "\", Arr, Count, Spec, Recursive
    End If
   End If
  Else
   If PathMatchSpec(FoundFile, Spec) Then
    Count = Count + 1
    If Count > UBound(Arr) Then
     ReDim Preserve Arr(1 To Count + ARR_MAX)
    End If
    Arr(Count) = StartPath & FoundFile
   End If
  End If
  Valid = FindNextFile(hFile, mFD)
 Loop
 FindClose hFile
End Sub


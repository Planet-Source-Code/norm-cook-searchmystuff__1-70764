VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mPath As String
Private mKey As String
Private mSection As String
Private mDefault As String

' Profile String functions:
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Property Let Default(sDefault As String)
 mDefault = sDefault
End Property
Property Get Default() As String
 Default = mDefault
End Property
Property Let Path(sPath As String)
 mPath = sPath
End Property
Property Get Path() As String
 Path = mPath
End Property
Property Let Key(sKey As String)
 mKey = sKey
End Property
Property Get Key() As String
 Key = mKey
End Property
Property Let Section(sSection As String)
 mSection = sSection
End Property
Property Get Section() As String
 Section = mSection
End Property
Property Get Value() As String
 Dim sBuf As String * 255
 Dim RetLen As Long
 RetLen = GetPrivateProfileString(mSection, mKey, mDefault, sBuf, &O255, mPath)
 Value = Left$(sBuf, RetLen)
End Property
Property Let Value(sValue As String)
 sValue = Replace$(sValue, vbNullChar, vbNullString)
 WritePrivateProfileString mSection, mKey, sValue, mPath
End Property
Public Sub DeleteKey()
 WritePrivateProfileString mSection, mKey, 0&, mPath
End Sub
Public Sub DeleteSection()
 WritePrivateProfileString mSection, 0&, 0&, mPath
End Sub
Private Property Get INISection() As String
 Dim sBuf As String * 8192
 Dim RetLen As Long
 RetLen = GetPrivateProfileString(mSection, 0&, mDefault, sBuf, 8192&, mPath)
 INISection = Left$(sBuf, RetLen - 1)
End Property
Private Property Get Sections() As String
 Dim sBuf As String * 8192
 Dim RetLen As Long
 RetLen = GetPrivateProfileString(0&, 0&, mDefault, sBuf, 8192&, mPath)
 Sections = Left$(sBuf, RetLen - 1)
End Property
Public Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef Count As Long)
 Dim Tmp() As String
 Dim sSection As String
 Dim i As Long
 Count = 0
 Erase sKey
 sSection = INISection
 If Len(sSection) Then
  Tmp = Split(sSection, vbNullChar)
  Count = UBound(Tmp) + 1
  ReDim sKey(1 To Count)
  For i = 0 To UBound(Tmp)
   sKey(i + 1) = Tmp(i)
  Next
 End If
End Sub
Public Sub EnumerateAllSections(ByRef sSections() As String, ByRef Count As Long)
 Dim sIniFile As String
 Dim i As Long
 Dim Tmp() As String
 Count = 0
 Erase sSections
 sIniFile = Sections
 If Len(sIniFile) Then
  Tmp = Split(sIniFile, vbNullChar)
  Count = UBound(Tmp) + 1
  ReDim sSections(1 To Count)
  For i = 0 To UBound(Tmp)
   sSections(i + 1) = Tmp(i)
  Next
 End If
End Sub


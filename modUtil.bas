Attribute VB_Name = "modUtil"
Option Explicit
Public oIni As cIni
Private Const LVM_FIRST = &H1000
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function LVSelCount(LVW As ListView)
 LVSelCount = SendMessage(LVW.hWnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
End Function
Public Function Quote(ByVal s As String) As String
 Quote = Chr$(34) & s & Chr$(34)
End Function
Public Function StripAttr(ByVal Txt As String) As String
 Dim a() As String, i As Long, p As Long
 On Error GoTo Errhdl
 a = Split(Txt, vbNewLine)
 For i = 0 To UBound(a) 'find 1st occurrence
  If Left$(a(i), 13) = "Attribute VB_" Then
   Exit For
  End If
 Next
 If i = UBound(a) + 1 Then 'never found it
  StripAttr = Txt
  Exit Function
 End If
 i = i + 1 'search until it doesn't appear
 Do Until Left$(a(i), 13) <> "Attribute VB_"
  i = i + 1
 Loop
 p = InStr(Txt, a(i))
 StripAttr = Mid$(Txt, p)
 Exit Function
Errhdl:
 StripAttr = Txt
End Function




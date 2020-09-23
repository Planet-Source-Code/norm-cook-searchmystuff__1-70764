VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmView 
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   Icon            =   "frmView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdFont 
      Caption         =   "Change Font"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6870
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17304
            Text            =   "Press 'n' for next occurrence of search string"
            TextSave        =   "Press 'n' for next occurrence of search string"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Search As String
Public Text As String
Private CurrPos As Long
Private NotInText As Boolean

Private Sub cmdFont_Click()
 With CD
  .Flags = cdlCFBoth
  On Error Resume Next
  .ShowFont
  If Err = cdlCancel Then Exit Sub
  Txt.Font.Name = .FontName
  Txt.Font.Size = .FontSize
  Txt.Font.Bold = .FontBold
  Txt.Font.Italic = .FontItalic
 End With
End Sub

Private Sub Form_Activate()
 Txt.Text = Text
 Text = LCase$(Text)
 Search = LCase(Search)
 If Len(Search) Then
  CurrPos = InStr(1, Text, Search)
  If CurrPos Then
   Txt.SelStart = CurrPos - 1
   Txt.SelLength = Len(Search)
  Else
   SB.Panels(1).Text = "Search String not in text"
   NotInText = True
  End If
 Else
  SB.Panels(1).Text = "No Search String specified"
  NotInText = True
 End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
  KeyAscii = 0
  FindNext
 End If
End Sub

Private Sub FindNext()
 If NotInText Then Exit Sub
 CurrPos = InStr(CurrPos + Len(Search), Text, Search)
 If CurrPos Then
  Txt.SelStart = CurrPos - 1
  Txt.SelLength = Len(Search)
 Else
  CurrPos = InStr(1, Text, Search)
  Txt.SelStart = CurrPos - 1
  Txt.SelLength = Len(Search)

 End If
End Sub

Private Sub Form_Load()
 Dim i As Long
 With oIni
  .Section = "View Settings"
  .Key = "FLeft": Left = .Value
  .Key = "FTop": Top = .Value
  .Key = "FWidth": Width = .Value
  .Key = "FHeight": Height = .Value
  .Key = "FontN":  Txt.Font.Name = .Value
  .Key = "FontS":  Txt.Font.Size = .Value
  .Key = "FontB":  Txt.Font.Bold = .Value
  .Key = "FontI":  Txt.Font.Italic = .Value
 End With

End Sub

Private Sub Form_Resize()
 If WindowState <> vbMinimized Then
  Txt.Move 0, cmdFont.Height, ScaleWidth, ScaleHeight - SB.Height - cmdFont.Height
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim i As Long
 With oIni
  .Section = "View Settings"
  .Key = "FLeft": .Value = Left
  .Key = "FTop": .Value = Top
  .Key = "FWidth": .Value = Width
  .Key = "FHeight": .Value = Height
  .Key = "FontN": .Value = Txt.Font.Name
  .Key = "FontS": .Value = Txt.Font.Size
  .Key = "FontB": .Value = Txt.Font.Bold
  .Key = "FontI": .Value = Txt.Font.Italic
 End With
End Sub

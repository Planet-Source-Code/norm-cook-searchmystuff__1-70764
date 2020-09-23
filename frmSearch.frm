VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search VB"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10590
   Begin VB.CheckBox chkRemAttr 
      Caption         =   "Remove VB Attributes When Viewing"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":08CA
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":0E64
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":13FE
            Key             =   "frm"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1998
            Key             =   "bas"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":1F32
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":24CC
            Key             =   "Dn"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSrch 
      Caption         =   "Search"
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.CheckBox chkT 
      Caption         =   "Controls"
      Height          =   255
      Index           =   3
      Left            =   7800
      TabIndex        =   6
      Tag             =   "ctl"
      Top             =   0
      Width           =   975
   End
   Begin VB.CheckBox chkT 
      Caption         =   "Cls Mods"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   5
      Tag             =   "cls"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox chkT 
      Caption         =   "Bas Mods"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   4
      Tag             =   "bas"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CheckBox chkT 
      Caption         =   "Forms"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Tag             =   "frm"
      Top             =   0
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Show All Records"
      Height          =   255
      Left            =   9000
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtSrch 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2646
         ImageKey        =   "Up"
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Path"
         Object.Width           =   6703
      EndProperty
   End
   Begin VB.Menu LVPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu LVArr 
         Caption         =   "View Selected File"
         Index           =   0
      End
      Begin VB.Menu LVArr 
         Caption         =   "Copy To..."
         Index           =   1
      End
      Begin VB.Menu LVArr 
         Caption         =   "Delete From Database"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program needs a reference to
' Microsoft DAO Object Library (dao360.dll)
'If it shows as missing reference, you
' may need to register it,
' Start|Run|regsvr32.exe c:\[path]\dao360.dll
'Possible Locations:
'C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll
'C:\WINDOWS\ServicePackFiles\i386
'Download site
'http://www.domainpunch.com/support/articles/dao.php
'Do NOT need MS Access installed
'Note you can also drag files to an open
' VB Ide Project Window, see the
' LV_OLEStartDrag routine below
Option Explicit
Private DBPath As String
Private DB As DAO.Database
Private RecCnt As Long
Private LVRatios As Variant
Private Sub Form_Load()
 Init
 Show
 DoEvents
 If RecCnt = 0 Then
  MsgBox "No records in the database." & vbNewLine & _
    "Run the included BuildMyStuff vbp to add records"
  Unload Me
 Else
  LV.ColumnHeaders(5).Text = "Path   (Database has " & RecCnt & " records)"
 End If
End Sub
Private Sub Init()
 DBPath = App.Path & "\Ref.mdb"
 Set DB = OpenDatabase(DBPath)
 RecCnt = DBRecCnt
 Set oIni = New cIni
 oIni.Path = App.Path & "\Search.ini"
 LVRatios = Array(, 0.142, 0.1, 0.164, 0.136, 0.45)
 LoadSettings
End Sub
Private Sub LoadSettings()
 Dim i As Long
 With oIni
  .Section = "Main Settings"
  .Key = "FLeft": Left = .Value
  .Key = "FTop": Top = .Value
  .Key = "FWidth": Width = .Value
  .Key = "FHeight": Height = .Value
  For i = 0 To 3
   .Key = "FType" & i: chkT(i).Value = .Value
  Next
  .Key = "FAttr":  chkRemAttr.Value = .Value
 End With
End Sub
Private Sub SaveSettings()
 Dim i As Long
 With oIni
  .Section = "Main Settings"
  .Key = "FLeft": .Value = Left
  .Key = "FTop": .Value = Top
  .Key = "FWidth": .Value = Width
  .Key = "FHeight": .Value = Height
  For i = 0 To 3
   .Key = "FType" & i: .Value = chkT(i).Value
  Next
  .Key = "FAttr": .Value = chkRemAttr.Value
 End With
End Sub
Private Sub Form_Resize()
 Dim i As Long
 If WindowState <> vbMinimized Then
 'listview column width ratios
 LV.Move 0, 480, ScaleWidth, ScaleHeight
 For i = 1 To LV.ColumnHeaders.Count
  LV.ColumnHeaders(i).Width = LVRatios(i) * LV.Width
 Next
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveSettings
 Set oIni = Nothing
 DB.Close
 Set DB = Nothing
End Sub
Private Sub cmdAll_Click()
 LoadRS DB.OpenRecordset("Main")
End Sub
Private Sub cmdSrch_Click()
 If Len(txtSrch.Text) Then
  DoSearch
 End If
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 Dim i As Long
 For i = 1 To LV.ColumnHeaders.Count
  LV.ColumnHeaders(i).Icon = 0
 Next
 With LV
  .SortKey = ColumnHeader.Index - 1
  .SortOrder = .SortOrder Xor 1
  LV.ColumnHeaders(ColumnHeader.Index).Icon = IIf(.SortOrder = lvwAscending, "Up", "Dn")
 End With
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
 Caption = Item.SubItems(4)
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim LI As ListItem
 If Button = vbRightButton Then
  Set LI = LV.HitTest(x, y)
  If Not LI Is Nothing Then
   LI.Selected = True
   LVArr(0).Enabled = LVSelCount(LV) < 2
   PopupMenu LVPop
  End If
 End If
End Sub

Private Sub LV_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
 Dim i As Long
 For i = 1 To LV.ListItems.Count
  With LV.ListItems(i)
   If .Selected Then
    Data.Files.Add .SubItems(4)
   End If
  End With
 Next
 Data.SetData , vbCFFiles
End Sub

Private Sub LVArr_Click(Index As Integer)
 Select Case Index
  Case 0 'view
   DoView LV.SelectedItem
  Case 1 'copy
   DoCopy
  Case 2 'delete
   DoDelete
 End Select
End Sub
Private Sub DoDelete()
 Dim i As Long
 Dim RS As Recordset
 For i = LV.ListItems.Count To 1 Step -1
  With LV.ListItems(i)
   If .Selected Then
    Set RS = DB.OpenRecordset("Select * From Main Where Index = " & Val(.Key))
    RS.Delete
    LV.ListItems.Remove i
   End If
  End With
 Next
End Sub
Private Sub DoView(LI As ListItem)
 Dim Frm As Form
 Dim RS As Recordset
 Set RS = DB.OpenRecordset("Select * From Main Where Index = " & Val(LI.Key))
 Set Frm = New frmView
 Frm.Caption = RS!Path
 If chkRemAttr.Value = vbChecked Then
  Frm.Text = StripAttr(RS!Text)
 Else
  Frm.Text = RS!Text
 End If
 Frm.Search = txtSrch.Text
 Frm.Show vbModal
End Sub
Private Sub txtSrch_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  KeyAscii = 0
  DoSearch
 End If
End Sub
Private Sub DoSearch()
 Dim RS As Recordset
 Dim SQL As String
 Dim Srch As String
 Srch = Quote("*" & txtSrch.Text & "*")
 SQL = "Select * From Main Where (Text Like " & Srch & " Or Name Like " & Srch & ")" & _
   " And (" & TypeSQL & ")"
 Set RS = DB.OpenRecordset(SQL)
 If RS.EOF Then
  Caption = "Nothing Found"
  Exit Sub
 End If
 LoadRS RS
End Sub
Private Function TypeSQL() As String
 Dim i As Long
 For i = 0 To 3
  If chkT(i).Value = vbChecked Then
   If Len(TypeSQL) Then
    TypeSQL = TypeSQL & "Or Type = " & Quote(chkT(i).Tag)
   Else
    TypeSQL = TypeSQL & " Type = " & Quote(chkT(i).Tag)
   End If
   TypeSQL = TypeSQL & " "
  End If
 Next
End Function
Private Sub LoadRS(RS As Recordset)
 Dim i As Long
 Screen.MousePointer = vbHourglass
 LV.ListItems.Clear
 RS.MoveLast: RS.MoveFirst
 For i = 1 To RS.RecordCount
  With LV.ListItems.Add
   .Selected = False
   .Text = RS!Name
   .SubItems(1) = RS!Type
   .SmallIcon = CStr(RS!Type)
   .SubItems(2) = Format$(RS!Date, "yyyy-mm-dd hh-nn-ss")
   .SubItems(3) = Format$(RS!Size, "@@@@@@@")
   .SubItems(4) = RS!Path
   .Key = RS!Index & "k"
  End With
  RS.MoveNext
 Next
 Screen.MousePointer = vbDefault
 Caption = LV.ListItems.Count & " Items Found"
End Sub
Private Sub DoCopy()
 Dim RS As Recordset
 Dim BF As String, i As Long
 BF = BrowseForFolderByPath("C:\VB6", hWnd, "Select Folder")
 If Len(BF) = 0 Then Exit Sub
 For i = 1 To LV.ListItems.Count
  With LV.ListItems(i)
   If .Selected Then
    Set RS = DB.OpenRecordset("Select Path From Main Where Index = " & Val(.Key))
    FileCopy RS!Path, QualifyPath(BF) & FileTitle(RS!Path)
   End If
  End With
 Next
End Sub
Private Function DBRecCnt() As Long
 Dim RS As Recordset
 Set RS = DB.OpenRecordset("Main")
 If Not RS.EOF Then
  RS.MoveLast: RS.MoveFirst
  DBRecCnt = RS.RecordCount
 End If
 RS.Close
End Function

VERSION 5.00
Begin VB.Form frmBuild 
   Caption         =   "Build Reference Database"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "frmBuild.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear Database"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Empties the Database before populating"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate Database"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      ToolTipText     =   "Store the files in the database"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Start Scan"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Scan Path above for the selected files"
      Top             =   1545
      Width           =   3015
   End
   Begin VB.TextBox txtScanPath 
      Height          =   285
      Left            =   1635
      TabIndex        =   6
      Text            =   "C:\"
      Top             =   1065
      Width           =   4095
   End
   Begin VB.CommandButton cmdBrScan 
      Caption         =   "..."
      Height          =   285
      Left            =   5715
      TabIndex        =   5
      ToolTipText     =   "Change the Scan Path"
      Top             =   1065
      Width           =   375
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "Usercontrols"
      Height          =   255
      Index           =   3
      Left            =   4830
      TabIndex        =   3
      Tag             =   "*.ctl"
      Top             =   465
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "Class Modules"
      Height          =   255
      Index           =   2
      Left            =   3190
      TabIndex        =   2
      Tag             =   "*.cls"
      Top             =   465
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "Bas Modules"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Tag             =   "*.bas"
      Top             =   465
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkExt 
      Caption         =   "Forms"
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   0
      Tag             =   "*.frm"
      Top             =   465
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Label lblPath 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label Label3 
      Caption         =   "Start Scanning At:"
      Height          =   255
      Left            =   195
      TabIndex        =   7
      Top             =   1125
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Scan For:"
      Height          =   255
      Left            =   1695
      TabIndex        =   4
      Top             =   105
      Width           =   2895
   End
End
Attribute VB_Name = "frmBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program needs a reference to dao360.dll
'Do not need to have MS Access installed
'May need to register it
'Possible Locations:
'C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll
'C:\WINDOWS\ServicePackFiles\i386
'Download site
'http://www.domainpunch.com/support/articles/dao.php

'As written, the db gets cleared before each populate
'If you have numerous sources, just uncheck
'the 'Clear Database' checkbox each time
'you scan/populate.  You will likely wind up
'with a lot of dupes, as I did, but the
'companion app allows you to delete entries
Option Explicit
Private DB As DAO.Database
Private DBPath As String
Private Files() As String
Private FCnt As Long
Private Sub Form_Load()
 'rename the file, change the path as desired
 DBPath = App.Path & "\Ref.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set DB = Nothing
End Sub

Private Sub cmdBrScan_Click()
 Dim Br As String
 Br = BrowseForFolderByPath(txtScanPath.Text, hWnd, "Select Path to Start Scanning")
 If Len(Br) Then
  txtScanPath.Text = Br
 End If
End Sub

Private Function GetSpec() As String
 Dim i As Long
 For i = 0 To 3
  If chkExt(i).Value = vbChecked Then
   If i <> 3 Then
    GetSpec = GetSpec & chkExt(i).Tag & "; "
   Else
    GetSpec = GetSpec & chkExt(i).Tag
   End If
  End If
 Next
End Function
Private Sub cmdScan_Click()
 DoScan
End Sub
Private Sub cmdPopulate_Click()
 Populate
End Sub
Private Sub DoScan()
 Dim Spec As String
 Spec = GetSpec
 If Len(Spec) = 0 Then
  MsgBox "No VB File types selected"
  chkExt(0).SetFocus
  Exit Sub
 End If
 If (Len(txtScanPath.Text) = 0) Or _
    (FolderExists(txtScanPath.Text) = False) Then
  MsgBox "Invalid Scan Path"
  txtScanPath.SetFocus
  Exit Sub
 End If
 'get the files
 EnumFilesStringArrayWildCard txtScanPath.Text, Files, FCnt, Spec, True
 lblPath.Caption = "Done. " & FCnt & " Files Found"
 cmdPopulate.Enabled = CBool(FCnt)
 cmdScan.Enabled = False
End Sub
Private Sub DeletePrev()
 Dim i As Long
 Dim RS As Recordset
 Set RS = DB.OpenRecordset("Main")
 RS.MoveLast
 RS.MoveFirst
 For i = 1 To RS.RecordCount
  RS.Delete
  RS.MoveNext
 Next
 RS.Close
End Sub
Private Sub Populate()
 Dim RS As DAO.Recordset
 Dim i As Long
 Dim FTitle As String
 Set DB = OpenDatabase(DBPath)
 If chkClear.Value = vbChecked Then
  DeletePrev
 End If
 Set RS = DB.OpenRecordset("Main")
 For i = 1 To FCnt
  lblPath.Caption = "Processing File " & i & " of " & FCnt
  lblPath.Refresh
  FTitle = FileTitle(Files(i))
  RS.AddNew
  RS.Fields("Path").Value = Files(i)
  RS.Fields("Type").Value = LCase$(Right$(Files(i), 3))
  RS.Fields("Name").Value = Left$(FTitle, Len(FTitle) - 4)
  RS.Fields("Text").Value = ReadFileBinary(Files(i))
  RS.Fields("Date").Value = FileDateTime(Files(i))
  RS.Fields("Size").Value = FileLen(Files(i))
  RS.Update
 Next
 RS.Close
 lblPath.Caption = lblPath.Caption & vbNewLine & "Compacting Database"
 DB.Close
 CompactDB DBPath
 lblPath.Caption = lblPath.Caption & vbNewLine & "Done!"
 cmdScan.Enabled = True
 cmdPopulate.Enabled = False
End Sub
Private Sub KillExists(ByVal FilePath As String)
 On Error Resume Next
 Kill FilePath
End Sub
Private Sub CompactDB(ByVal FilePath As String)
 Dim TmpFN As String
 TmpFN = App.Path & "\Temp.mdb"
 KillExists TmpFN
 CompactDatabase FilePath, TmpFN, dbLangGeneral
 KillExists FilePath
 Name TmpFN As FilePath
End Sub


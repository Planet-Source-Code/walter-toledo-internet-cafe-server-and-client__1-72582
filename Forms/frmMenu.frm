VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC Internet Cafe"
   ClientHeight    =   7950
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13830
   ClipControls    =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2400
      Top             =   8040
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtTriger 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   8040
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   5175
      Left            =   12480
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdVoid 
         Caption         =   "&Void"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdTender 
         Caption         =   "T&ender"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PC"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Account Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Service Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Log-In Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Log-Out Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Amount    "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Access Key"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   35
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3D84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   1508
      ButtonWidth     =   1270
      ButtonHeight    =   1455
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limited"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transfer"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Extend"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "-"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   12255
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Caption         =   " Service Rate "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtGame 
            Height          =   285
            Left            =   1560
            TabIndex        =   9
            Text            =   "25.00"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtNet 
            Height          =   285
            Left            =   1560
            TabIndex        =   8
            Text            =   "20.00"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Games/Rental:"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Internet: "
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Today is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   720
         Width           =   9255
      End
      Begin VB.Label lblToday 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Today is"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11760
      Top             =   960
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOP 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuLT 
         Caption         =   "Limited"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTR 
         Caption         =   "Transfer"
      End
      Begin VB.Menu mnuEx 
         Caption         =   "Extend"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New Recordset

Private Sub Form_Activate()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Connect ' Database Connection

lblToday.Caption = Format(Date, "dddd, mmmm dd, yyyy")
lblTime.Caption = Format(Time, "h:mm:ss AMPM")

LoadME
End Sub

Sub LoadME()
SQL = "SELECT * FROM log_temp ORDER BY id"
If rst.State = 1 Then rst.Close
rst.Open SQL, con, 3, 3, 1
lv1.ListItems.Clear
For i = 1 To rst.RecordCount
  With lv1.ListItems.Add(, , rst!pc)
    .ListSubItems.Add , , UCase(rst!Name)
    .ListSubItems.Add , , rst!Account
    .ListSubItems.Add , , rst!Service
    .ListSubItems.Add , , Format(rst!Time_in, "hh:nn AMPM")
    .ListSubItems.Add , , Format(rst!Time_out, "hh:nn AMPM")
    .ListSubItems.Add , , rst!Duration
    .ListSubItems.Add , , Format(rst!Amount, "0.00")
    .ListSubItems.Add , , rst!access
    .ListSubItems.Add , , rst!IDCode
  End With
  rst.MoveNext
Next i
End Sub

Private Sub cmdTender_Click()
'On Error GoTo Hell
Dim i As Integer
i = txtID
frmTender.lblComp.Caption = "Computer No. " & txtID.Text 'lv1.ListItems(i).ListSubItems().Text
frmTender.txtService.Text = lv1.ListItems(i).ListSubItems(3).Text

frmTender.txtID = lv1.ListItems(i).ListSubItems(9).Text

If lv1.ListItems(i).ListSubItems(2).Text = "Open" Then
'Dim a, b, c
Dim x, y, z

x = lv1.ListItems(i).ListSubItems(4).Text
y = Format(Time, "hh:mm AMPM")

frmTender.dtStart = x
frmTender.dtEnd = y

x = frmTender.dtStart
y = frmTender.dtEnd

z = y - x
frmTender.dtDuration = z
frmTender.txtTrigger = Format(z, "hh:mm AMPM")
frmTender.txtDuration = frmTender.dtDuration

frmTender.txtDuration = Mid(frmTender.txtTrigger, 1, 5)
'frmTender.dtEnd = Time

dtEnd = z 'Time 'frmTender.dtEnd

cboService = lv1.ListItems(i).ListSubItems(3).Text
If cboService = "Internet" Then
  txtFee = frmMenu.txtNet
Else
  txtFee = frmMenu.txtGame
End If

a = Format(dtEnd, "hh")
b = Format(dtEnd, "nn")

If a <> 0 Then
  If a > 12 Then
    a = 24 / a
    a = Round(a)
    a = a * txtFee
  Else
    a = a * txtFee
  End If
End If

If b <> 0 Then
  c = txtFee / 60
  b = b * c '0.333333333
End If

txtAmount = Val(Round(a, 2)) + Val(Round(b, 2))
frmTender.txtAmount.Text = Format(txtAmount, "#,##0.00")
Else
frmTender.txtDuration.Text = Mid(lv1.ListItems(i).ListSubItems(6).Text, 1, 5)
frmTender.txtAmount.Text = lv1.ListItems(i).ListSubItems(7).Text
End If

frmTender.Show 1

Exit Sub
Hell:
MsgBox Err.Description, vbInformation, "Info Message"
End Sub

Private Sub cmdVoid_Click()
Dim i As Integer
Dim rs1 As New Recordset

i = txtID
id = lv1.ListItems(i).ListSubItems(9).Text
If MsgBox("This void this transaction?", vbQuestion + vbYesNo + vbDefaultButton2, "Info Message") = vbYes Then
  For j = 1 To 9
    lv1.ListItems(i).ListSubItems(j).Text = ""
  Next
  
  SQL = "SELECT * FROM log WHERE id='" & id & "'"
  If rs.State = 1 Then rs.Close
  rs.Open SQL, con, 3, 3, 1
  If rs.RecordCount <> 0 Then
    With rs
      !Status = "Void"
      .Update
    End With
  End If
  
  SQL = "SELECT * FROM log_temp WHERE id='" & txtID & "'"
  If rs1.State = 1 Then rs1.Close
  rs1.Open SQL, con, 3, 3, 1
  For i = 1 To rs1.RecordCount
    rs1!Name = ""
    rs1!Account = ""
    rs1!Service = ""
    rs1!Time_in = "00:00:00"
    rs1!Time_out = "00:00:00"
    rs1!Duration = ""
    rs1!Amount = "0"
    rs1!access = ""
    rs1!IDCode = ""
    rs1.Update
    rs1.MoveNext
  Next
Else
  Exit Sub
End If
LoadME
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure you want exit the application?", vbQuestion + vbYesNo + vbDefaultButton2, "Info Message") = vbYes Then
  End
Else
  Cancel = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
End Sub

Private Sub lv1_Click()
On Error Resume Next
Dim i As Integer
i = lv1.SelectedItem.Text
txtID = i
txtTriger.Text = lv1.ListItems(i).ListSubItems(2).Text
If txtTriger.Text <> "" Then
  cmdTender.Enabled = True
  cmdVoid.Enabled = True
Else
  cmdTender.Enabled = False
  cmdVoid.Enabled = False
End If
End Sub

Private Sub lv1_Keyup(KeyCode As Integer, Shift As Integer)
lv1_Click
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuEx_Click()
  Extend
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuExport_Click()
  frmExport.Show
End Sub

Private Sub mnuLT_Click()
  frmAccount.txtAccount.Text = "Limited"
  frmAccount.Show 1
End Sub

Private Sub mnuOP_Click()
  frmAccount.txtAccount.Text = "Open"
  frmAccount.Show 1
End Sub

Private Sub mnuReport_Click()
frmReport1.Show
End Sub

Private Sub mnuTR_Click()
  frmTransfer.Show 1
End Sub

Private Sub Timer1_Timer()
lblTime = Time
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 2
    frmAccount.txtAccount.Text = "Open"
    frmAccount.Show 1
  Case 3
    frmAccount.txtAccount.Text = "Limited"
    frmAccount.Show 1
  Case 5
    frmTransfer.Show 1
  Case 6
    Extend
    'frmExtend.Show 1
End Select
End Sub

Sub Extend()
On Error Resume Next
Dim j As Integer
j = txtID
If lv1.ListItems(j).ListSubItems(2).Text = "" Then
  MsgBox "Empty or Null Value, Check your account!", vbInformation, "Info Message"
  Exit Sub
Else
  If lv1.ListItems(j).ListSubItems(2).Text = "Open" Then
    MsgBox "You can only extend limited account!", vbInformation, "Info Message"
    Exit Sub
  Else
    frmExtend.Show 1
  End If
End If
End Sub

Private Sub cmdClear_Click()
If MsgBox("You really want to clear the log-files?", vbQuestion + vbYesNo + vbDefaultButton2, "Info Message") = vbYes Then
  ClearAll
Else
  Exit Sub
End If
End Sub

Sub ClearAll()
SQL = "SELECT * FROM log_temp"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
For i = 1 To rs.RecordCount
  rs!Name = ""
  rs!Account = ""
  rs!Service = ""
  rs!Time_in = "00:00:00"
  rs!Time_out = "00:00:00"
  rs!Duration = ""
  rs!Amount = "0"
  rs!access = ""
  rs!IDCode = ""
  rs.Update
  rs.MoveNext
Next
LoadME
End Sub

Private Sub Timer2_Timer()
TriggerMan
End Sub

Sub TriggerMan()
Dim x, y
SQL = "SELECT * FROM log_temp"
If rs.State = 1 Then rs.Close
rs.Open SQL, con
For i = 1 To rs.RecordCount
  x = Format(Time, "hh:mm AMPM")
  y = lv1.ListItems(i).ListSubItems(5).Text
  If y = x Then
    With lv1.ListItems(i)
      .ListSubItems(1).ForeColor = &HFF&
      .ListSubItems(2).ForeColor = &HFF&
      .ListSubItems(3).ForeColor = &HFF&
      .ListSubItems(4).ForeColor = &HFF&
      .ListSubItems(5).ForeColor = &HFF&
      .ListSubItems(6).ForeColor = &HFF&
      .ListSubItems(7).ForeColor = &HFF&
      .ListSubItems(8).ForeColor = &HFF&
    End With
  End If
Next i
End Sub


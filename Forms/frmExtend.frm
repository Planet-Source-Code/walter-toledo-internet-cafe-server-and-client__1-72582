VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExtend 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extend Account"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   Icon            =   "frmExtend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboComp 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmExtend.frx":014A
      Left            =   1680
      List            =   "frmExtend.frx":014C
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtAccount 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox cboService 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmExtend.frx":014E
      Left            =   1680
      List            =   "frmExtend.frx":0158
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtFee 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtDuration 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777215
      CustomFormat    =   "00:00:00"
      Format          =   76677122
      CurrentDate     =   40071.8640509259
   End
   Begin MSComCtl2.DTPicker dtStart 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "hh:mm AMPM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   76677122
      CurrentDate     =   40069
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      CalendarBackColor=   16777215
      CustomFormat    =   "00:00:00"
      Format          =   76677122
      CurrentDate     =   40073.8306597222
   End
   Begin VB.TextBox txtDuration 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Log-Out Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4200
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Fee As String

Private Sub cboComp_Click()
'txtName.SetFocus
End Sub

Private Sub cboService_Change()
 If cboService = "Internet" Then
   txtFee = frmMenu.txtNet
 Else
   txtFee = frmMenu.txtGame
 End If
End Sub

Private Sub cboService_Click()
'cboService_Change
End Sub

Private Sub cmdCancel_Click()
If cmdCancel.Caption = "&Cancel" Then
  Unload Me
End If

If cmdCancel.Caption = "&Back" Then
  'cmdOK.Caption = "&Ok"
  cmdConfirm.Visible = False
  cmdCancel.Caption = "&Cancel"
  EnabledME True
End If
End Sub

Private Sub cmdOK_Click()
If cboComp.Text <> "" Then
  If cmdOK = True Then
    cmdConfirm.Visible = True
    'cmdOK.Caption = "&Confirm"
    cmdCancel.Caption = "&Back"
    EnabledME False
  End If
Else
  MsgBox "Required fields need to filled-up!", vbInformation, "Info Message"
End If
End Sub

Private Sub dtDuration_Change()
  txtDuration = dtDuration
End Sub

Private Sub dtDuration_KeyPress(KeyAscii As Integer)
  KeyAscii = 0
End Sub

Private Sub dtEnd_Change()
'On Error Resume Next
Dim a, b, c

dtDuration = dtStart + dtEnd
txtDuration = Format(dtDuration, "hh:mm AMPM")

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

txtAmount.Text = Val(Round(a, 2)) + Val(Round(b, 2))
txtAmount.Text = Format(txtAmount, "#,##0.00")
End Sub

Private Sub dtEnd_KeyPress(KeyAscii As Integer)
dtEnd_Change
End Sub

Private Sub dtEnd_KeyUp(KeyCode As Integer, Shift As Integer)
'dtEnd_Change
End Sub

Private Sub Form_Activate()
'If txtAccount = "Open" Then
'  dtEnd = "00:00"
'  dtDuration = dtStart
'  txtAmount = "0.00"
'End If
End Sub

Private Sub Form_Load()
Dim j As Integer
j = frmMenu.txtID.Text
With frmMenu
  cboComp.Text = j
  txtName.Text = .lv1.ListItems(j).ListSubItems(1).Text
  txtAccount.Text = .lv1.ListItems(j).ListSubItems(2).Text
  cboService.Text = .lv1.ListItems(j).ListSubItems(3).Text
  dtStart = .lv1.ListItems(j).ListSubItems(4).Text
  dtDuration = .lv1.ListItems(j).ListSubItems(5).Text
  dtEnd = .lv1.ListItems(j).ListSubItems(6).Text
  txtAmount.Text = .lv1.ListItems(j).ListSubItems(7).Text
  'txtName.Text = .lv1.ListItems(j).ListSubItems(1).Text
End With
End Sub

Private Sub cmdConfirm_Click()
LoadME
Unload Me
frmMenu.LoadME
End Sub

Public Sub LoadME()
On Error GoTo Hell
Dim i As Integer
i = cboComp.Text
'With frmMenu.lv1.ListItems(i)
  '.ListSubItems(1) = UCase(txtName.Text)
  '.ListSubItems(2) = txtAccount
  '.ListSubItems(3) = cboService
  '.ListSubItems(4) = Format(dtStart, "hh:nn AMPM")
  '.ListSubItems(5) = Format(dtDuration, "hh:nn AMPM")
  '.ListSubItems(6) = Format(dtEnd, "hh:nn AMPM")
  '.ListSubItems(7) = txtAmount
  '.ListSubItems(8) = txtAmount
  '.ListSubItems(9) = txtAmount
'End With
id = frmMenu.lv1.ListItems(i).ListSubItems(9).Text
SQL = "SELECT * FROM log WHERE id='" & id & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
If rs.RecordCount <> 0 Then
  rs!Time_in = Format(dtStart, "hh:nn AMPM")
  rs!Time_out = Format(dtDuration, "hh:nn AMPM")
  rs!Duration = Format(dtEnd, "hh:nn AMPM")
  rs!Amount = txtAmount
  rs.Update
End If
SQL = "SELECT * FROM log_temp WHERE id='" & i & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
If rs.RecordCount <> 0 Then
  rs!Time_in = Format(dtStart, "hh:nn AMPM")
  rs!Time_out = Format(dtDuration, "hh:nn AMPM")
  rs!Duration = Format(dtEnd, "hh:nn AMPM")
  rs!Amount = txtAmount
  rs.Update
End If

Exit Sub
Hell:
MsgBox Err.Description, vbInformation, "Info Message"
End Sub

Sub EnabledME(Hit As Boolean)
'cboComp.Enabled = Hit
'txtName.Enabled = Hit
'txtAccount.Enabled = Hit
'cboService.Enabled = Hit
dtStart.Enabled = Hit
dtEnd.Enabled = Hit
dtDuration.Enabled = Hit
txtAmount.Enabled = Hit
End Sub



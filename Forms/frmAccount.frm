VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAccount 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Type"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm"
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtFee 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtDuration 
      Height          =   315
      Left            =   1680
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777215
      CustomFormat    =   "00:00:00"
      Format          =   20643842
      CurrentDate     =   40071.8640509259
   End
   Begin VB.TextBox txtDuration 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3600
      Width           =   1335
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
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   40069
   End
   Begin VB.ComboBox cboService 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmAccount.frx":014A
      Left            =   1680
      List            =   "frmAccount.frx":0154
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtAccount 
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox cboComp 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmAccount.frx":0169
      Left            =   1680
      List            =   "frmAccount.frx":016B
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1680
      TabIndex        =   13
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
      Format          =   20643842
      CurrentDate     =   40073.8306597222
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4200
      Y1              =   4080
      Y2              =   4080
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
      TabIndex        =   7
      Top             =   3600
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
      TabIndex        =   6
      Top             =   2640
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
      TabIndex        =   5
      Top             =   3120
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
      TabIndex        =   4
      Top             =   2160
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
      TabIndex        =   3
      Top             =   1680
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
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
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
      TabIndex        =   1
      Top             =   720
      Width           =   1455
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
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code As String

Private Sub Form_Activate()
If txtAccount = "Open" Then
  dtEnd = "00:00"
  dtDuration = dtStart
  txtAmount = "0.00"
End If
End Sub

Private Sub Form_Load()
Dim h As String
Dim m As String
Dim s As String

dtStart = Format(Time, "hh:mm AMPM")
dtEnd = "01:00"
txtDuration = dtStart + dtEnd
dtDuration = Format(txtDuration, "hh:mm AMPM")

cboService = "Internet"

txtAmount = "20.00"

For i = 1 To frmMenu.lv1.ListItems.Count
  If frmMenu.lv1.ListItems(i).ListSubItems(2).Text = "" Then
    cboComp.AddItem i
  End If
Next
End Sub

Private Sub cboComp_Click()
txtName.SetFocus
End Sub

Private Sub cboService_Change()
 If cboService = "Internet" Then
   txtFee = frmMenu.txtNet
 Else
   txtFee = frmMenu.txtGame
 End If
End Sub

Private Sub cboService_Click()
cboService_Change
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

Sub LoadME1()
Dim i As Integer
i = cboComp.Text
With frmMenu.lv1.ListItems(i)
  .ListSubItems(1) = UCase(txtName.Text)
  .ListSubItems(2) = txtAccount
  .ListSubItems(3) = cboService
  .ListSubItems(4) = Format(dtStart, "hh:nn AMPM")
  .ListSubItems(5) = Format(dtDuration, "hh:nn AMPM")
  .ListSubItems(6) = Format(dtEnd, "hh:nn AMPM")
  .ListSubItems(7) = txtAmount
  .ListSubItems(8) = "Password"
  .ListSubItems(9) = "ID"
End With
End Sub

Private Sub cmdConfirm_Click()
LoadME
Unload Me
frmMenu.LoadME
End Sub

Public Sub LoadME()
On Error GoTo Hell
Dim rs1 As New Recordset
id = cboComp.Text
SQL = "SELECT * FROM log_temp WHERE PC='" & id & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
For i = 1 To rs.RecordCount
  SQL = "SELECT * FROM log ORDER BY id"
  rs1.Open SQL, con, 3, 3, 1
  cRt = Val(rs1.RecordCount) + 1
  AccessCode
  With rs1
    .AddNew
    !Date = Format(Date, "yyyy/mm/dd")
    !pc = id
    !Name = txtName.Text
    !account = txtAccount.Text
    !service = cboService.Text
    !Time_in = dtStart
    !time_out = dtDuration
    !duration = dtEnd
    !amount = txtAmount.Text
    '!Status = ""
    .Update
  End With
  'Next j
  With rs
    !Name = txtName.Text
    !account = txtAccount.Text
    !service = cboService.Text
    !Time_in = dtStart
    !time_out = dtDuration
    !duration = dtEnd
    !amount = txtAmount.Text
    !access = Code
    !IDCode = cRt
    .Update
  End With
Next

Exit Sub
Hell:
MsgBox Err.Description, vbInformation, "Info Message"
End Sub

Sub EnabledME(Hit As Boolean)
cboComp.Enabled = Hit
txtName.Enabled = Hit
txtAccount.Enabled = Hit
cboService.Enabled = Hit
dtStart.Enabled = Hit
dtEnd.Enabled = Hit
dtDuration.Enabled = Hit
txtAmount.Enabled = Hit
End Sub

Sub AccessCode()
Dim a, b, c, d, e, f, g, h
a = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
b = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
c = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
d = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
e = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
f = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
g = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)
h = Array(65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90)

Randomize

a = a(Int((25 * Rnd) + 1))
b = b(Int((25 * Rnd) + 1))
c = c(Int((25 * Rnd) + 1))
d = d(Int((25 * Rnd) + 1))
e = e(Int((25 * Rnd) + 1))
f = f(Int((25 * Rnd) + 1))
g = g(Int((25 * Rnd) + 1))
h = h(Int((25 * Rnd) + 1))

Code = Chr(a) & Chr(b) & Chr(c) & Chr(d) & Chr(e) & Chr(f) & Chr(g) & Chr(h)
End Sub



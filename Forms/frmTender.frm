VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTender 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   Icon            =   "frmTender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   40076
   End
   Begin VB.TextBox txtTrigger 
      Height          =   285
      Left            =   3720
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   3720
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSub 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0.00"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "&Post"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Tender"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtVat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtTender 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtDuration 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtService 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   40076
   End
   Begin MSComCtl2.DTPicker dtDuration 
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20643842
      CurrentDate     =   40076
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4560
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   3480
      Top             =   2280
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   4560
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "VAT (12%):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tendered:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computer No. 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmTender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPost_Click()
On Error GoTo Hell
SQL = "SELECT * FROM income ORDER BY id"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
If rs.RecordCount <> 0 Then
  With rs
    .AddNew
    !Date = Format(Date, "yyyy/mm/dd")
    !IDCode = txtID.Text
    !amount = txtAmount.Text
    !Change = txtChange.Text
    !Balance = txtBalance.Text
    !Vat = txtVat.Text
    !Subtotal = txtSub.Text
    !Total = txtTotal.Text
    .Update
  End With
End If

SQL = "SELECT * FROM log WHERE id='" & txtID.Text & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
If rs.RecordCount <> 0 Then
  rs!Status = "Tender"
  rs.Update
End If

id = frmMenu.txtID
SQL = "SELECT * FROM log_temp WHERE id='" & id & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1
For i = 1 To rs.RecordCount
  rs!Name = ""
  rs!account = ""
  rs!service = ""
  rs!Time_in = "00:00:00"
  rs!time_out = "00:00:00"
  rs!duration = ""
  rs!amount = "0"
  rs!access = ""
  rs!IDCode = ""
  rs.Update
  rs.MoveNext
Next
frmMenu.LoadME
Unload Me
Exit Sub
Hell:
MsgBox Err.Description, vbInformation, "Info Message"
End Sub

Private Sub Form_Activate()
If txtTender.Text = "" Then
  cmdOK.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
If cmdCancel.Caption = "&Back" Then
  cmdPost.Visible = False
  cmdCancel.Caption = "&Cancel"
  txtTender.SetFocus
  txtTender.Locked = False
Else
  Unload Me
End If
End Sub

Private Sub cmdOK_Click()
If cmdOK.Caption = "&Tender" Then
  cmdPost.Visible = True
  cmdCancel.Caption = "&Back"
  txtTender.Locked = True
End If
'-----------------------------------------------------------------------------------
Dim Total As Double
Dim Tot As Double
txtChange = 0
txtBalance = 0
txtTender = Format(txtTender, "#,##0.00")
txtAmount = Format(txtAmount, "#,##0.00")
If txtTender.Text <> "" Then
  If txtTender = txtAmount Then
    txtTotal = txtAmount
  Else
    Tot = Val(txtTender) - Val(txtAmount)
    If Tot > 0 Then
      txtChange = Tot
      'txtBalance = 0
      txtTotal = txtAmount
      txtChange.FontBold = True
      txtBalance.FontBold = False
    ElseIf Tot < 0 Then
      'txtChange = 0
      txtBalance = Tot
      txtTotal = Tot
      txtBalance.FontBold = True
      txtChange.FontBold = False
      If MsgBox("Insufficient Fund!", vbInformation, "Balance: [" & txtBalance & "]") = vbOK Then
        cmdCancel.Caption = "&Back"
        cmdCancel_Click
      End If
    End If
  End If
  txtVat = txtAmount / 1.12
  txtVat = txtAmount - txtVat
  txtVat = Format(txtVat, "#,##0.00")
  txtSub = txtTotal - txtVat
  txtChange = Format(txtChange, "#,##0.00")
  txtBalance = Format(txtBalance, "#,##0.00")
  txtSub = Format(txtSub, "#,##0.00")
  txtTotal = Format(txtTotal, "#,##0.00")
ElseIf Not IsNumeric(txtTender.Text) Then
  MsgBox "Invalid: Your amount you entered is incorrect", vbInformation, "Info Message"
  Exit Sub
Else
  MsgBox "Invalid: You must enter the amount first!", vbInformation, "Info Message"
  Exit Sub
End If
End Sub

Private Sub txtTender_Change()
If txtTender.Text <> "" Then
  cmdOK.Enabled = True
Else
  cmdOK.Enabled = False
End If
End Sub


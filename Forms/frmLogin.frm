VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Login"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -120
      ScaleHeight     =   615
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "l"
      TabIndex        =   0
      Text            =   "***"
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Access Code"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  End
End Sub

Private Sub cmdOK_Click()
On Error GoTo Hell
SQL = "SELECT * FROM login"
If rs.State = 1 Then rs.Close
rs.Open SQL, con
If rs.RecordCount <> 0 Then
  If rs!password = txtCode Then
    'frmMenu.Show
    frmFlash.Show
    Unload Me
  Else
    MsgBox "Access Denied! Please contact your system administrator...", vbInformation, "Info Message"
    Exit Sub
  End If
Else
  MsgBox "Invalid! Access Code", vbInformation, "Info Message"
  Exit Sub
End If
'con.Close
Exit Sub
Hell:
MsgBox Err.Description, vbExclamation, "Info Message"
End Sub

Private Sub Form_Activate()
'If KeyCode = 13 Then
 ' KeyCode = 9
'End If
End Sub

Private Sub Form_Load()
Connect
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
cmdOK_Click
End If
End Sub

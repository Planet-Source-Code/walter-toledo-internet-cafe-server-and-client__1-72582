VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1665
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboComp 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         ItemData        =   "frmMain.frx":08CA
         Left            =   120
         List            =   "frmMain.frx":08EC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2535
      End
      Begin VB.TextBox txtAccess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Activate"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboComp_Change()
Select Case cboComp.ListIndex
  Case 0
    Text1 = 1
  Case 1
    Text1 = 2
  Case 2
    Text1 = 3
  Case 3
    Text1 = 4
  Case 4
    Text1 = 5
  Case 5
    Text1 = 6
  Case 6
    Text1 = 7
  Case 7
    Text1 = 8
  Case 8
    Text1 = 9
  Case 9
    Text1 = 10
End Select
End Sub

Private Sub cboComp_Click()
cboComp_Change
txtAccess.SetFocus
End Sub

Private Sub Command1_Click()
id = Text1.Text
xes = txtAccess.Text
SQL = "SELECT * FROM log_temp WHERE id='" & id & "' AND access='" & xes & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con
If rs.RecordCount <> 0 Then
  MsgBox "Welcome " & UCase(rs!Name) & "! Enjoy and have fun...", vbInformation, "Info Message"
  frmTimer.lblComp = "Computer No. " & id
  
  frmTimer.lblLapse = rs!time_out
  frmTimer.Show
  Unload Me
Else
  MsgBox "Access Denied! Please check your access id and try again...", vbInformation, "Info Message"
  Exit Sub
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Keycode = 121 Then
 End
End If
End Sub

Private Sub Form_Load()
Connect
'Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.Left + 200)
'Picture1.Top = (Me.ScaleHeight / 2) - Picture1.Top
End Sub

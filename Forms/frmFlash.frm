VERSION 5.00
Begin VB.Form frmFlash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   6000
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   7000
      Left            =   5400
      Top             =   1560
   End
   Begin VB.Label lblBar 
      BackStyle       =   0  'Transparent
      Caption         =   "||||||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5970
      TabIndex        =   4
      Top             =   2175
      Width           =   255
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   2170
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ||||||||||||||||||||||||||||||||||||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Cafe'"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2450
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Cafe'"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2400
      Left            =   120
      Picture         =   "frmFlash.frx":0000
      Top             =   120
      Width           =   2325
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
frmMenu.Show
Unload Me
End Sub

Private Sub Timer2_Timer()
If Timer2.Interval Then
lblTime.Caption = Val(lblTime.Caption) + 1
If lblTime.Caption = "10" Or lblTime.Caption = "20" Or lblTime.Caption = "30" _
  Or lblTime.Caption = "40" Or lblTime.Caption = "50" Or lblTime.Caption = "60" _
  Or lblTime.Caption = "70" Or lblTime.Caption = "80" Or lblTime.Caption = "90" _
  Or lblTime.Caption = "100" Then
  lblBar.Caption = lblBar.Caption + "|||"
End If
End If
If lblTime.Caption = 100 Then
  Timer2.Enabled = False
  Timer1.Enabled = True
End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTimer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20512770
      CurrentDate     =   40078
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   240
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label lblLapse 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblComp 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer No. 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Timer2.Enabled = True
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = Me.ScaleWidth
End Sub

Private Sub Text1_Change()
DTPicker3 = Text1
If Text1 = lblLapse Then
  MsgBox "Time's Up... Game is over! Hahaha... ", vbInformation, "Info Message"
  frmMain.Show
  Unload Me
End If
End Sub

Private Sub Timer1_Timer()
lblTime = Time
End Sub

Private Sub Timer3_Timer()
Text1 = Time
End Sub

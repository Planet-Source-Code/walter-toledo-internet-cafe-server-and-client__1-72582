VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1380
      ScaleWidth      =   1365
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Develop By: Roderick Bautista, Aurelia Alvarez, Davenly Mendoza, Jessica Perez, Jovani Edu"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0 "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":06AE
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   6135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2009 - 2010 S.G. Systems"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DIT2.1c Group One - Internet Cafe' POS Systems"
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
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub


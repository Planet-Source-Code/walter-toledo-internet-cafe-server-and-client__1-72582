VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReport1 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Report"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40076
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   40076
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
x = Format(dtStart, "yyyy/mm/dd")
y = Format(dtEnd, "yyyy/mm/dd")
SQL = "SELECT * FROM log WHERE date BETWEEN '" & x & "' AND '" & y & "'"
SQL = SQL & " AND status='Tender'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con
If rs.RecordCount <> 0 Then
  Set DataReport1.DataSource = rs.DataSource
  DataReport1.Sections("Section2").Controls.Item("lbltitle").Caption = "Statement of Account"
  DataReport1.Sections("Section2").Controls.Item("lbldate").Caption = Format(dtStart, "mmmm dd, yyyy") & " to " & Format(dtEnd, "mmmm dd, yyyy")
  DataReport1.Show
  Unload Me
End If
End Sub

Private Sub Form_Load()
dtStart = Date
dtEnd = Date
End Sub

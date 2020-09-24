VERSION 5.00
Begin VB.Form frmTransfer 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   Icon            =   "frmTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Confirm"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Computer:"
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
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Computer:"
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
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If cboFrom <> "" And cboTo <> "" Then
  If Command1.Caption = "&OK" Then
    Command3.Caption = "&Back"
    Command2.Visible = True
    cboFrom.Enabled = False
    cboTo.Enabled = False
  End If
Else
  MsgBox "Operation can't continue due insufficient data!", vbInformation, "Info Message"
End If
End Sub

Private Sub Command2_Click()
On Error GoTo Hell
Dim rs1 As New Recordset
Dim i As Integer
Dim j As Integer
i = cboTo
j = cboFrom
With frmMenu
  SQL = "SELECT * FROM log_temp WHERE id='" & i & "'"
  If rs.State = 1 Then rs.Close
  rs.Open SQL, con, 3, 3, 1
  For x = 1 To rs.RecordCount
    With rs
      !pc = i
      !Name = frmMenu.lv1.ListItems(j).ListSubItems(1).Text
      !Account = frmMenu.lv1.ListItems(j).ListSubItems(2).Text
      !Service = frmMenu.lv1.ListItems(j).ListSubItems(3).Text
      !Time_in = frmMenu.lv1.ListItems(j).ListSubItems(4).Text
      !Time_out = frmMenu.lv1.ListItems(j).ListSubItems(5).Text
      !Duration = frmMenu.lv1.ListItems(j).ListSubItems(6).Text
      !Amount = frmMenu.lv1.ListItems(j).ListSubItems(7).Text
      !access = frmMenu.lv1.ListItems(j).ListSubItems(8).Text
      !IDCode = frmMenu.lv1.ListItems(j).ListSubItems(9).Text
      .Update
      .MoveNext
    End With
  Next x
  SQL = "SELECT * FROM log_temp WHERE id='" & j & "'"
  If rs1.State = 1 Then rs1.Close
  rs1.Open SQL, con, 3, 3, 1
  For x = 1 To rs1.RecordCount
    With rs1
      !pc = j
      !Name = ""
      !Account = ""
      !Service = ""
      !Time_in = "00:00:00"
      !Time_out = "00:00:00"
      !Duration = ""
      !Amount = "0"
      !access = ""
      !IDCode = ""
      .Update
      '.MoveNext
    End With
  Next x
  
  id = .lv1.ListItems(j).ListSubItems(9).Text
  SQL = "SELECT * FROM log WHERE id='" & id & "'"
  If rs.State = 1 Then rs.Close
  rs.Open SQL, con, 3, 3, 1
  For x = 1 To rs.RecordCount
    With rs
      '!id = i
      !pc = i
      '!Name = ""
      '!Account = ""
      '!Service = ""
      '!Time_in = ""
      '!Time_out = ""
      '!Durarion = ""
      '!Amount = ""
      '!Status = ""
      .Update
      .MoveNext
    End With
  Next x
End With
Unload Me
frmMenu.LoadME
Exit Sub
Hell:
MsgBox Err.Description, vbInformation, "Info Message"
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Cancel" Then
  Unload Me
Else
  Command3.Caption = "&Cancel"
  Command2.Visible = False
  cboFrom.Enabled = True
  cboTo.Enabled = True
End If
End Sub

Private Sub Form_Load()
For i = 1 To frmMenu.lv1.ListItems.Count
  If frmMenu.lv1.ListItems(i).ListSubItems(2).Text <> "" Then
    cboFrom.AddItem i
  End If
Next
For i = 1 To frmMenu.lv1.ListItems.Count
  If frmMenu.lv1.ListItems(i).ListSubItems(2).Text = "" Then
    cboTo.AddItem i
  End If
Next

End Sub

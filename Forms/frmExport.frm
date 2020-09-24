VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Export to Excel"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   40076
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   40076
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "       TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Xport()
On Error GoTo Hell
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim nRctr As Integer

Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

Set oSheet = oBook.Worksheets(1)

oSheet.Range("A1").Value = "ID"
oSheet.Range("B1").Value = "Date"
oSheet.Range("C1").Value = "PC No."
oSheet.Range("D1").Value = "Customer Name"
oSheet.Range("E1").Value = "Account Type"
oSheet.Range("F1").Value = "Service Type"
oSheet.Range("G1").Value = "Time-In"
oSheet.Range("H1").Value = "Time-Out"
oSheet.Range("I1").Value = "Duration"
oSheet.Range("J1").Value = "Amount"
oSheet.Range("K1").Value = "Status"

oSheet.Range("A1:K1").Font.Bold = True

dtStart = Format(DTPicker1, "yyyy/mm/dd")
dtEnd = Format(DTPicker2, "yyyy/mm/dd")
SQL = "SELECT * FROM log WHERE date BETWEEN '" & dtStart & "' AND '" & dtEnd & "'"
If rs.State = 1 Then rs.Close
rs.Open SQL, con, 3, 3, 1

With rs
    If .RecordCount > 0 Then
        .MoveFirst
        nRctr = 1
        Do While Not .EOF
            nRctr = nRctr + 1
            dt = Format(!Date, "mm/dd/yy")      'Date
            tn = Format(!Time_in, "hh:mm:ss AMPM")
            tf = Format(!time_out, "hh:mm:ss AMPM")
            oSheet.Range("A" & Trim(Str(nRctr))).Value = !id
            oSheet.Range("B" & Trim(Str(nRctr))).Value = dt
            oSheet.Range("C" & Trim(Str(nRctr))).Value = !pc
            oSheet.Range("D" & Trim(Str(nRctr))).Value = !Name
            oSheet.Range("E" & Trim(Str(nRctr))).Value = !account
            oSheet.Range("F" & Trim(Str(nRctr))).Value = !service
            oSheet.Range("G" & Trim(Str(nRctr))).Value = tn '!time_in
            oSheet.Range("H" & Trim(Str(nRctr))).Value = tf '!time_out
            oSheet.Range("I" & Trim(Str(nRctr))).Value = !duration
            oSheet.Range("J" & Trim(Str(nRctr))).Value = !amount
            oSheet.Range("K" & Trim(Str(nRctr))).Value = !Status
            .MoveNext
        Loop
    End If
End With

oBook.SaveAs App.Path & "\" & "Back-up_" & Format(Date, "mmddyy") & ".xls" '"C:\Export.Xls"

MsgBox "Exporting Completed.", vbInformation, "Export to Excel"

oExcel.Quit

Exit Sub
Hell:
MsgBox Err.Description, vbExclamation, "Infom Message"
End Sub

Private Sub Command1_Click()
 Xport
 Unload Me
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form28 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "Form28.frx":0000
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form28.frx":57E2
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Center"
         Object.Width           =   6703
      EndProperty
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Command1_Click()
    Form29.Show vbModal
End Sub

Private Sub Command2_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation, "Bank of Paracale": Exit Sub

With Form8
.Text1.Text = ListView1.SelectedItem
.Text2.Text = ListView1.SelectedItem
.Show vbModal
End With

End Sub

Private Sub Command3_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Bank of Paracale": Exit Sub

Dim cndelete As New ADODB.connection
Dim rsdelete As New ADODB.recordset
Dim reply As String

Call connection(cndelete, App.Path & "\db1.mdb", "rbp")
Call recordset(rsdelete, cndelete, "SELECT * FROM Table6 WHERE c_center='" & ListView1.SelectedItem & "'")

'If i

'

reply = MsgBox("Are you sure you want to delete this center name.", vbExclamation + vbYesNo, "Bank of Paracale")
If reply = vbNo Then
Exit Sub
End If

'delete record
With rsdelete
.Delete
'Call
End With

MsgBox "Record has been successfully deleted.", vbInformation, "Bank of Paracale"

Call Viewcenter

'close connection
Set cndelete = Nothing
'close recordset
Set rsdelete = Nothing

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call Viewcenter
End Sub

Public Sub Viewcenter()
Dim cncenter As New ADODB.connection
Dim rscenter As New ADODB.recordset
Dim X

Call connection(cncenter, App.Path & "\db1.mdb", "rbp")
Call recordset(rscenter, cncenter, " SELECT * FROM Table6 ORDER BY c_center ASC")

ListView1.ListItems.clear

With rscenter
    While Not .EOF
        Set X = Me.ListView1.ListItems.Add(, , .Fields!c_center)
        .MoveNext
    Wend
End With


Set cncenter = Nothing
Set rscenter = Nothing

End Sub

Private Sub Form_Load()
    Call Viewcenter
End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form26 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   Icon            =   "Form26.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
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
      MouseIcon       =   "Form26.frx":57E2
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Technical Officer"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Command1_Click()
    Form27.Show vbModal
End Sub

Private Sub Command2_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation, "Bank of Paracale": Exit Sub

    With Form7
    .Text1.Text = ListView1.SelectedItem
    .Text2.Text = ListView1.SelectedItem
    .Show vbModal
    End With
    
End Sub

Private Sub Command3_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation, "Bank of Paracale": Exit Sub

Dim cndelete As New ADODB.connection
Dim rsdelete As New ADODB.recordset
Dim reply As String

Call connection(cndelete, App.Path & "\db1.mdb", "rbp")
Call recordset(rsdelete, cndelete, "SELECT * FROM Table5 WHERE t_officer='" & ListView1.SelectedItem & "'")

'If if

reply = MsgBox("Are you sure you want to delete this record,", vbExclamation + vbYesNo, "Bank of Paracale")
If reply = vbNo Then
Exit Sub
End If


With rsdelete
'delete record
.Delete
'call del
End With

MsgBox "Record has been successfully deleted.", vbInformation, "Bank of Paracale"

Call Viewofficer

Set cndelete = Nothing
Set rsdelete = Nothing

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call Viewofficer
End Sub

Public Sub Viewofficer()

Dim cnofficer As New ADODB.connection
Dim rsofficer As New ADODB.recordset
Dim X

Call connection(cnofficer, App.Path & "\db1.mdb", "rbp")
Call recordset(rsofficer, cnofficer, "SELECT * FROM Table5 ORDER BY t_officer ASC")

ListView1.ListItems.clear

    With rsofficer
    While Not .EOF
    
        Set X = ListView1.ListItems.Add(, , .Fields!t_officer)
                .MoveNext
        
    Wend
    End With
    
Set cnofficer = Nothing
Set rsofficer = Nothing

End Sub

Private Sub Form_Load()
    Call Viewofficer
End Sub



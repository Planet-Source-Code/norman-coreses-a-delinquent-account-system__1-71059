VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   4140
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
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
      MouseIcon       =   "Form15.frx":57E2
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Center"
         Object.Width           =   6703
      EndProperty
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Command1_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation, "Bank of Paracale": Exit Sub

    With Form14
    .Text1.Text = ListView1.SelectedItem
    End With

    Unload Me

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Viewcenter
End Sub


Public Sub Viewcenter()
Dim cnviewcenter As New ADODB.connection
Dim rsviewcenter As New ADODB.recordset
Dim X

Call connection(cnviewcenter, App.Path & "\db1.mdb", "rbp")
Call recordset(rsviewcenter, cnviewcenter, "SELECT * FROM Table6 ORDER BY c_center ASC")

ListView1.ListItems.clear
            
    With rsviewcenter
        While Not .EOF
        
            Set X = ListView1.ListItems.Add(, , .Fields!c_center)
            .MoveNext
        Wend
    End With

Set cnviewcenter = Nothing
Set rsviewcenter = Nothing
End Sub

Private Sub ListView1_DblClick()
    Call Command1_Click
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    Call Command1_Click
End Sub

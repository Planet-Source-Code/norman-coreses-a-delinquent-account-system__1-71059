VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   2295
   ClientTop       =   2745
   ClientWidth     =   11400
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11400
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11415
      TabIndex        =   11
      Top             =   0
      Width           =   11415
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&View Center"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   11355
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   4080
         Picture         =   "Form14.frx":57E2
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "..."
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&CLOSE"
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
      Left            =   6840
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&PRINT"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3960
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6985
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TECHNICAL OFFICER"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "TYPE"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CLIENT NAME"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "LOAN AMOUNT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "DATE GRANTED"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "MATURITY DATE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "TERM OF LOAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "PRINCIPAL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "INTEREST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "PENALTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "BALANCE"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Borrower:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   10
      Top             =   5760
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   1920
      TabIndex        =   9
      Top             =   5760
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Total Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8400
      TabIndex        =   5
      Top             =   5760
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   9720
      TabIndex        =   4
      Top             =   5760
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Enter Center"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   1245
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741





Private Sub Command1_Click()
Dim cnviewcenter As New ADODB.connection
Dim rsviewcenter As New ADODB.recordset
Dim X, tot

Text1.Text = Trim(Text1.Text)
If Trim(Text1.Text) = "" Then
Exit Sub
End If

Call connection(cnviewcenter, App.Path & "\db1.mdb", "rbp")
Call recordset(rsviewcenter, cnviewcenter, "SELECT * FROM Table1 WHERE n_center='" & Text1.Text & "'")

If rsviewcenter.RecordCount = 0 Then
    MsgBox "The record you requested could not be found.", vbExclamation, "Bank of Paracale"
    Text1.Text = ""
    Label2.Caption = ""
    Label4.Caption = ""
    Text1.SetFocus
    ListView1.ListItems.clear
    Exit Sub
    End If
        
        ListView1.ListItems.clear
                                      
         With rsviewcenter
            While Not .EOF
                
                Set X = ListView1.ListItems.Add(, , .Fields!t_officer)
                    X.SubItems(1) = .Fields!Type
                    X.SubItems(2) = .Fields!c_name
                    X.SubItems(3) = .Fields!l_amount
                    X.SubItems(4) = .Fields!d_granted
                    X.SubItems(5) = .Fields!m_date
                    X.SubItems(6) = .Fields!t_loan
                    X.SubItems(7) = .Fields!d_principal
                    X.SubItems(8) = .Fields!d_interest
                    X.SubItems(9) = .Fields!d_penalty
                    X.SubItems(10) = .Fields!t_balance
                    tot = tot + Val(X.SubItems(10))
                    .MoveNext
            Wend
         End With
          
          Label4.Caption = ListView1.ListItems.Count
          Label2.Caption = tot
          
Set cnviewcenter = Nothing
Set rsviewcenter = Nothing
          
End Sub

Private Sub Command2_Click()
Dim cnprintcenter As New ADODB.connection
Dim rsprintcenter As New ADODB.recordset

Call connection(cnprintcenter, App.Path & "\db1.mdb", "rbp")
Call recordset(rsprintcenter, cnprintcenter, "SELECT * FROM Table1 WHERE n_center='" & Text1.Text & "'")
       
Set RPT3.DataSource = rsprintcenter
           
Unload Me

RPT3.Show

Set cnprintcenter = Nothing
Set rsprintcenter = Nothing

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Form15.Show vbModal
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call positionform(Form14)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Form15.Show vbModal
End Sub

Private Sub Label2_Change()
    Label2.Caption = Format(Label2, "#,##0.00")
End Sub

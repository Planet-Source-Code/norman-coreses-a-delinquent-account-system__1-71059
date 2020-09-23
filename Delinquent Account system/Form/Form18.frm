VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6285
   ClientLeft      =   3795
   ClientTop       =   3120
   ClientWidth     =   11400
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6254.04
   ScaleMode       =   0  'User
   ScaleWidth      =   11400
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   11415
      TabIndex        =   9
      Top             =   0
      Width           =   11415
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   4200
         Picture         =   "Form18.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&View by Type"
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
         Left            =   4905
         TabIndex        =   10
         Top             =   240
         Width           =   1665
      End
   End
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
      Left            =   6840
      TabIndex        =   3
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7011
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
         Text            =   "CENTER"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "TECHNICAL OFFICER"
         Object.Width           =   4410
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
   Begin VB.Label Label6 
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
      TabIndex        =   8
      Top             =   5760
      Width           =   60
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
      TabIndex        =   7
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
      Left            =   9720
      TabIndex        =   6
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
      Caption         =   "&Type:"
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
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   540
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Combo1_Click()
Dim cndisplaycombo As New ADODB.connection
Dim rsdisplaycombo As New ADODB.recordset
Dim cn1displaycombo As New ADODB.connection
Dim rs1displaycombo As New ADODB.recordset

Dim tot As Double
Dim X

If Combo1.Text = "MFWO" Then

Call connection(cndisplaycombo, App.Path & "\db1.mdb", "rbp")
Call recordset(rsdisplaycombo, cndisplaycombo, "SELECT * FROM Table1 ORDER BY type ASC")

ListView1.ListItems.clear

    With rsdisplaycombo
        While Not .EOF
            
            Set X = ListView1.ListItems.Add(, , .Fields!n_center)
                X.SubItems(1) = .Fields!t_officer
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
            
                Label4.Caption = tot
                Label6.Caption = ListView1.ListItems.Count
Else

Call connection(cn1displaycombo, App.Path & "\db1.mdb", "rbp")
Call recordset(rs1displaycombo, cn1displaycombo, "SELECT * FROM Table1 WHERE type='" & Combo1.Text & "'")

ListView1.ListItems.clear

    With rs1displaycombo
        While Not .EOF
        
            Set X = ListView1.ListItems.Add(, , .Fields!n_center)
                X.SubItems(1) = .Fields!t_officer
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

End If
            
            Label4.Caption = tot
            Label6.Caption = ListView1.ListItems.Count

Set cndisplaycombo = Nothing
Set rsdisplaycombo = Nothing
Set cn1displaycombo = Nothing
Set rs1displaycombo = Nothing

End Sub

Private Sub Command1_Click()
Dim cnprintmfund As New ADODB.connection
Dim rsprintmfund As New ADODB.recordset
Dim rs1printmfund As New ADODB.recordset

Call connection(cnprintmfund, App.Path & "\db1.mdb", "rbp")
Call recordset(rs1printmfund, cnprintmfund, "SELECT * FROM Table1 ORDER BY type ASC")
Call recordset(rsprintmfund, cnprintmfund, "SELECT * FROM Table1 WHERE type='" & Combo1.Text & "'")


If Combo1.Text = "MFWO" Then

    Set RPT4.DataSource = rs1printmfund

Else

    Set RPT4.DataSource = rsprintmfund
    
End If

Unload Me

RPT4.Show

Set cnprintmfund = Nothing
Set rsprintmfund = Nothing
Set rs1printmfund = Nothing

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Public Sub Viewtype()

Dim cnviewcombo As New ADODB.connection
Dim rsviewcombo As New ADODB.recordset

Call connection(cnviewcombo, App.Path & "\db1.mdb", "rbp")
Call recordset(rsviewcombo, cnviewcombo, "SELECT * FROM Table4")

ListView1.ListItems.clear

    With rsviewcombo
        While Not .EOF
        
            Combo1.AddItem .Fields!category
            .MoveNext
        Wend
    End With
    
    
Set cnviewcombo = Nothing
Set rsviewcombo = Nothing

End Sub

Private Sub Form_Load()
    Call positionform(Form18)
    Call Viewtype
End Sub

Private Sub Label4_Change()
Label4.Caption = Format(Label4, "#,##0.00")
End Sub


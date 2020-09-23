VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   10740
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10815
      TabIndex        =   36
      Top             =   0
      Width           =   10815
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   3000
         Picture         =   "Form12.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&View Individual Record"
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
         Left            =   3840
         TabIndex        =   37
         Top             =   240
         Width           =   2835
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
      Left            =   9120
      TabIndex        =   34
      Top             =   6240
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
      Left            =   7920
      TabIndex        =   33
      Top             =   6240
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   6000
      TabIndex        =   32
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form12.frx":66AC
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date of O.R"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Payment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "OR Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   6000
      ScaleHeight     =   5175
      ScaleWidth      =   4575
      TabIndex        =   27
      Top             =   840
      Width           =   4575
      Begin VB.TextBox Text12 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox Text13 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Payment:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   4200
         Width           =   930
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4200
         Width           =   3855
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3840
         Width           =   3855
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3480
         Width           =   3855
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Penalty:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   3840
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Interest:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Principal:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Term of Loan:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Maturity Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Date  Granted:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Loan Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Center:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Technical Officer:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label14 
         Caption         =   "&Account Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5775
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
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lbltot 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Client Name:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1785
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Combo1_Change()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
ListView1.ListItems.clear

End Sub

Private Sub Combo1_Click()
Dim cnviewclient As New ADODB.connection
Dim rsviewclient As New ADODB.recordset

Call connection(cnviewclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsviewclient, cnviewclient, "SELECT * FROM Table1 WHERE c_name='" & Combo1.Text & "'")
                                       
If rsviewclient.RecordCount = 0 Then
MsgBox "The record you requested could not be found", vbExclamation, "Bank of Paracale"
Exit Sub
End If
                         
With rsviewclient
    Text1.Text = .Fields!t_officer
    Text2.Text = .Fields!n_center
    Text3.Text = .Fields!Type
    Text4.Text = .Fields!l_amount
    Text5.Text = .Fields!d_granted
    Text6.Text = .Fields!m_date
    Text7.Text = .Fields!t_loan
    Text8.Text = .Fields!d_principal
    Text9.Text = .Fields!d_interest
    Text10.Text = .Fields!d_penalty
    Text14.Text = .Fields!contrac_n
End With

Call fill

Set cnviewclient = Nothing
Set rsviewclient = Nothing
    
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
                            
If KeyAscii = 13 Then
    Call Combo1_Click
End If
                        
End Sub

Private Sub Command1_Click()
Dim cnprintclient As New ADODB.connection
Dim rsprintclient As New ADODB.recordset
                        
Call connection(cnprintclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsprintclient, cnprintclient, "SELECT * FROM Table3 WHERE c_name='" & Combo1.Text & "'")
                                                                         
Set RPT1.DataSource = rsprintclient
                                     
Unload Me
                
RPT1.Show
                                    
Set cnprintclient = Nothing
Set rsprintclient = Nothing
                                    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Combo1.SetFocus
End Sub

Private Sub Form_Load()

Call positionform(Form12)

Call loadclient

End Sub

Public Sub loadclient()
Dim cnloadclient As New ADODB.connection
Dim rsloadclient As New ADODB.recordset

Call connection(cnloadclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsloadclient, cnloadclient, "SELECT * FROM Table1 ORDER BY c_name ASC")

With rsloadclient
    While Not .EOF
        Combo1.AddItem .Fields!c_name
        .MoveNext
    Wend
    End With
      
Set cnloadclient = Nothing
Set rsloadclient = Nothing
End Sub

Public Sub fill()
Dim cnviewtransac As New ADODB.connection
Dim rsviewtransac As New ADODB.recordset
Dim tot, X
                                  
Call connection(cnviewtransac, App.Path & "\db1.mdb", "rbp")
Call recordset(rsviewtransac, cnviewtransac, "SELECT * FROM Table3 WHERE contrac_n='" & Text14.Text & "'")
                                                                           

                                                                           
ListView1.ListItems.clear
                                       
    With rsviewtransac
        While Not .EOF
            Set X = ListView1.ListItems.Add(, , .Fields!d_or)
                X.SubItems(1) = .Fields!d_payment
                X.SubItems(2) = .Fields!or_number
                                                                
                tot = tot + Val(X.SubItems(1))
                .MoveNext
        Wend
    End With
                 
    Text12.Text = tot
    
Set cnviewtransac = Nothing
Set rsviewtransac = Nothing
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count < 1 Then: MsgBox "No record to search.", vbExclamation, "Bank of Paracale": Exit Sub

With Form13
    .Text1.Text = Combo1.Text
    .Text2.Text = ListView1.SelectedItem.SubItems(1)
    .Text3.Text = ListView1.SelectedItem.SubItems(2)
    .Show vbModal
End With

End Sub

Private Sub ListView1_GotFocus()
    Call fill
End Sub

Private Sub Text10_Change()
    Call Fillcompute
End Sub

Private Sub Text11_Change()
    Call compute
End Sub

Private Sub Text12_Change()
    Call compute
End Sub

Private Sub Text13_Change()
    Call updatepayment
End Sub

Public Sub compute()
    Text13.Text = Val(Text11.Text) - Val(Text12.Text)
End Sub
Public Sub Fillcompute()
    Text11.Text = Val(Text8.Text) + Val(Text9.Text) + Val(Text10.Text)
End Sub

Private Sub Text8_Change()
    Call Fillcompute
End Sub

Private Sub Text9_Change()
    Call Fillcompute
End Sub

Public Sub updatepayment()
Dim cnupdatepayment As New ADODB.connection
Dim rsupdatepayment As New ADODB.recordset

Call connection(cnupdatepayment, App.Path & "\db1.mdb", "rbp")
Call recordset(rsupdatepayment, cnupdatepayment, "SELECT * FROM Table1 WHERE c_name='" & Combo1.Text & "'")

If rsupdatepayment.RecordCount = 0 Then
Exit Sub
End If
                                   
    With rsupdatepayment
        .Fields!t_balance = Text13.Text
        .Update
        .Requery
    End With
       
Set cnupdatepayment = Nothing
Set rsupdatepayment = Nothing
   
End Sub





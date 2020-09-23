VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   6090
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -840
      ScaleHeight     =   735
      ScaleWidth      =   7455
      TabIndex        =   35
      Top             =   0
      Width           =   7455
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Existing Record"
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
         Left            =   945
         TabIndex        =   36
         Top             =   240
         Width           =   5955
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   0
         Left            =   1800
         Picture         =   "Form4.frx":57E2
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
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
      Left            =   3360
      TabIndex        =   31
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4560
      TabIndex        =   30
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   5520
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1800
         TabIndex        =   29
         Top             =   2760
         Width           =   1575
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
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text9 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   13
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text8 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox Text7 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox Text6 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox Text5 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox Text4 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text2 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   3735
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox Text10 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text11 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   33023
         CalendarTrailingForeColor=   16777215
         Format          =   3866625
         CurrentDate     =   39500
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   2040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   33023
         CalendarTrailingForeColor=   16777215
         Format          =   3866625
         CurrentDate     =   39500
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Client Name:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Type"
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
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label11 
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
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Penalty:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Interest:"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Principal:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Term of Loan:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Maturity Date:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Date Granted:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Loan Amount:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Center:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Technical Officer:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5895
      Begin VB.ComboBox Combo3 
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
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Client No:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Add new Client:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   34
      Top             =   240
      Width           =   6165
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   4
      Left            =   240
      Picture         =   "Form4.frx":66AC
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741

Dim sString As String
Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_Change()
    Call clear
End Sub

Private Sub Combo3_Click()
Dim cneditclient As New ADODB.connection
Dim rseditclient As New ADODB.recordset
                                                   
Call connection(cneditclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rseditclient, cneditclient, "SELECT * FROM Table1 WHERE contrac_n='" & Combo3.Text & "'")

If rseditclient.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Bank of Paracale"

Exit Sub
End If

With rseditclient
        
            Text1.Text = .Fields("t_officer")
            Text2.Text = .Fields("n_center")
            Text9.Text = .Fields("c_name")
            Text4.Text = .Fields("l_amount")
            Text10.Text = .Fields("d_granted")
            Text11.Text = .Fields("m_date")
            Combo1.Text = .Fields("type")
            Combo2.Text = .Fields("t_loan")
            Text5.Text = .Fields("d_principal")
            Text6.Text = .Fields("d_interest")
            Text7.Text = .Fields("d_penalty")
            Text8.Text = .Fields("t_balance")
           
End With
    
               
'close connection
Set cneditclient = Nothing
'close recordset
Set rseditclient = Nothing

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Combo3_Click
        
    End If
End Sub

Private Sub Command1_Click()
Dim cnupdateclient As New ADODB.connection
Dim rsupdateclient As New ADODB.recordset

If sempty(Combo3) = True Then Exit Sub
If sempty(Text9) = True Then Exit Sub
If sempty(Combo1) = True Then Exit Sub
If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text4) = True Then Exit Sub
If sempty(Text10) = True Then Exit Sub
If sempty(Text11) = True Then Exit Sub
If sempty(Combo2) = True Then Exit Sub
If sempty(Text5) = True Then Exit Sub
If sempty(Text6) = True Then Exit Sub
If sempty(Text7) = True Then Exit Sub

If snumber(Text4) = True Then Text4.SetFocus: Call hlfocus(Text4): Exit Sub
If snumber(Text5) = True Then Text5.SetFocus: Call hlfocus(Text5): Exit Sub
If snumber(Text6) = True Then Text6.SetFocus: Call hlfocus(Text6): Exit Sub
If snumber(Text7) = True Then Text7.SetFocus: Call hlfocus(Text7): Exit Sub
                    
Call connection(cnupdateclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsupdateclient, cnupdateclient, "SELECT * FROM Table1 WHERE contrac_n='" & Combo3.Text & "'")

If rsupdateclient.RecordCount = 0 Then
MsgBox "No Record found. Please check it.", vbExclamation, "Bank of Paracale"
Combo3.SetFocus
Exit Sub
End If
 
sString = rsupdateclient.Fields!c_name


 
With rsupdateclient
    .Fields!Type = Combo1.Text
    .Fields!c_name = Text9.Text
    .Fields!t_officer = Text1.Text
    .Fields!n_center = Text2.Text
    .Fields!l_amount = Text4.Text
    .Fields!d_granted = Text10.Text
    .Fields!m_date = Text11.Text
    .Fields!t_loan = Combo2.Text
    .Fields!d_principal = Text5.Text
    .Fields!d_interest = Text6.Text
    .Fields!d_penalty = Text7.Text
    .Fields!t_balance = Text8.Text
    .Update
    .Requery
End With

    MsgBox "Record successfully updated.", vbInformation, "Bank of Paracale"
    
    Unload Me
    
Set cnupdateclient = Nothing
Set rsupdateclient = Nothing

End Sub

Private Sub Command10_Click()
    Form6.Show vbModal
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Form6.Show vbModal
End Sub

Private Sub Command4_Click()
    Form5.Show vbModal
End Sub

Private Sub DTPicker1_Change()
    Text10.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
    Text10.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
    Text11.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_Click()
    Text11.Text = DTPicker2.Value
End Sub

Private Sub Form_Load()

Call positionform(Form4)
                  
Call viewcontract
                    
With Combo1
.AddItem "MF"
.AddItem "WO"
End With

With Combo2
.AddItem "25 Weeks"
.AddItem "50 Weeks"
.AddItem "75 Weeks"
End With

DTPicker1.Value = Date
DTPicker2.Value = Date

End Sub

Public Sub compute()
Dim totalbalance As String

totalbalance = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
Text8.Text = totalbalance
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Form5.Show vbModal
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Form6.Show vbModal
End Sub

Private Sub Text4_LostFocus()
    Text4.Text = Format(Text4, "##,#0.00")
End Sub

Private Sub Text5_Change()
    Call compute
End Sub

Private Sub Text6_Change()
    Call compute
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If sempty(Text6) = True Then Exit Sub
Text7.SetFocus
End If
End Sub

Private Sub Text7_Change()
    Call compute
End Sub

Sub clear()

Combo1.Text = ""
Text9.Text = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text10.Text = ""
Text11.Text = ""
Combo2.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click
End If
End Sub

Public Sub viewcontract()
Dim cnloadclient As New ADODB.connection
Dim rsloadclient As New ADODB.recordset

Call connection(cnloadclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsloadclient, cnloadclient, "SELECT * FROM Table1 ORDER BY contrac_n ASC")

With rsloadclient
    While Not .EOF
    Combo3.AddItem .Fields!contrac_n
    .MoveNext
    Wend
End With

Set cnloadclient = Nothing
Set rsloadclient = Nothing
End Sub

Private Sub Text9_LostFocus()
Text9.Text = StrConv(Text9, vbProperCase)
End Sub

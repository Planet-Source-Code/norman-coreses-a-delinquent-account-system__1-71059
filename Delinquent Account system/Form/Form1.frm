VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Bookman Old Style"
      Size            =   8.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   6240
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -480
      ScaleHeight     =   735
      ScaleWidth      =   7455
      TabIndex        =   31
      Top             =   0
      Width           =   7455
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   1800
         Picture         =   "Form1.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Add New Client"
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
         Left            =   480
         TabIndex        =   32
         Top             =   240
         Width           =   6225
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   17
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   29
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Top             =   4680
         Width           =   2535
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
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1080
         Width           =   3735
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
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   23
         Top             =   1800
         Width           =   3735
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
         TabIndex        =   22
         Top             =   2160
         Width           =   2775
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
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3240
         Width           =   1575
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
         TabIndex        =   20
         Top             =   3600
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
         TabIndex        =   19
         Top             =   3960
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
         TabIndex        =   18
         Top             =   4320
         Width           =   2535
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3120
         TabIndex        =   26
         Top             =   2880
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
         Format          =   56360961
         CurrentDate     =   39500
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   2520
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
         Format          =   56360961
         CurrentDate     =   39500
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Account Number:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1680
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   720
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   4680
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   4320
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   10
         Top             =   3960
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   3600
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   3240
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   2880
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   2520
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1140
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   3
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741

Dim num, ans As Integer
        
Private Sub Combo1_KeyPress(KeyAscii As Integer)
                
    If sempty(Combo1) = True Then Exit Sub
    Text1.SetFocus
                    
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
                                                                            
    If sempty(Combo2) = True Then Exit Sub
    Text5.SetFocus
                                                
End Sub

Private Sub Command1_Click()
Dim cnaddnew As New ADODB.connection
Dim rsaddnew As New ADODB.recordset

If sempty(Combo1) = True Then Exit Sub
If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If sempty(Text4) = True Then Exit Sub
If sempty(Text9) = True Then Exit Sub
If sempty(Text10) = True Then Exit Sub
If sempty(Combo2) = True Then Exit Sub
If sempty(Text5) = True Then Exit Sub
If sempty(Text6) = True Then Exit Sub
If sempty(Text7) = True Then Exit Sub

If snumber(Text4.Text) = True Then Text4.SetFocus: Call hlfocus(Text4): Exit Sub
If snumber(Text5.Text) = True Then Text5.SetFocus: Call hlfocus(Text5): Exit Sub
If snumber(Text6.Text) = True Then Text6.SetFocus: Call hlfocus(Text6): Exit Sub
If snumber(Text7.Text) = True Then Text7.SetFocus: Call hlfocus(Text7): Exit Sub

Call connection(cnaddnew, App.Path & "\db1.mdb", "rbp")
Call recordset(rsaddnew, cnaddnew, "SELECT  * FROM Table1")

'If i

ans = MsgBox("Are you sure with the information listed?", vbInformation + vbYesNo, "Bank of Paracale")
If ans = vbNo Then
Exit Sub
End If
     
With rsaddnew
    .AddNew
    .Fields!contrac_n = Label12.Caption
    .Fields!t_officer = Text1.Text
    .Fields!n_center = Text2.Text
    .Fields!Type = Combo1.Text
    .Fields!c_name = Text3.Text
    .Fields!l_amount = Text4.Text
    .Fields!d_granted = Text9.Text
    .Fields!m_date = Text10.Text
    .Fields!t_loan = Combo2.Text
    .Fields!d_principal = Text5.Text
    .Fields!d_interest = Text6.Text
    .Fields!d_penalty = Text7.Text
    .Fields!t_balance = Text8.Text
    .Update
    .Requery
End With

Call counter

Call clear

Combo1.SetFocus
 
MsgBox "New record record has been saved", vbInformation, "Bank of Paracale"

Set cnaddnew = Nothing
Set rsaddnew = Nothing

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Form2.Show vbModal
End Sub

Private Sub Command4_Click()
    Form3.Show vbModal
End Sub

Private Sub DTPicker1_Change()
Text9.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
Text9.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text10.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_Click()
Text10.Text = DTPicker2.Value
End Sub

Private Sub Form_Load()

Call counter

Call positionform(Form1)

Combo1.clear
Combo2.clear

DTPicker1.Value = Date
DTPicker2.Value = Date
Text9.Text = Date
Text10.Text = Date
Combo1.clear

With Combo1
    .AddItem "MF"
    .AddItem "WO"
End With

With Combo2
    .AddItem "25 weeks"
    .AddItem "50 weeks"
    .AddItem "75 weeks"
End With

End Sub

Public Sub compute()
    Text8.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Form2.Show vbModal
    If sempty(Text1) = True Then Exit Sub
    Text2.SetFocus
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sempty(Text10) = True Then Exit Sub
    Combo2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Form3.Show vbModal
    If sempty(Text2) = True Then Exit Sub
    Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sempty(Text3) = True Then Exit Sub
    Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
    Text3.Text = StrConv(Text3, vbProperCase)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sempty(Text4) = True Then Exit Sub
    Text9.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
    Text4.Text = Format(Text4, "##,#0.00")
End Sub

Private Sub Text5_Change()
    Call compute
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If sempty(Text5) = True Then Exit Sub
        Text6.SetFocus
    End If
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
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.ListIndex = -1
Combo2.ListIndex = -1
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call Command1_Click
End If
End Sub



Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If sempty(Text9) = True Then Exit Sub
    Text10.SetFocus
End If
End Sub

Public Sub counter()
Dim cncounter As New ADODB.connection
Dim rscounter As New ADODB.recordset

num = 999

Call connection(cncounter, App.Path & "\db1.mdb", "rbp")
Call recordset(rscounter, cncounter, "SELECT * FROM Table1")

With rscounter
If .RecordCount = 0 Then
Label12.Caption = 1000

Else

num = num + .RecordCount + 1
Label12.Caption = num

End If
End With

Set cncounter = Nothing
Set rscounter = Nothing
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   5220
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -120
      ScaleHeight     =   735
      ScaleWidth      =   5415
      TabIndex        =   20
      Top             =   0
      Width           =   5415
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Payment Transaction"
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
         TabIndex        =   21
         Top             =   240
         Width           =   5205
      End
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   960
         Picture         =   "Form11.frx":57E2
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Saved"
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
      Left            =   2400
      TabIndex        =   19
      Top             =   4560
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
      Left            =   3600
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4935
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   2655
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2160
         Width           =   2655
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1800
         Width           =   2655
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
         Format          =   56426497
         CurrentDate     =   39507
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1920
         TabIndex        =   23
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         TabIndex        =   22
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
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
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Payment:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Date of O.R:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "O.R Number:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Loan Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Center: "
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
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acount Number:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4935
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
         TabIndex        =   1
         Top             =   240
         Width           =   2880
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter Client Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741




Dim ans As Integer

Private Sub Combo1_Change()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text7.Text = ""
Label5.Caption = ""
Label7.Caption = ""

End Sub

Private Sub Combo1_Click()
Dim cnloadcombo As New ADODB.connection
Dim rsloadcombo As New ADODB.recordset

Call connection(cnloadcombo, App.Path & "\db1.mdb", "rbp")
Call recordset(rsloadcombo, cnloadcombo, "SELECT * FROM Table1 WHERE c_name='" & Combo1.Text & "'")

If rsloadcombo.RecordCount = 0 Then
MsgBox "The record you requested could not be found", vbExclamation, "Bank of Paracale"
Exit Sub
End If

With rsloadcombo
    
    Text1.Text = .Fields!t_officer
    Text2.Text = .Fields!n_center
    Text3.Text = .Fields!l_amount
    Label5.Caption = .Fields!Type
    Label7.Caption = .Fields!contrac_n
    
End With
                                  
Set cnloadcombo = Nothing
Set rsloadcombo = Nothing
                             
End Sub
                       
                                                           
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Combo1_Click
End If
End Sub

Private Sub Command1_Click()
Dim cntransacclient As New ADODB.connection
Dim rstransacclient As New ADODB.recordset

If Label5.Caption = "" And Label7.Caption = "" And Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
MsgBox "No record found. please check it.", vbExclamation, "Bank of Paracale"
Combo1.SetFocus
Exit Sub
End If

If sempty(Text5) = True Then Exit Sub
If sempty(Text7) = True Then Exit Sub
If snumber(Text5) = True Then Text5.SetFocus: Call hlfocus(Text5): Exit Sub
If snumber(Text7) = True Then Text7.SetFocus: Call hlfocus(Text7): Exit Sub

Call connection(cntransacclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rstransacclient, cntransacclient, "SELECT * FROM Table3")

'If if

With rstransacclient
    .AddNew
    .Fields("c_name") = Combo1.Text
    .Fields("t_officer") = Text1.Text
    .Fields("n_center") = Text2.Text
    .Fields("l_amount") = Text3.Text
    .Fields("type") = Label5.Caption
    .Fields("contrac_n") = Label7.Caption
    .Fields("d_payment") = Text5.Text
    .Fields("or_number") = Text7.Text
    .Fields("d_or") = DTPicker1.Value
    .Update
 
 End With
 
Call compute

ans = MsgBox("Successfuly saved. transaction another.", vbInformation + vbYesNo, "Bank of Paracale")
If ans = vbYes Then
Call clear
Else
Unload Me
End If

Set cntransacclient = Nothing
Set rstransacclient = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
                                
Call positionform(Form11)
                                                    
Call Viewcombo
                                                            
DTPicker1.Value = Date
                                                
End Sub

Public Sub compute()

Dim cnupdatepayment As New ADODB.connection
Dim rsupdatepayment As New ADODB.recordset

Call connection(cnupdatepayment, App.Path & "\db1.mdb", "rbp")
Call recordset(rsupdatepayment, cnupdatepayment, "SELECT * FROM Table1 WHERE c_name='" & Combo1.Text & "'")

    With rsupdatepayment
        .Fields!t_balance = .Fields!t_balance - Val(Text5.Text)
        .Update
    End With
    
Set cnupdatepayment = Nothing
Set rsupdatepayment = Nothing

End Sub

Public Sub Viewcombo()
Dim cnsearchcombo As New ADODB.connection
Dim rssearchcombo As New ADODB.recordset

Call connection(cnsearchcombo, App.Path & "\db1.mdb", "rbp")
Call recordset(rssearchcombo, cnsearchcombo, "SELECT * FROM Table1 ORDER BY c_name ASC")

Combo1.clear

    With rssearchcombo
        While Not .EOF
            Combo1.AddItem .Fields!c_name
            .MoveNext
        Wend
    End With
    
Set cnsearchcombo = Nothing
Set rssearchcombo = Nothing

End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
Text7.Text = ""
Label5.Caption = ""
Label7.Caption = ""
Combo1.Text = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
End If
End Sub

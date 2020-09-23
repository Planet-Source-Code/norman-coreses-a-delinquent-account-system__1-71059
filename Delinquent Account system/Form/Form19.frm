VERSION 5.00
Begin VB.Form Form19 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "Form19.frx":0000
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   1200
      Top             =   6360
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
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image imgclose 
      Height          =   480
      Left            =   480
      Picture         =   "Form19.frx":57E2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
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
      TabIndex        =   2
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
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
      TabIndex        =   1
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741




Dim login As Boolean

Private Sub Command1_Click()
Dim cnlogin As New ADODB.connection
Dim rslogin As New ADODB.recordset

Call connection(cnlogin, App.Path & "\db1.mdb", "rbp")
Call recordset(rslogin, cnlogin, "SELECT * FROM UserAccount")

If Text1.Text = "" Then Text1.SetFocus: Exit Sub
If Text2.Text = "" Then Text2.SetFocus: Exit Sub

If recfound(rslogin, "Username", Text1.Text) = False Then

    MsgBox "User Name doesn't Exist", vbExclamation, "Bank of Paracale"
    terminate
    Text1.SetFocus
    Call hlfocus(Text1)
        
    Else

    If Password = Text2.Text Then
    Me.Hide
    MDIForm1.Show
        
    Else

    MsgBox "You did not enter the correct password....", vbExclamation, "Bank of Paracale"
    terminate
    Text2.SetFocus
    Call hlfocus(Text2)
        
    End If
    End If
    
    login = True
    
Set cnlogin = Nothing
Set rslogin = Nothing

End Sub

Private Sub Command2_Click()
login = False
Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then MsgBox "The application is already running.", vbInformation, "Bank of Paracale": Unload Me: Exit Sub



'Call positionform(Form19)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1, vbProperCase)
End Sub
Public Sub terminate()
'Static attempt As Integer

'attempt = attempt + 1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
'If imgclose.Enable = True Then
'imgopen.Visible = True
'ElseIf imgopen.Enable = True Then
'imgclose.Visible = True
'End If
End Sub


VERSION 5.00
Begin VB.Form Form21 
   BorderStyle     =   0  'None
   Caption         =   "Form21"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   -17040
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Password:"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   -13320
      TabIndex        =   3
      Top             =   960
      Width           =   45
   End
   Begin VB.Image imgkey 
      Height          =   315
      Left            =   4080
      Picture         =   "Form21.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   330
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text <> Password Then
        MsgBox "INVALID PASSWORD. Please type the correct password.", vbCritical, "Bank of Paracale"
        Call hlfocus(Text1)
    Else
        Unload Me
    End If
End If
    
End Sub

Private Sub Timer1_Timer()
 If Label2.Visible = True Then
        'imgopen.Enable = False
        imgkey.Left = 3360
    ElseIf Label2.Visible = True Then
        'imgclose.Enable = True
        imgkey.Left = 4080
    End If
End Sub

VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&Copyright @ 2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&Bank of Paracale"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2145
         TabIndex        =   2
         Top             =   720
         Width           =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Delinquent Account's version 1.0"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   1320
         Width           =   4365
      End
      Begin VB.Image Image1 
         Height          =   885
         Left            =   240
         Picture         =   "Form10.frx":0000
         Top             =   600
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
'Call center
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub



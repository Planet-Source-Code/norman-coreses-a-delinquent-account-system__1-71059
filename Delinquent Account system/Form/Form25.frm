VERSION 5.00
Begin VB.Form Form25 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   Icon            =   "Form25.frx":0000
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -2160
      ScaleHeight     =   615
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   2280
         Picture         =   "Form25.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Log in Information"
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
         Left            =   2760
         TabIndex        =   1
         Top             =   120
         Width           =   2385
      End
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Command1_Click()
Dim cnviewuser As New ADODB.connection
Dim rsviewuser As New ADODB.recordset

Dim ans As String



'Set RPT5.DataSource = rsviewuser

'Unload Me

'Set rsviewuser = Nothing

End Sub



'End Sub

Private Sub DTPicker1_Click()
'Text1.Text =
End Sub

Private Sub Form_Load()
Date
End Sub



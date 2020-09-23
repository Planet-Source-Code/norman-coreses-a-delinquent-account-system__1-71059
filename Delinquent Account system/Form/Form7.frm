VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Technical Officer"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
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
         Left            =   3000
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&New T.O :"
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
         TabIndex        =   6
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Old T.O :"
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
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Command1_Click()
Dim cnedit As New ADODB.connection
Dim rsedit As New ADODB.recordset

If sempty(Text2) = True Then Exit Sub

Call connection(cnedit, App.Path & "\db1.mdb", "rbp")
Call recordset(rsedit, cnedit, "SELECT * FROM Table5 WHERE t_officer='" & Text1.Text & "'")

'If Text
'
'

'

With rsedit
.Fields!t_officer = Text2.Text
.Update
End With

MsgBox "Technical Officer successfully updated.", vbInformation, "Bank of Paracale"

Unload Me

Set cnedit = Nothing
Set rsedit = Nothing

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text2_LostFocus()
Text2.Text = StrConv(Text2, vbProperCase)
End Sub



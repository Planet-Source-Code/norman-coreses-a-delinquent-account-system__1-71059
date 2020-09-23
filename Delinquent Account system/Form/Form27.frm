VERSION 5.00
Begin VB.Form Form27 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
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
         Left            =   2520
         TabIndex        =   3
         Top             =   720
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
         Left            =   3720
         TabIndex        =   2
         Top             =   720
         Width           =   1215
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Add New T.O:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1290
      End
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Command1_Click()

Dim cnnewofficer As New ADODB.connection
Dim rsnewofficer As New ADODB.recordset
Dim reply As String

If sempty(Text1) = True Then Exit Sub

Call connection(cnnewofficer, App.Path & "\db1.mdb", "rbp")
Call recordset(rsnewofficer, cnnewofficer, "SELECT * FROM Table5")

'If i

With rsnewofficer
.AddNew
.Fields!t_officer = Text1.Text
.Update
End With

reply = MsgBox("New technical officer has been saved. add another,", vbExclamation + vbYesNo, "Bank of Paracale")
If reply = vbYes Then
    Text1.Text = ""
    Text1.SetFocus
Else
Unload Me
End If

'close connection
Set cnnewofficer = Nothing
'close recordset
Set rsnewofficer = Nothing

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1, vbProperCase)
End Sub


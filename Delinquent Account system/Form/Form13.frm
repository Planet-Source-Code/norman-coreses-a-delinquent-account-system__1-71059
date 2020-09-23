VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7455
      TabIndex        =   11
      Top             =   0
      Width           =   7455
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   480
         Picture         =   "Form13.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Edit Payment"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1665
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
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
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&OR Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Payment:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "&Client Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741




Dim ans As Integer

Private Sub Command1_Click()
Dim cnupdateclient As New ADODB.connection
Dim rsupdateclient As New ADODB.recordset
Dim reply As String

If sempty(Text2) = True Then Exit Sub
If snumber(Text2) = True Then Exit Sub
       
    reply = MsgBox("Are you sure with the information listed", vbInformation + vbYesNo, "Bank of Paracale")
    If reply = vbNo Then
        Exit Sub
        Else
    End If
           
Call connection(cnupdateclient, App.Path & "\db1.mdb", "rbp")
Call recordset(rsupdateclient, cnupdateclient, "SELECT * FROM Table3 WHERE or_number='" & Text3.Text & "'")
      
    With rsupdateclient
        .Fields!d_payment = Text2.Text
        .Update
        .Requery
    End With

    Unload Me

Set cnupdateclient = Nothing
Set rsupdateclient = Nothing

End Sub

Private Sub Command2_Click()
Dim cndeletetransac As New ADODB.connection
Dim rsdeletetransac As New ADODB.recordset

ans = MsgBox("Delete O.R number.", vbInformation + vbYesNo, "Bank of Paracale")
If ans = vbNo Then
Exit Sub
Else
End If
 
Call connection(cndeletetransac, App.Path & "\db1.mdb", "rbp")
Call recordset(rsdeletetransac, cndeletetransac, "SELECT * FROM Table3 WHERE or_number='" & Text3.Text & "'")

With rsdeletetransac
'    Call de
 .Delete
End With

MsgBox "O.R number successfully deleted.", vbInformation, "Bank of Paracale"

Unload Me

Set cndeletetransac = Nothing
Set rsdeletetransac = Nothing

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub




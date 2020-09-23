VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form22 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   Icon            =   "Form22.frx":0000
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -1320
      ScaleHeight     =   735
      ScaleWidth      =   11415
      TabIndex        =   19
      Top             =   0
      Width           =   11415
      Begin VB.Image Image1 
         Height          =   720
         Index           =   4
         Left            =   3480
         Picture         =   "Form22.frx":57E2
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&View System User"
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
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   8325
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4095
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text4 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&View password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Firstname:"
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
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Mi:"
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
         TabIndex        =   17
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Lastname:"
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
         TabIndex        =   16
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Username:"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
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
         TabIndex        =   14
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Confirm password:"
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
         TabIndex        =   13
         Top             =   2040
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Full Name"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text5.PasswordChar = "*"
    Text6.PasswordChar = "*"
Else
    Text5.PasswordChar = ""
    Text6.PasswordChar = ""
End If

End Sub

Private Sub Command1_Click()
    Form23.Show vbModal
End Sub

Private Sub Command2_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No Record found.", vbInformation, "Bank of Paracale": Exit Sub

If Text1.Text = "" Then
MsgBox "Select UserName.", vbInformation, "Bank of Paracale"
Exit Sub
End If

With Form9
.Text1.Text = Text1.Text
.Text2.Text = Text2.Text
.Text3.Text = Text3.Text
.Text4.Text = Text4.Text
.Text5.Text = Text5.Text
.Text6.Text = Text6.Text
.Text7.Text = Text4.Text
.Show vbModal
End With

Call clear

End Sub

Private Sub Command3_Click()
Dim cndelete As New ADODB.connection
Dim rsdelete As New ADODB.recordset
Dim reply As String

If ListView1.ListItems.Count < 1 Then MsgBox "No Record found.", vbExclamation, "Bank of Paracale": Exit Sub

Call connection(cndelete, App.Path & "\db1.mdb", "rbp")
Call recordset(rsdelete, cndelete, "SELECT * FROM UserAccount WHERE UserName='" & ListView1.SelectedItem & "'")

reply = MsgBox("Are you sure you want to delete this username.", vbExclamation + vbYesNo, "Bank of Paracale")
If reply = vbYes Then

With rsdelete
.Delete
'Call d
MsgBox "username successfully deleted.", vbInformation, "Bank of Paracale"
End With
End If

Call fillGotfocus
Call clear

Set cndelete = Nothing
Set rsdelete = Nothing

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call fillGotfocus
End Sub

Private Sub Form_Load()
Call positionform(Form22)
Check1.Value = 1

Call fillGotfocus

End Sub

Private Sub ListView1_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No Record to search.", vbExclamation, "Bank of Paracale": Exit Sub

Text4.Text = ListView1.SelectedItem

Call filllv

End Sub

Public Sub filllv()
Dim cnfill As New ADODB.connection
Dim rsfill As New ADODB.recordset

Call connection(cnfill, App.Path & "\db1.mdb", "rbp")
Call recordset(rsfill, cnfill, "SELECT * FROM UserAccount WHERE UserName='" & Text4.Text & "'")

With rsfill
Text1.Text = .Fields!Firstname
Text2.Text = .Fields!Mi
Text3.Text = .Fields!Lastname
Text5.Text = .Fields!Password
Text6.Text = .Fields!Password
End With

Set cnfill = Nothing
Set rsfill = Nothing

End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Public Sub fillGotfocus()
Dim cnuser As New ADODB.connection
Dim rsuser As New ADODB.recordset
Dim X

Call connection(cnuser, App.Path & "\db1.mdb", "rbp")
Call recordset(rsuser, cnuser, "SELECT * FROM UserAccount ORDER BY UserName ASC")

ListView1.ListItems.clear

With rsuser

    While Not .EOF
            Set X = ListView1.ListItems.Add(, , .Fields!Username)
            X.SubItems(1) = Trim(.Fields!Firstname) & " " & Trim(.Fields!Mi) & " " & Trim(.Fields!Lastname)
            .MoveNext
    Wend
End With

Set cnuser = Nothing
Set rsuser = Nothing

End Sub

Private Sub ListView1_GotFocus()
'fillGotfocus
End Sub

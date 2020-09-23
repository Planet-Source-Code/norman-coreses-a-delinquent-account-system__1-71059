VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Delinquent Account's ver. 1.0"
   ClientHeight    =   4290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7800
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   7740
      TabIndex        =   3
      Top             =   975
      Width           =   7800
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   1005
         ButtonWidth     =   2434
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Payment"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Client"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&T.O"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Center"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Type"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Lock"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Exit"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   7800
      TabIndex        =   2
      Top             =   1590
      Width           =   7800
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   7800
      TabIndex        =   1
      Top             =   0
      Width           =   7800
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Bank of Paracale Inc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   630
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   5580
      End
      Begin VB.Image Image1 
         Height          =   765
         Left            =   840
         Picture         =   "MDIForm1.frx":57E2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1035
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4005
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Username:"
            TextSave        =   "Username:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in:"
            TextSave        =   "Time Log-in:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -38160
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8842
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":911C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B700
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1185C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":12136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13010
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":13E62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuhype33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadd 
         Caption         =   "Add new client"
      End
      Begin VB.Menu mnuhypenc 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditexistingclient 
         Caption         =   "Edit existing client"
      End
      Begin VB.Menu mnuhypeene 
         Caption         =   "-"
      End
      Begin VB.Menu mnutechnicalofficer 
         Caption         =   "Technical Officer"
      End
      Begin VB.Menu mnuhypen94 
         Caption         =   "-"
      End
      Begin VB.Menu mnucenters 
         Caption         =   "Center"
      End
      Begin VB.Menu mnuhypen09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuusersetting 
         Caption         =   "User Setting"
      End
      Begin VB.Menu mnuhypen331 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexitd 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Search"
      Begin VB.Menu mnuhypena 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclient 
         Caption         =   "By Client Name"
      End
      Begin VB.Menu mnuhypen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuto 
         Caption         =   "By Technical Officer"
      End
      Begin VB.Menu mnuhypen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucenter 
         Caption         =   "By Center"
      End
      Begin VB.Menu mnuhypen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnutype 
         Caption         =   "By Type"
      End
      Begin VB.Menu mnuhypen4 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnuhide 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhideshow 
         Caption         =   "Hide/Show Shortcut menu"
      End
      Begin VB.Menu mnushow 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhidetoolbar 
         Caption         =   "Hide/Show Toolbar"
      End
      Begin VB.Menu mnubotton 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuhypen98 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexplorer 
         Caption         =   "Explorer"
      End
      Begin VB.Menu mnuhypen9 
         Caption         =   "-"
      End
      Begin VB.Menu mnucalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuhypeen 
         Caption         =   "-"
      End
      Begin VB.Menu mnunotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnuhypendd 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741


Private Sub MDIForm_Load()
    
    Me.Show
    Form10.Show vbModal
    
    StatusBar1.Panels(2).Text = "" & Username
    
    StatusBar1.Panels(4).Text = Now
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Dim ans As Integer
        ans = MsgBox("Terminate the application.", vbExclamation + vbYesNo, "Bank of Paracale")
                If ans = vbYes Then
                Call timelogout
                End
                    Else
                        Cancel = 1
            End If
End Sub


Private Sub mnuabout_Click()
Form20.Show vbModal
End Sub

Private Sub mnuadd_Click()
    Form1.Show
End Sub





Private Sub mnucalculator_Click()
    On Error Resume Next
        Shell "calc.exe", vbNormalFocus
        
End Sub

Private Sub mnucenter_Click()
Form14.Show
End Sub

Private Sub mnucenters_Click()
Form28.Show vbModal
End Sub

Private Sub mnuclient_Click()
Form12.Show
End Sub

Private Sub mnueditexistingclient_Click()
    Form4.Show
End Sub

Private Sub mnuexitd_Click()
    Dim ans As Integer
        ans = MsgBox("Terminate the application.", vbExclamation + vbYesNo, "Bank of Paracale")
            If ans = vbYes Then
                    Call timelogout
                    End
                Else
                    Exit Sub
            End If
        
End Sub

Private Sub mnuexplorer_Click()
On Error Resume Next
    Shell "explorer.exe", vbNormalFocus
End Sub

Private Sub mnuhideshow_Click()
Picture3.Visible = Not Picture3.Visible
End Sub

Private Sub mnuhidetoolbar_Click()
StatusBar1.Visible = Not StatusBar1.Visible
End Sub

Private Sub mnunotepad_Click()
    On Error Resume Next
        Shell "notepad.exe", vbNormalFocus
        
End Sub

Private Sub mnutechnicalofficer_Click()
Form26.Show vbModal
End Sub

Private Sub mnuto_Click()
Form16.Show
End Sub





Private Sub mnutype_Click()
Form18.Show
End Sub

Private Sub mnuusersetting_Click()
If StatusBar1.Panels(2).Text <> "Admin" Then MsgBox "This function is for Administrator only.", vbCritical, "Bank of Paracale": Exit Sub
Form22.Show vbModal

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2: Form1.Show
        Case 4: Form4.Show
        Case 6: Form11.Show
        Case 8: Form12.Show
        Case 10: Form16.Show
        Case 12: Form14.Show
        Case 14: Form18.Show
        Case 16: Form21.Show vbModal
        Case 18: mnuexitd_Click
    End Select
End Sub

Public Sub timelogout()
Dim cnlogout As New ADODB.connection
Dim rslogout As New ADODB.recordset


Call connection(cnlogout, App.Path & "\db1.mdb", "rbp")

'
'
'

Call recordset(rslogout, cnlogout, "SELECT * FROM Table2 ORDER BY UserName ASC")
    
    
With rslogout
.AddNew
.Fields("UserName") = StatusBar1.Panels(2).Text
.Fields("Time-in") = StatusBar1.Panels(4).Text
.Fields("Date") = Date
.Fields("Time-out") = Now
.Update
 End With

'close connection
Set cnlogout = Nothing
'close recordset
Set rslogout = Nothing

End Sub

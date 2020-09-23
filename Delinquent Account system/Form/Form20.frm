VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " About"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&Contact#: +639212733741"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   1920
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&E-mail : cell_nor@yahoo.com"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Copyright © 2000"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "Version 1.0.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "Delinquent Account's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Image Image1 
         Height          =   885
         Left            =   120
         Picture         =   "Form20.frx":57E2
         Top             =   720
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Close"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This system was designed for the borrower's who  under write - off and under mutual fund application."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'comments or  suggestion please email @ cell_nor@yahoo.com
'if you want full code o f  this system just contact @: 639212733741



Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Command1.Caption = "&Close"
End Sub

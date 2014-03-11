VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SourceSafe Login"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   6975
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "$"
      TabIndex        =   5
      Top             =   960
      Width           =   5655
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
   Begin VB.TextBox txtSourceSafeDatabase 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "INI Location:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOkClicked As Boolean

Private Sub Command1_Click()
    IsOkClicked = True
    Me.Hide
End Sub

Private Sub Command2_Click()
    IsOkClicked = False
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim sCurrentSourceSafeDir As String
    sCurrentSourceSafeDir = GetRegistryString(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\SourceSafe", "Current Database", vbNullString)

    txtSourceSafeDatabase.Text = sCurrentSourceSafeDir
    IsOkClicked = False
End Sub

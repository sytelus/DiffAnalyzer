VERSION 5.00
Begin VB.Form frmSSafeProjectSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select SourceSafe Project"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "SSafeProjectSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin DiffAnalyser.SSafeTree SSafeTree1 
      Height          =   3615
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   6376
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   3750
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1590
      TabIndex        =   0
      Top             =   3750
      Width           =   1065
   End
End
Attribute VB_Name = "frmSSafeProjectSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mbOKPressed As Boolean

Public Function DislayForm(ByVal voVSSDataBase As VSSDatabase, ByRef rsSelectedFolder As String, Optional vsDefaultFolder As String = "$") As Boolean
    mbOKPressed = False
    Set SSafeTree1.oVSSDataBase = voVSSDataBase
    frmSSafeProjectSelector.SSafeTree1.RefreshTree
    frmSSafeProjectSelector.SSafeTree1.SelectedNode = vsDefaultFolder
    frmSSafeProjectSelector.SSafeTree1.ExpandOnClick = False
    frmSSafeProjectSelector.Show vbModal
    Set SSafeTree1.oVSSDataBase = Nothing
    rsSelectedFolder = SSafeTree1.SelectedNode
    DislayForm = mbOKPressed
End Function

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbOKPressed = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    SSafeTree1.SetFocus
End Sub


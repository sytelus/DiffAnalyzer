VERSION 5.00
Begin VB.Form frmSaveToDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Database"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "AddToDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -30
      TabIndex        =   6
      Top             =   1110
      Width           =   4545
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3270
      TabIndex        =   5
      Top             =   1290
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   1290
      Width           =   1065
   End
   Begin VB.CheckBox chkDeleteOldData 
      Caption         =   "&Delete Old Data"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Value           =   1  'Checked
      Width           =   3405
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   345
      Left            =   3390
      TabIndex        =   2
      Top             =   330
      Width           =   975
   End
   Begin VB.TextBox txtDSN 
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection Parameters:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1695
   End
End
Attribute VB_Name = "frmSaveToDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbIsOkPressed As Boolean

Public Function DisplayForm(ByRef rsDefaultDSN As String, ByRef rbDeleteOldData As Boolean) As Boolean

    Dim bFormValidated As Boolean
    Dim sMessage As String
    
    txtDSN.Text = rsDefaultDSN
    chkDeleteOldData.Value = BoolToCheck(rbDeleteOldData)
    
    bFormValidated = False
    
    Do While Not bFormValidated
    
        mbIsOkPressed = False
        
        Me.Show vbModal
        
        If mbIsOkPressed Then
            bFormValidated = ValidateForm(sMessage)
            If bFormValidated Then
                rsDefaultDSN = txtDSN.Text
                rbDeleteOldData = CheckToBool(chkDeleteOldData)
            Else
                MsgBox sMessage
            End If
        Else
            bFormValidated = True
        End If
    Loop
        
    DisplayForm = mbIsOkPressed
    
    Unload Me

End Function

Private Function ValidateForm(ByRef rsMessage As String) As Boolean
    Dim bReturn As Boolean
    rsMessage = vbNullString
    bReturn = True
    
    If IsTextBoxEmpty(txtDSN) Then
        bReturn = False
        rsMessage = rsMessage & "Connection parameters can't be blank" & vbCrLf
    End If
    
    Dim sConnectionError As String
    sConnectionError = CheckConnectionParametersCorrect(txtDSN.Text)
    
    If sConnectionError <> vbNullString Then
        bReturn = False
        rsMessage = rsMessage & "Connection Parameter incorrect: " & sConnectionError
    End If
    
    ValidateForm = bReturn
    
End Function

Private Function IsTextBoxEmpty(ByVal voTextBox As TextBox) As Boolean
    If Trim$(voTextBox.Text) = vbNullString Then
        IsTextBoxEmpty = True
    Else
        IsTextBoxEmpty = False
    End If
End Function

Private Sub cmdBrowse_Click()
    
    On Error GoTo ErrHandler
    
    Dim oConn As ADODB.Connection
    Set oConn = New ADODB.Connection
    
    If CheckConnectionParametersCorrect(txtDSN) = vbNullString Then
        oConn.ConnectionString = txtDSN.Text
    End If
    
    oConn.Properties("Prompt") = 1
    
    oConn.Open
    
    txtDSN = oConn.ConnectionString
    
    Set oConn = Nothing

Exit Sub
ErrHandler:
    Set oConn = Nothing
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbIsOkPressed = True
    Me.Hide
End Sub

Private Function CheckConnectionParametersCorrect(ByVal vsConnectionString As String) As String
    
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    
    CheckConnectionParametersCorrect = vbNullString
    
    Dim oConn As ADODB.Connection
    Set oConn = New ADODB.Connection
    
    oConn.ConnectionString = vsConnectionString
    
    oConn.Open
    
    Set oConn = Nothing
    Me.MousePointer = vbDefault

Exit Function
ErrHandler:
    Me.MousePointer = vbDefault
    CheckConnectionParametersCorrect = Err.Description
End Function


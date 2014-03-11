VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diff Analyser"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowDiffFile 
      Caption         =   "Show Diff File..."
      Height          =   345
      Left            =   4530
      TabIndex        =   16
      Top             =   5790
      Width           =   1545
   End
   Begin VB.CommandButton cmdSaveToDataBase 
      Caption         =   "&Save To Database..."
      Height          =   345
      Left            =   6210
      TabIndex        =   13
      Top             =   5790
      Width           =   1665
   End
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   12
      Top             =   6225
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10398
            Text            =   "Ready"
            TextSave        =   "Ready"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Files Processed:"
            TextSave        =   "Files Processed:"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   30
      TabIndex        =   11
      Top             =   1530
      Width           =   7875
   End
   Begin VB.CommandButton cmdStopProcessing 
      Caption         =   "S&top Analyses"
      Height          =   375
      Left            =   2130
      TabIndex        =   7
      Top             =   1050
      Width           =   1575
   End
   Begin DiffAnalyser.SSafeList SSafeList1 
      Height          =   3885
      Left            =   2370
      TabIndex        =   10
      Top             =   1800
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   6853
   End
   Begin DiffAnalyser.SSafeTree SSafeTree1 
      Height          =   3885
      Left            =   90
      TabIndex        =   9
      Top             =   1800
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   6853
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   30
      TabIndex        =   8
      Top             =   870
      Width           =   7875
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "&Analyse Now"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txtProject1Path 
      Height          =   315
      Index           =   1
      Left            =   4260
      TabIndex        =   4
      Top             =   390
      Width           =   2685
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Br&owse..."
      Height          =   345
      Index           =   1
      Left            =   6990
      TabIndex        =   5
      Top             =   390
      Width           =   885
   End
   Begin VB.TextBox txtProject1Path 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2685
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "B&rowse..."
      Height          =   345
      Index           =   0
      Left            =   2850
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&File Types To Analyse:"
      Height          =   195
      Left            =   4290
      TabIndex        =   15
      Top             =   1110
      Width           =   1605
   End
   Begin MSForms.ComboBox cmbFileExtensions 
      Height          =   315
      Left            =   5940
      TabIndex        =   14
      ToolTipText     =   "Type of files to process"
      Top             =   1080
      Width           =   1935
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "3413;556"
      MatchEntry      =   1
      ListStyle       =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Project &2 to compare:"
      Height          =   195
      Left            =   4260
      TabIndex        =   3
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project &1 to compare:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1


Private mbProcessingCanceled As Boolean
Private mlFilesProcessed As Long

Private msSSafeDir As String
Private msTempDir As String

Dim oVSSDataBase As New SourceSafeTypeLib.VSSDatabase

Private Sub cmdAnalyse_Click()
    
    On Error GoTo ERR_cmdAnalyse_Click
    
    cmdAnalyse.Enabled = False
    cmdStopProcessing.Enabled = True
    
    mbProcessingCanceled = False
    mlFilesProcessed = 0
    
    SSafeList1.ClearFileInfo
    
    
    txtProject1Path(0).Text = RemoveSlashAtEnd(txtProject1Path(0).Text)
    txtProject1Path(1).Text = RemoveSlashAtEnd(txtProject1Path(1).Text)
    
    If Not IsValidSSafeProject(txtProject1Path(0).Text) Then
    
        txtProject1Path(0).SetFocus
        Err.Raise 1000, , "Note valid SourceSafe project: '" & txtProject1Path(0).Text & "'"
        
    ElseIf Not IsValidSSafeProject(txtProject1Path(1).Text) Then
    
        txtProject1Path(1).SetFocus
        Err.Raise 1000, , "Note valid SourceSafe project: '" & txtProject1Path(1).Text & "'"
        
    Else
    
        'Load the tree
        Me.MousePointer = vbHourglass
        Set SSafeTree1.oVSSDataBase = oVSSDataBase
        SSafeTree1.VSSRoot = txtProject1Path(0).Text
        SSafeTree1.TreeRoot = txtProject1Path(0).Text
        
        Call UpdateStatus("Loading items to analyse...")
        
        
        Call SSafeTree1.RefreshTree(True)
        
        Dim oDiffGenerator As DiffGenerator
        Dim oDiffProcessor As DiffProcessor
        
        Set oDiffGenerator = New DiffGenerator
        Set oDiffProcessor = New DiffProcessor
        
        oDiffGenerator.SourceSafePath = msSSafeDir
        oDiffGenerator.TempDir = msTempDir
                
        Dim oclAllowedFileExtensions As Collection
        Set oclAllowedFileExtensions = New Collection
        
        'Quick & dirty - hard coded constants for the time being
        Select Case cmbFileExtensions.ListIndex
            Case 0
                oclAllowedFileExtensions.Add "bas"
                oclAllowedFileExtensions.Add "cls"
                oclAllowedFileExtensions.Add "ctl"
                oclAllowedFileExtensions.Add "dob"
                oclAllowedFileExtensions.Add "frm"
            Case 1
                oclAllowedFileExtensions.Add "c"
                oclAllowedFileExtensions.Add "cpp"
                oclAllowedFileExtensions.Add "h"
            Case 2
                oclAllowedFileExtensions.Add "txt"
                oclAllowedFileExtensions.Add "asc"
                oclAllowedFileExtensions.Add "htm"
                oclAllowedFileExtensions.Add "html"
                oclAllowedFileExtensions.Add "java"
                oclAllowedFileExtensions.Add "js"
                oclAllowedFileExtensions.Add "asp"
                oclAllowedFileExtensions.Add "css"
                oclAllowedFileExtensions.Add "inc"
            Case 3
                oclAllowedFileExtensions.Add "pas"
                oclAllowedFileExtensions.Add "dfm"
            Case 4
                oclAllowedFileExtensions.Add "sql"
                oclAllowedFileExtensions.Add "txt"
            Case 5
                'No restrictions
            Case Else
                cmbFileExtensions.SetFocus
                Err.Raise 1000, , "Unsupported option :" & cmbFileExtensions.Text
        End Select
        
        
                
        'Analyse file in path
        Call AnalyseFilesInProject(txtProject1Path(0).Text, txtProject1Path(1).Text, oDiffGenerator, oDiffProcessor, oclAllowedFileExtensions)
        
        Set oDiffGenerator = Nothing
        Set oDiffProcessor = Nothing
        Set oclAllowedFileExtensions = Nothing
        
    End If
    
    mbProcessingCanceled = False
    Call UpdateStatus("Processing Completed")
    Me.MousePointer = vbDefault
    cmdAnalyse.Enabled = True
    cmdStopProcessing.Enabled = False
    
Exit Sub
ERR_cmdAnalyse_Click:
    ShowError
End Sub

Private Sub AnalyseFilesInProject(ByVal vsPath1 As String, vsPath2 As String, ByVal voDiffGenerator As DiffGenerator, ByVal voDiffProcessor As DiffProcessor, ByVal voclAllowedFileExtensions As Collection)

    Dim lItemIndex As Long
    Dim oVSSItem As VSSItem
    Dim sFile1 As String
    Dim sFile2 As String
    Dim sDiff As String
    Dim sFileExtension As String
    Dim lAllowableExtensionIndex As Long
    Dim bFileExtensionOK As Boolean
    
    If Not mbProcessingCanceled Then
    
        SSafeTree1.SelectedNode = vsPath1
        SSafeList1.PathFilter = vsPath1
        
        Set oVSSItem = oVSSDataBase.VSSItem(vsPath1)
        
        Call UpdateStatus("Analysing items in " & vsPath1 & " ...")
        
        'For each file in Path 1
        For lItemIndex = 1 To oVSSItem.Items.Count
            
            If mbProcessingCanceled Then Exit For
            
            Call UpdateStatus("Finding next item to analyse...")
            
            If oVSSItem.Items.Item(lItemIndex).Type = VSSITEM_FILE Then
            
                sFileExtension = ExtractFileExtension(oVSSItem.Items.Item(lItemIndex).Name)
                
                If voclAllowedFileExtensions.Count >= 1 Then
                    bFileExtensionOK = False
                    For lAllowableExtensionIndex = 1 To voclAllowedFileExtensions.Count
                        If LCase(sFileExtension) = LCase(voclAllowedFileExtensions(lAllowableExtensionIndex)) Then
                            bFileExtensionOK = True
                            Exit For
                        End If
                    Next lAllowableExtensionIndex
                Else
                    bFileExtensionOK = True
                End If
                            
                If bFileExtensionOK Then
                    'Build the path for file 1
                    sFile1 = vsPath1 & "/" & oVSSItem.Items.Item(lItemIndex).Name
                    
                    Call UpdateStatus("Analysing file " & sFile1 & " ...")
                    
                    'Build the path for file 2
                    sFile2 = vsPath2 & "/" & oVSSItem.Items.Item(lItemIndex).Name
                    
                    'Generate the diff
                    sDiff = voDiffGenerator.GetDiff(sFile1, sFile2)
                    
                    'Ananlyse the diff
                    Call voDiffProcessor.ProcessDiff(sDiff)
                    
                    'Update list view
                    Call SSafeList1.UpdateFileInfo(vsPath1, oVSSItem.Items.Item(lItemIndex).Name, voDiffProcessor)
                    
                    mlFilesProcessed = mlFilesProcessed + 1
                    
                    stbMain.Panels(2).Text = "Files Processed: " & mlFilesProcessed
                    
                End If
            End If
            
        Next lItemIndex
        
        'For each file in Path 1
        For lItemIndex = 1 To oVSSItem.Items.Count
            
            If mbProcessingCanceled Then Exit For
            
            Call UpdateStatus("Finding next project to analyse...")
            
            If oVSSItem.Items.Item(lItemIndex).Type = VSSITEM_PROJECT Then
            
                'Build the path for project1
                sFile1 = vsPath1 & "/" & oVSSItem.Items.Item(lItemIndex).Name
                
                Call UpdateStatus("Analysing project " & sFile1 & " ...")
                
                'Build the path for file 2
                sFile2 = vsPath2 & "/" & oVSSItem.Items.Item(lItemIndex).Name
                
                Call AnalyseFilesInProject(sFile1, sFile2, voDiffGenerator, voDiffProcessor, voclAllowedFileExtensions)
                
            End If
            
        Next lItemIndex
    
    End If
    
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim bIsCanceled As Boolean
    Dim sSelectedProject As String
    bIsCanceled = Not frmSSafeProjectSelector.DislayForm(oVSSDataBase, sSelectedProject, txtProject1Path(Index).Text)
    If Not bIsCanceled Then
        txtProject1Path(Index).Text = sSelectedProject
    End If
End Sub

Private Sub cmdSaveToDataBase_Click()
    
    On Error GoTo ErrHandler
    
    Dim bSaveToDatabaseDialogReturn As Boolean
    Dim sConnectionString As String
    Dim bDeleteOldData As Boolean
    
    sConnectionString = GetSetting("DiffAnalyser", "Settings", "ConnectionString", vbNullString)
    bDeleteOldData = False
    bSaveToDatabaseDialogReturn = frmSaveToDatabase.DisplayForm(sConnectionString, bDeleteOldData)
    If bSaveToDatabaseDialogReturn Then
        Call SaveSetting("DiffAnalyser", "Settings", "ConnectionString", sConnectionString)
        Call SaveToDatabase(sConnectionString, bDeleteOldData)
    End If
Exit Sub
ErrHandler:
    ShowError
End Sub

Private Sub SaveToDatabase(ByVal vsConnectionString As String, ByVal vboolDeleteOldData As Boolean)
    
    On Error GoTo ERRTrap
    
    Dim oConn As ADODB.Connection
    Set oConn = New ADODB.Connection
    
    Const sMAIN_TABLE As String = "DiffAnalyses"
    
    Me.MousePointer = vbHourglass
    oConn.Open vsConnectionString
    
    Dim rsMainTable As ADODB.Recordset
    Set rsMainTable = New ADODB.Recordset
    
    If vboolDeleteOldData Then
        Call oConn.Execute("DELETE FROM " & sMAIN_TABLE)
    End If
    
    Dim sSQL As String
    
    sSQL = "SELECT * FROM " & sMAIN_TABLE & " WHERE 1=2"
    
    Call rsMainTable.Open(sSQL, oConn, adOpenForwardOnly, adLockOptimistic)
    
    Dim lListIndex As Long
        
    'For each row in all file info list, add row in table
    '
    'Currently data is bound with UI - will be improved later...
    '
    With SSafeList1.AllFileInfoList.ListItems
        For lListIndex = 1 To .Count
            rsMainTable.AddNew
            rsMainTable("FileName") = .Item(lListIndex).Text
            rsMainTable("FilePath") = ExtractFilePath(.Item(lListIndex).Key, "/")
            rsMainTable("DeletedLineCount") = .Item(lListIndex).SubItems(1)
            rsMainTable("InsertedLineCount") = .Item(lListIndex).SubItems(2)
            rsMainTable("ChangedLineCount") = .Item(lListIndex).SubItems(3)
            rsMainTable("IgnoredLineCount") = .Item(lListIndex).SubItems(4)
            rsMainTable("IsFileExist") = (.Item(lListIndex).SmallIcon <> 2)
            rsMainTable.Update
        Next lListIndex
    End With
    
    Me.MousePointer = vbDefault
    Set rsMainTable = Nothing
    Set oConn = Nothing
    
Exit Sub
ERRTrap:
    Set rsMainTable = Nothing
    Set oConn = Nothing
    ReRaiseError
End Sub

Private Sub cmdShowDiffFile_Click()

    On Error GoTo ERRTrap

    Dim sDiffFile As String
    Dim sFile1 As String
    Dim sFile2 As String
    
    Me.MousePointer = vbHourglass
    sFile1 = SSafeList1.SelectedFile
    If sFile1 <> vbNullString Then
        
        sFile2 = txtProject1Path(1).Text & Mid$(sFile1, Len(txtProject1Path(1).Text) + 1)
        
        Dim oDiffGenerator As DiffGenerator
        Set oDiffGenerator = New DiffGenerator
        oDiffGenerator.SourceSafePath = msSSafeDir
        oDiffGenerator.TempDir = msTempDir
        sDiffFile = oDiffGenerator.MakeDiffFile(sFile1, sFile2)
        Call ShellExecute(0, "open", sDiffFile, "", "", SW_NORMAL)
        Set oDiffGenerator = Nothing
    Else
        Err.Raise 1000, , "No file selected. Please select a file for which you want to see the difference."
    End If
    Me.MousePointer = vbDefault
    
Exit Sub
ERRTrap:
    ShowError
    Call DeleteFile(sDiffFile)
    Set oDiffGenerator = Nothing
End Sub

Private Sub cmdStopProcessing_Click()
    mbProcessingCanceled = True
End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load

    msSSafeDir = GetRegistryString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SourceSafe", "API Current Database", vbNullString)
    
    If msSSafeDir = vbNullString Then
        Err.Raise 1000, , "SourceSafe Client path is not found in your machine. This could be probably because SourceSafe client is not installed in machine. This program can not work if SourceSafe client is not available."
    End If
    
    msTempDir = GetTempDirPath
    
    txtProject1Path(0).Text = GetSetting("DiffAnalyser", "LastValues", "Project1Path", vbNullString)
    txtProject1Path(1).Text = GetSetting("DiffAnalyser", "LastValues", "Project2Path", vbNullString)
    
    Call FillFilterExtensionCombo(cmbFileExtensions)
    Load frmLogin
    frmLogin.Show 1, Me
    
    If frmLogin.IsOkClicked = False Then
        Call oVSSDataBase.Open
    Else
        Call oVSSDataBase.Open(frmLogin.txtSourceSafeDatabase.Text, frmLogin.txtUserName.Text, frmLogin.txtPassword.Text)
    End If
    Unload frmLogin
    
    cmbFileExtensions.ListIndex = 0
    cmdStopProcessing.Enabled = False
    mbProcessingCanceled = False
    mlFilesProcessed = 0
    
Exit Sub
Form_Load:
    ShowError
    End 'If loading failed then end! Currently no way to test sevearity.
End Sub

Private Sub FillFilterExtensionCombo(ByVal vcmbComboBox As MSForms.ComboBox)
    With vcmbComboBox
        .AddItem "Analyse VB Files (*.bas;*.cls;*.frm;*.ctl;*.dob)"
        .AddItem "Analyse C/C + Files (*.c;*.cpp;*.h)"
        .AddItem "Analyse Web Files (*.txt;*.htm;*.html;*.java;*.js;*.css;*.inc;*.asp)"
        .AddItem "Analyse Delphi Files (*.pas;*.dfm)"
        .AddItem "Analyse Text And Data Scripts (*.sql;*.txt;*.asc)"
        .AddItem "Analyse All Files (*.*)"
        
        Dim lListIndex As Long
        Dim lMaxItemWidth As Long
        
        lMaxItemWidth = 0
        For lListIndex = 0 To .ListCount - 1
            If lMaxItemWidth < Me.TextWidth(.List(lListIndex) & "           ") Then
                lMaxItemWidth = Me.TextWidth(.List(lListIndex) & "           ")
            End If
        Next lListIndex
        If ((.Width / 20) < lMaxItemWidth) Then
            .ListWidth = CLng(lMaxItemWidth) / 20  'Twips to Points conversion
        End If
        
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("DiffAnalyser", "LastValues", "Project1Path", txtProject1Path(0).Text)
    Call SaveSetting("DiffAnalyser", "LastValues", "Project2Path", txtProject1Path(1).Text)
    Set oVSSDataBase = Nothing
    SSafeList1.ClearFileInfo
    End
End Sub

Private Function IsPathValid(ByVal vsPath As String) As Boolean
    On Error Resume Next
    Dim oVSSItem As VSSItem
    Set oVSSDataBase = oVSSDataBase.VSSItem(vsPath)
    IsPathValid = (Err.Number = 0)
End Function

Private Function IsValidSSafeProject(ByVal vsPath As String) As Boolean
    On Error Resume Next
    Dim oVSSItem As VSSItem
    Set oVSSItem = oVSSDataBase.VSSItem(vsPath)
    IsValidSSafeProject = (Err.Number = 0)
End Function

Private Sub ShowError()
    cmdAnalyse.Enabled = True
    cmdStopProcessing.Enabled = False
    mbProcessingCanceled = False
    Me.MousePointer = vbDefault
    Call UpdateStatus("Error occured: " & Err.Description)
    Utils.ShowError
End Sub

Private Sub UpdateStatus(ByVal vsStatusString As String)
    stbMain.Panels(1).Text = vsStatusString
    stbMain.Panels(1).ToolTipText = stbMain.Panels(1).Text
    DoEvents
End Sub

Private Sub SSafeTree1_NodeClicked(Node As ComctlLib.Node)
    SSafeList1.PathFilter = Node.FullPath
End Sub

Private Function IsValidDir(ByVal vsPath As String) As Boolean
    On Error Resume Next
    Dim sTemp As String
    sTemp = Dir(GetPathWithSlash(vsPath) & "*.*")
    IsValidDir = (Err.Number = 0)
End Function


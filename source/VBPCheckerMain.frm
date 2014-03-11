VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmVBPCheckerMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SourceSafe File Checker"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   Icon            =   "VBPCheckerMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwFiles 
      Height          =   3840
      Left            =   2400
      TabIndex        =   11
      Top             =   2325
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "FileName"
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "User"
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "Date"
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   "Comment"
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   4304
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtFileDateToCheckAfter 
      Height          =   315
      Left            =   2700
      TabIndex        =   12
      Top             =   825
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      Format          =   22806528
      CurrentDate     =   37223
   End
   Begin SSafeFileChecker.SSafeTree SSafeTree1 
      Height          =   3840
      Left            =   75
      TabIndex        =   10
      Top             =   2325
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   6773
   End
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   6315
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15769
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Files Processed:"
            TextSave        =   "Files Processed:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   30
      TabIndex        =   6
      Top             =   2055
      Width           =   10725
   End
   Begin VB.CommandButton cmdStopProcessing 
      Caption         =   "S&top Checking"
      Height          =   375
      Left            =   1830
      TabIndex        =   4
      Top             =   1575
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   30
      TabIndex        =   5
      Top             =   1395
      Width           =   10800
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "&Check Now"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1575
      Width           =   1575
   End
   Begin VB.TextBox txtProject1Path 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4560
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "B&rowse..."
      Height          =   345
      Index           =   0
      Left            =   4725
      TabIndex        =   2
      Top             =   375
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Find Files Changed After This Date:"
      Height          =   195
      Left            =   150
      TabIndex        =   13
      Top             =   825
      Width           =   2505
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&File Types To Check:"
      Height          =   195
      Left            =   4290
      TabIndex        =   9
      Top             =   1635
      Width           =   1515
   End
   Begin MSForms.ComboBox cmbFileExtensions 
      Height          =   315
      Left            =   5940
      TabIndex        =   8
      ToolTipText     =   "Type of files to process"
      Top             =   1605
      Width           =   4860
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "8572;556"
      MatchEntry      =   1
      ListStyle       =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project To Check:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmVBPCheckerMain"
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
    
    lvwFiles.ListItems.Clear
    
    
    txtProject1Path(0).Text = RemoveSlashAtEnd(txtProject1Path(0).Text)
    
    If Not IsValidSSafeProject(txtProject1Path(0).Text) Then
    
        txtProject1Path(0).SetFocus
        Err.Raise 1000, , "Note valid SourceSafe project: '" & txtProject1Path(0).Text & "'"
        
    Else
    
        'Load the tree
        Me.MousePointer = vbHourglass
        Set SSafeTree1.oVSSDataBase = oVSSDataBase
        SSafeTree1.VSSRoot = txtProject1Path(0).Text
        SSafeTree1.TreeRoot = txtProject1Path(0).Text
        
        Call UpdateStatus("Loading items to analyse...")
        
        
        Call SSafeTree1.RefreshTree(True)
                
        Dim oclAllowedFileExtensions As Collection
        Set oclAllowedFileExtensions = New Collection
        
        'Quick & dirty - hard coded constants for the time being
        Select Case cmbFileExtensions.ListIndex
            Case 0
                oclAllowedFileExtensions.Add "vbp"
            Case 1
                oclAllowedFileExtensions.Add "bas"
                oclAllowedFileExtensions.Add "cls"
                oclAllowedFileExtensions.Add "ctl"
                oclAllowedFileExtensions.Add "dob"
                oclAllowedFileExtensions.Add "frm"
            Case 2
                oclAllowedFileExtensions.Add "c"
                oclAllowedFileExtensions.Add "cpp"
                oclAllowedFileExtensions.Add "h"
            Case 3
                oclAllowedFileExtensions.Add "txt"
                oclAllowedFileExtensions.Add "asc"
                oclAllowedFileExtensions.Add "htm"
                oclAllowedFileExtensions.Add "html"
                oclAllowedFileExtensions.Add "java"
                oclAllowedFileExtensions.Add "js"
                oclAllowedFileExtensions.Add "asp"
                oclAllowedFileExtensions.Add "css"
                oclAllowedFileExtensions.Add "inc"
            Case 4
                oclAllowedFileExtensions.Add "pas"
                oclAllowedFileExtensions.Add "dfm"
            Case 5
                oclAllowedFileExtensions.Add "sql"
                oclAllowedFileExtensions.Add "txt"
            Case 6
                'No restrictions
            Case Else
                cmbFileExtensions.SetFocus
                Err.Raise 1000, , "Unsupported option :" & cmbFileExtensions.Text
        End Select
                
        'Analyse file in path
        Call AnalyseFilesInProject(txtProject1Path(0).Text, oclAllowedFileExtensions)
        
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

Private Sub AnalyseFilesInProject(ByVal vsPath1 As String, ByVal voclAllowedFileExtensions As Collection)

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
        
        Set oVSSItem = oVSSDataBase.VSSItem(vsPath1)
        
        Call UpdateStatus("Analysing items in " & vsPath1 & " ...")
        
        'For each file in Path 1
        For lItemIndex = 1 To oVSSItem.Items.Count
            
            If mbProcessingCanceled Then Exit For
            
            Call UpdateStatus("Finding next item to analyse...")
            
            With oVSSItem.Items.Item(lItemIndex)
                If .Type = VSSITEM_FILE Then
                        sFileExtension = ExtractFileExtension(.Name)
                        
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
                            sFile1 = vsPath1 & "/" & .Name
                            
                            Call UpdateStatus("Analysing file " & sFile1 & " ...")
                            
                            Dim lThisVersion As Long
                            Dim oVersion As VSSVersion
                            Dim dtDateForThisVersion As Date
                            dtDateForThisVersion = Now
                            lThisVersion = .VersionNumber
                            Dim oThisVersion As VSSVersion
                            For Each oVersion In .Versions
                                If oVersion.VersionNumber = lThisVersion Then
                                    dtDateForThisVersion = oVersion.Date
                                    Set oThisVersion = oVersion
                                    Exit For
                                End If
                            Next oVersion
                                
                            If dtDateForThisVersion >= dtFileDateToCheckAfter.Value Then
                                'Update list view
                                Dim lvlFileDetail As ListItem
                                Set lvlFileDetail = lvwFiles.ListItems.Add(, , .Name)
                                If Not (oThisVersion Is Nothing) Then
                                    lvlFileDetail.SubItems(1) = oThisVersion.Username
                                    lvlFileDetail.SubItems(2) = oThisVersion.Date
                                    lvlFileDetail.SubItems(3) = oThisVersion.Comment
                                Else
                                    lvlFileDetail.SubItems(1) = "n/a"
                                    lvlFileDetail.SubItems(2) = "n/a"
                                    lvlFileDetail.SubItems(3) = "n/a"
                                End If
                            End If
                            
                            mlFilesProcessed = mlFilesProcessed + 1
                            
                            stbMain.Panels(2).Text = "Files Processed: " & mlFilesProcessed
                            
                        End If
                End If
            End With
            
        Next lItemIndex
        
        'For each file in Path 1
        For lItemIndex = 1 To oVSSItem.Items.Count
            
            If mbProcessingCanceled Then Exit For
            
            Call UpdateStatus("Finding next project to analyse...")
            
            If oVSSItem.Items.Item(lItemIndex).Type = VSSITEM_PROJECT Then
            
                'Build the path for project1
                sFile1 = vsPath1 & "/" & oVSSItem.Items.Item(lItemIndex).Name
                
                Call UpdateStatus("Analysing project " & sFile1 & " ...")
                
                Call AnalyseFilesInProject(sFile1, voclAllowedFileExtensions)
                
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
    
    Call FillFilterExtensionCombo(cmbFileExtensions)
        
    oVSSDataBase.Open "\\atidevsql\vsstest"
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
        .AddItem "Analyse VBP Files (*.vbp)"
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
    Set oVSSDataBase = Nothing
    lvwFiles.ListItems.Clear
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

Private Function IsValidDir(ByVal vsPath As String) As Boolean
    On Error Resume Next
    Dim sTemp As String
    sTemp = Dir(GetPathWithSlash(vsPath) & "*.*")
    IsValidDir = (Err.Number = 0)
End Function


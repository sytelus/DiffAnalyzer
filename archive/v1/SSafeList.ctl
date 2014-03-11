VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.UserControl SSafeList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin ComctlLib.ListView lvwMain 
      Height          =   3285
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "file"
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "deleted lines"
         Object.Tag             =   ""
         Text            =   "Deleted Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "inserted lines"
         Object.Tag             =   ""
         Text            =   "Inserted Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   "changed lines"
         Object.Tag             =   ""
         Text            =   "Changed Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   "ignored lines"
         Object.Tag             =   ""
         Text            =   "Ignored Lines"
         Object.Width           =   706
      EndProperty
   End
   Begin ComctlLib.ListView lvwAllFileInfo 
      Height          =   3285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "file"
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   "deleted lines"
         Object.Tag             =   ""
         Text            =   "Deleted Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   "inserted lines"
         Object.Tag             =   ""
         Text            =   "Inserted Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   "changed lines"
         Object.Tag             =   ""
         Text            =   "Changed Lines"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   "ignored lines"
         Object.Tag             =   ""
         Text            =   "Ignored Lines"
         Object.Width           =   706
      EndProperty
   End
   Begin ComctlLib.ImageList imlMain 
      Left            =   4350
      Top             =   1530
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SSafeList.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SSafeList.ctx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SSafeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sPathFilter As String
Private mlFilterPathLen As Long

Public Property Get SelectedFile() As String
    If Not (lvwMain.SelectedItem Is Nothing) Then
        SelectedFile = lvwMain.SelectedItem.Key
    Else
        SelectedFile = vbNullString
    End If
End Property

Public Property Let SelectedFile(ByVal vsFileSpec As String)
    lvwMain.ListItems(vsFileSpec).Selected = True
End Property

Public Property Get AllFileInfoList() As ListView
    Set AllFileInfoList = lvwAllFileInfo
End Property

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwMain.SortKey = ColumnHeader.Index - 1
    If ColumnHeader.Tag = 0 Then
        lvwMain.SortOrder = lvwDescending
        ColumnHeader.Tag = 1
    Else
        lvwMain.SortOrder = lvwAscending
        ColumnHeader.Tag = 0
    End If
    lvwMain.Sorted = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lNosOfColumns As Long
    Dim lColumnWidth As Double
    Dim lColumnHeaderIndex As Long
    
    lColumnWidth = (lvwMain.Width - lvwMain.ColumnHeaders("file").Width) / (lvwMain.ColumnHeaders.Count - 1)
    lvwMain.ColumnHeaders(1).Tag = 0
    For lColumnHeaderIndex = 2 To lvwMain.ColumnHeaders.Count
        lvwMain.ColumnHeaders(lColumnHeaderIndex).Width = lColumnWidth
        lvwMain.ColumnHeaders(lColumnHeaderIndex).Tag = 0
    Next lColumnHeaderIndex
    
End Sub

Public Sub UpdateFileInfo(ByVal vsPath As String, ByVal vsFileName As String, ByVal voDiffProcessor As Object)
    Dim sFullFileSpec As String
    
    sFullFileSpec = vsPath & "/" & vsFileName
    
    Call UpdateListItem(lvwAllFileInfo, sFullFileSpec, vsFileName, voDiffProcessor)
    
    If IsPathFiltered(sFullFileSpec) Then
        Call UpdateListItem(lvwMain, sFullFileSpec, vsFileName, voDiffProcessor)
    End If
End Sub

Private Sub UpdateListItem(ByVal vlvwListView As ListView, ByVal vsFullFileSpec As String, ByVal vsFileName As String, ByVal voDiffProcessor As Object)
    With vlvwListView
        If Not ItemExistInList(vlvwListView, vsFullFileSpec) Then
            .ListItems.Add , vsFullFileSpec, vsFileName
        End If
        
        With .ListItems(vsFullFileSpec)
            .SubItems(1) = CStr(voDiffProcessor.DeletedLineCount)
            .SubItems(2) = CStr(voDiffProcessor.InsertedLineCount)
            .SubItems(3) = CStr(voDiffProcessor.ChangedLineCount)
            .SubItems(4) = CStr(voDiffProcessor.IgnoredLineCount)
            If voDiffProcessor.IsFileExist Then
                If voDiffProcessor.DeletedLineCount <> 0 Then
                    .SmallIcon = 1
                Else
                    .SmallIcon = 0
                End If
            Else
                .SmallIcon = 2
            End If
        End With
    End With
End Sub

Private Function IsPathFiltered(ByVal vsPath As String) As Boolean
    IsPathFiltered = (StrComp(Left(vsPath, mlFilterPathLen), m_sPathFilter, 1) = 0)
End Function

Public Property Get PathFilter() As String
    PathFilter = m_sPathFilter
End Property

Public Property Let PathFilter(ByVal vsPath As String)
    m_sPathFilter = RemoveSlashAtEnd(vsPath)
    mlFilterPathLen = Len(m_sPathFilter)
    Call RefreshList
End Property

Public Sub RefreshList()
    
    lvwMain.ListItems.Clear
    
    Dim lAllFilesListIndex As Long
    Dim sFileName As String
    Dim sFullFileSpec As String
    Dim lSubItemIndex As Long
    Dim lsiListItem As ListItem
    Dim sFilePath As String
    
    'For each item in all file info list
    For lAllFilesListIndex = 1 To lvwAllFileInfo.ListItems.Count
        sFullFileSpec = lvwAllFileInfo.ListItems(lAllFilesListIndex).Key
        sFilePath = ExtractFilePath(sFullFileSpec, "/")
        If StrComp(sFilePath, m_sPathFilter, 1) = 0 Then
            sFileName = Mid(sFullFileSpec, mlFilterPathLen + 2)
            Set lsiListItem = lvwMain.ListItems.Add(, sFullFileSpec, sFileName)
            lsiListItem.SmallIcon = lvwAllFileInfo.ListItems(lAllFilesListIndex).SmallIcon
            For lSubItemIndex = 1 To lvwAllFileInfo.ColumnHeaders.Count - 1
                lsiListItem.SubItems(lSubItemIndex) = lvwAllFileInfo.ListItems(lAllFilesListIndex).SubItems(lSubItemIndex)
            Next lSubItemIndex
        End If
    Next lAllFilesListIndex
    
End Sub

Public Sub ClearFileInfo()
    lvwAllFileInfo.ListItems.Clear
    lvwMain.ListItems.Clear
End Sub

Private Function ItemExistInList(ByVal vlvwListView As ListView, ByVal vsKey As String) As Boolean
    On Error Resume Next
    Dim litItem As ListItem
    
    Set litItem = vlvwListView.ListItems(vsKey)
    ItemExistInList = (Err.Number = 0)
End Function
Private Sub UserControl_Resize()
    With lvwMain
        .Height = UserControl.Height
        .Width = UserControl.Width
        .Left = 0
        .Top = 0
    End With
End Sub

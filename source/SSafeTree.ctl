VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl SSafeTree 
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4065
   ScaleWidth      =   4800
   Begin ComctlLib.TreeView tvwMain 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   6482
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   71
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "/"
      Style           =   7
      ImageList       =   "imlMain"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ComctlLib.ImageList imlMain 
      Left            =   2550
      Top             =   0
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
            Picture         =   "SSafeTree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SSafeTree.ctx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SSafeTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public oVSSDataBase  As VSSDatabase
Public VSSRoot As String
Public TreeRoot As String
Public ExpandOnClick As Boolean

Public Event NodeClicked(Node As Node)

Const sSSAFE_ROOT As String = "$"


Public Property Get TreeView() As TreeView
    Set TreeView = tvwMain
End Property

Public Property Get SelectedNode() As String
    Dim oSelectedNode As Node
    Set oSelectedNode = tvwMain.SelectedItem
    If Not (oSelectedNode Is Nothing) Then
        SelectedNode = oSelectedNode.FullPath
    Else
        SelectedNode = vbNullString
    End If
End Property

Public Property Let SelectedNode(ByVal vsPath As String)
    If IsValidSSafeProject(vsPath) Then
        Dim sSelectedPath As String
        sSelectedPath = SelectPath(vsPath)
        Call SilentSelectNode(sSelectedPath)
    Else
        Set tvwMain.SelectedItem = Nothing
    End If
End Property

Private Sub UserControl_Initialize()
    VSSRoot = sSSAFE_ROOT & "/"
    TreeRoot = sSSAFE_ROOT
    ExpandOnClick = True
End Sub

Private Sub UserControl_Resize()
    tvwMain.Width = UserControl.Width
    tvwMain.Height = UserControl.Height
    tvwMain.Top = 0
    tvwMain.Left = 0
End Sub

Public Function RefreshTree(Optional ByVal vboolFullrefresh As Boolean = False)
    
    Dim oVSSItem As SourceSafeTypeLib.VSSItem
    
    Set oVSSItem = oVSSDataBase.VSSItem(VSSRoot)

    If vboolFullrefresh Then
        tvwMain.Nodes.Clear
    End If

    If tvwMain.Nodes.Count = 0 Then
        Dim oNode As Node
        Set oNode = tvwMain.Nodes.Add(, , TreeRoot, TreeRoot, 1, 2)
        oNode.Tag = 0
        Call FillNodeWithSubItems(oNode)
        oNode.Expanded = True
    End If
    
End Function

Private Sub FillNodeWithSubItems(ByVal voNode As Node)
    
    On Error GoTo ERR_FillNodeWithSubItems
    
    Dim oTempGrandChild As Node
    Dim oChildNode As Node
    Dim vlVSSItemIndex As Long
    Dim oVSSItems As VSSItem
    Dim oGrandChildNode As Node
    Dim lGrandChildIndex As Long
    Dim bGrandChildExist As Boolean
    Dim oVSSChildItem As VSSItem
    
    Screen.MousePointer = vbHourglass
    
    'If this node is not filled
    If voNode.Tag <> 1 Then
        Call SilentRemove(voNode.FullPath & "/Temp")
        Set oVSSItems = oVSSDataBase.VSSItem(voNode.FullPath & "/")
        For vlVSSItemIndex = 1 To oVSSItems.Items.Count
            If oVSSItems.Items.Item(vlVSSItemIndex).Type = VSSITEM_PROJECT Then
                Set oChildNode = tvwMain.Nodes.Add(voNode, tvwChild, , oVSSItems.Items.Item(vlVSSItemIndex).Name, 1, 2)
                oChildNode.Key = oChildNode.FullPath
                oChildNode.Tag = 0
                bGrandChildExist = False
                Set oVSSChildItem = oVSSItems.Items.Item(vlVSSItemIndex)
                For lGrandChildIndex = 1 To oVSSChildItem.Items.Count
                    If oVSSChildItem.Items.Item(lGrandChildIndex).Type = VSSITEM_PROJECT Then
                        bGrandChildExist = True
                        Exit For
                    End If
                Next lGrandChildIndex
                If bGrandChildExist Then
                    Set oTempGrandChild = tvwMain.Nodes.Add(oChildNode, tvwChild, oChildNode.FullPath & "/Temp", "Temp")
                End If
            End If
        Next vlVSSItemIndex
        voNode.Tag = 1
    End If
    
    Screen.MousePointer = vbDefault
    
Exit Sub
ERR_FillNodeWithSubItems:
    Screen.MousePointer = vbDefault
    ReRaiseError
End Sub

Private Sub tvwMain_Expand(ByVal Node As ComctlLib.Node)
    Call FillNodeIfNotFilled(Node)
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As ComctlLib.Node)
    If ExpandOnClick Then
        Call FillNodeIfNotFilled(Node)
    End If
    RaiseEvent NodeClicked(Node)
End Sub

Private Sub FillNodeIfNotFilled(ByVal Node As ComctlLib.Node)
    On Error GoTo ErrorTrap
    If Node.Tag <> 1 Then
        Call FillNodeWithSubItems(Node)
        Node.Expanded = True
    End If
    
Exit Sub
ErrorTrap:
    ShowError
End Sub

Private Sub SilentRemove(ByVal vsPath As String)
    On Error Resume Next
    Call tvwMain.Nodes.Remove(vsPath)
End Sub

Private Function IsValidSSafeProject(ByVal vsPath As String) As Boolean
    On Error Resume Next
    Dim oVSSItem As VSSItem
    Set oVSSItem = oVSSDataBase.VSSItem(vsPath)
    IsValidSSafeProject = (Err.Number = 0)
End Function

Public Function SelectPath(ByVal vsPath As String) As String
    Dim lProjectNameStart As Long
    Dim lProjectNameEnd As Long
    Dim sProjectName As String
    Dim sBuildedPath As String
    
    sBuildedPath = VSSRoot
    
    'Expand this node
    tvwMain.Nodes(TreeRoot).Expanded = True
    tvwMain.Nodes(TreeRoot).EnsureVisible
    
    
    lProjectNameStart = Len(TreeRoot) + 2
    lProjectNameEnd = InStr(lProjectNameStart, vsPath, "/")
    
    Do While Not (lProjectNameEnd = 0)
    
        sProjectName = Mid(vsPath, lProjectNameStart, lProjectNameEnd - lProjectNameStart)
        
        If sBuildedPath <> vbNullString Then
            sBuildedPath = GetPathWithSlash(sBuildedPath, "/") & sProjectName
        Else
            sBuildedPath = sProjectName
        End If
        
        'Expand this node
        tvwMain.Nodes(sBuildedPath).Expanded = True
        tvwMain.Nodes(sBuildedPath).EnsureVisible
        
        lProjectNameStart = lProjectNameEnd + 1
        lProjectNameEnd = InStr(lProjectNameStart, vsPath, "/")

    Loop
    
    If lProjectNameStart < Len(vsPath) Then
    
        sProjectName = Mid(vsPath, lProjectNameStart)
        
        If sBuildedPath <> vbNullString Then
            sBuildedPath = GetPathWithSlash(sBuildedPath, "/") & sProjectName
        Else
            sBuildedPath = sProjectName
        End If
        
        'Expand this node
        tvwMain.Nodes(sBuildedPath).Expanded = True
        tvwMain.Nodes(sBuildedPath).EnsureVisible
    
    End If
    
    SelectPath = sBuildedPath
    
End Function

Private Sub SilentSelectNode(ByVal vsPath As String)
    On Error Resume Next
    tvwMain.Nodes(vsPath).Selected = True
End Sub


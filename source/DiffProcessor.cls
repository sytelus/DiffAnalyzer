VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarChangedLineCount As Long 'local copy
Private mvarDeletedLineCount As Long 'local copy
Private mvarInsertedLineCount As Long 'local copy
Private mvarIgnoredLineCount As Long 'local copy
Private mvarTotalLineCount As Long 'local copy
Private mvarBlankLineCount As Long 'local copy
Private mvarCommentedLineCount As Long 'local copy

Private mbIsFileExist As Boolean

    Const lDIFF_INDICATOR_POS As Long = 39


Public Property Let CommentedLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommentedLineCount = 5
    mvarCommentedLineCount = vData
End Property


Public Property Get CommentedLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommentedLineCount
    CommentedLineCount = mvarCommentedLineCount
End Property



Public Property Let BlankLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BlankLineCount = 5
    mvarBlankLineCount = vData
End Property


Public Property Get BlankLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BlankLineCount
    BlankLineCount = mvarBlankLineCount
End Property



Public Property Let TotalLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalLineCount = 5
    mvarTotalLineCount = vData
End Property


Public Property Get TotalLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalLineCount
    TotalLineCount = mvarTotalLineCount
End Property



Public Property Let IgnoredLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IgnoredLineCount = 5
    mvarIgnoredLineCount = vData
End Property


Public Property Get IgnoredLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IgnoredLineCount
    IgnoredLineCount = mvarIgnoredLineCount
End Property



Public Property Let InsertedLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InsertedLineCount = 5
    mvarInsertedLineCount = vData
End Property


Public Property Get InsertedLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InsertedLineCount
    InsertedLineCount = mvarInsertedLineCount
End Property



Public Property Let DeletedLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeletedLineCount = 5
    mvarDeletedLineCount = vData
End Property


Public Property Get DeletedLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeletedLineCount
    DeletedLineCount = mvarDeletedLineCount
End Property



Public Property Let ChangedLineCount(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChangedLineCount = 5
    mvarChangedLineCount = vData
End Property


Public Property Get ChangedLineCount() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChangedLineCount
    ChangedLineCount = mvarChangedLineCount
End Property




Public Sub ProcessDiff(ByVal vsDiff As String)
    
    On Error GoTo ERR_ProcessDiff
    
    mvarBlankLineCount = 0
    mvarChangedLineCount = 0
    mvarCommentedLineCount = 0
    mvarDeletedLineCount = 0
    mvarIgnoredLineCount = 0
    mvarInsertedLineCount = 0
    mvarTotalLineCount = 0
    
    Dim oclLines As Collection
    
    Set oclLines = MakeLines(vsDiff)
    
    Const sDIFF_CHAR_CHANGED_LINE As String = "|"
    Const sDIFF_CHAR_DELETED_LINE As String = "<"
    Const sDIFF_CHAR_INSERTED_LINE As String = ">"
    
    'Go thr' each line to check diff indicator
    Dim lLineIndex As Long
    Dim sLine As String
    Dim sDiffChar As String
    
    mbIsFileExist = True
    If oclLines.Count > 1 Then
        If oclLines(1) = "File or project not found" Then
            mbIsFileExist = False
        End If
    End If
    
    If mbIsFileExist Then
        For lLineIndex = 1 To oclLines.Count
            sLine = oclLines(lLineIndex)
            If Not IsIgnorable(sLine) Then
                sDiffChar = Mid(sLine, lDIFF_INDICATOR_POS, 1)
                Select Case sDiffChar
                    Case sDIFF_CHAR_CHANGED_LINE
                        mvarChangedLineCount = mvarChangedLineCount + 1
                    Case sDIFF_CHAR_DELETED_LINE
                        mvarDeletedLineCount = mvarDeletedLineCount + 1
                    Case sDIFF_CHAR_INSERTED_LINE
                        mvarInsertedLineCount = mvarInsertedLineCount + 1
                    Case Else
                        mvarIgnoredLineCount = mvarIgnoredLineCount + 1
                End Select
            Else
                mvarIgnoredLineCount = mvarIgnoredLineCount + 1
            End If
        Next lLineIndex
        
        'First 4 line, if existed should not be counted
        If mvarIgnoredLineCount >= 4 Then
            mvarIgnoredLineCount = mvarIgnoredLineCount - 4
        End If
    End If
    
Exit Sub
ERR_ProcessDiff:
    Set oclLines = Nothing
    ReRaiseError
End Sub

Private Function IsIgnorable(ByVal vsLine As String) As Boolean
    Dim sLeftPart As String
    Dim sRightPart As String
    
    sLeftPart = Trim(Mid(vsLine, 1, lDIFF_INDICATOR_POS - 1))
    sRightPart = Trim(Mid(vsLine, lDIFF_INDICATOR_POS + 1))
    
    If ((sLeftPart = "") Or (Left(sLeftPart, 1) = "'")) And ((sRightPart = "") Or (Left(sRightPart, 1) = "'")) Then
        IsIgnorable = True
    Else
        IsIgnorable = False
    End If
    
End Function

Private Function MakeLines(ByVal vsString As String) As Collection
    Dim oclLines As Collection
    Set oclLines = New Collection
    
    Const sNEW_LINE As String = vbCrLf
    Dim lNewLinePos As Long
    Dim lLastNewLinePos As Long
    Dim lLineLen As Long
    Dim lNewLineConstLen As Long
    Dim bExitLoop As Boolean
    
    lNewLineConstLen = Len(sNEW_LINE)
    lLastNewLinePos = 1
    bExitLoop = False
    vsString = sNEW_LINE & vsString
    
    Do
        lNewLinePos = InStr(lLastNewLinePos + lNewLineConstLen, vsString, sNEW_LINE)
        If lNewLinePos <> 0 Then
            lLineLen = lNewLinePos - (lLastNewLinePos + lNewLineConstLen)
            If lLineLen > 0 Then
                oclLines.Add Mid(vsString, lLastNewLinePos + lNewLineConstLen, lLineLen)
            End If
            lLastNewLinePos = lNewLinePos
        Else
            oclLines.Add Mid(vsString, lLastNewLinePos + lNewLineConstLen - 1)
            bExitLoop = True
        End If
    Loop While Not bExitLoop
    Set MakeLines = oclLines
End Function

Public Function IsFileExist() As Boolean
    IsFileExist = mbIsFileExist
End Function


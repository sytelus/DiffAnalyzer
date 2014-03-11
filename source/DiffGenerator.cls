VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiffGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable(s) to hold property value(s)
Private mvarTempDir As String 'local copy
Private mvarTempFileName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarSourceSafePath As String 'local copy

Const msDEFAULT_TEMP_FILE As String = "DiffGenerator.txt"

Public Property Let SourceSafePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceSafePath = 5
    mvarSourceSafePath = vData
End Property


Public Property Get SourceSafePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceSafePath
    SourceSafePath = mvarSourceSafePath
End Property



Public Property Let TempFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempFileName = 5
    mvarTempFileName = vData
End Property


Public Property Get TempFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempFileName
    TempFileName = mvarTempFileName
End Property

Public Property Let TempDir(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempDir = 5
    mvarTempDir = vData
End Property


Public Property Get TempDir() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempDir
    TempDir = mvarTempDir
End Property

Public Function GetDiff(ByVal vsFile1 As String, ByVal vsFile2 As String) As String
    
    Dim sDiffFile As String
    
    sDiffFile = MakeDiffFile(vsFile1, vsFile2)
    
    'Read the file in string
    Dim iFileHandle As Integer
    Dim sFileContent As String
    Dim sLine As String
    sFileContent = vbNullString
    iFileHandle = FreeFile
    sFileContent = Space$(FileLen(sDiffFile))
    Open sDiffFile For Binary Access Read As #iFileHandle
            Get #iFileHandle, , sFileContent ' Read line into variable.
            'sFileContent = sFileContent & sLine
    Close #iFileHandle
    Call DeleteFile(sDiffFile)
    GetDiff = sFileContent
End Function

Private Sub Class_Initialize()
    mvarTempFileName = msDEFAULT_TEMP_FILE
End Sub

Public Function MakeDiffFile(ByVal vsFile1 As String, ByVal vsFile2 As String, Optional ByVal vsDiffFile As String = vbNullString) As String

    'Build SourceSafe command line
    'Build ss.EXE path
    Dim sSSEXE As String
    Const sSSEXE_SUB_DIR As String = "win32"
    Const sSSEXE_NAME As String = "ss.exe"
    
    
    sSSEXE = GetPathWithSlash(GetPathWithSlash(mvarSourceSafePath) & sSSEXE_SUB_DIR) & sSSEXE_NAME
    
    'Build Output file path
    Dim sOutFile As String
    
    If vsDiffFile = vbNullString Then
        sOutFile = GetPathWithSlash(mvarTempDir) & mvarTempFileName
    Else
        sOutFile = GetPathWithSlash(vsDiffFile)
    End If
    
    Call DeleteFile(sOutFile)
    
    'Concate to make command line
    Const sQUOTE As String = """"
    Const sSPACE As String = " "
    Const sDIFF_COMMAND As String = "diff"
    Const sDIFF_OPTIONS As String = "-DX0 -IWE"
    Const sDIFF_REDIRECT_OPTION As String = "-O@"
    
    Dim sCommandLine As String
    sCommandLine = sQUOTE & sSSEXE & sQUOTE & sSPACE & sDIFF_COMMAND & sSPACE
    sCommandLine = sCommandLine & sQUOTE & vsFile1 & sQUOTE & sSPACE
    sCommandLine = sCommandLine & sQUOTE & vsFile2 & sQUOTE & sSPACE
    sCommandLine = sCommandLine & sDIFF_OPTIONS & sSPACE
    sCommandLine = sCommandLine & sDIFF_REDIRECT_OPTION & sQUOTE & sOutFile & sQUOTE & sSPACE
    'Run the command

    Dim oExecAndWait As ExecAndWait
    Set oExecAndWait = New ExecAndWait
        oExecAndWait.WindowState = vbMinimized
         Call oExecAndWait.ExecAndWait(sCommandLine, False)
    Set oExecAndWait = Nothing

    MakeDiffFile = sOutFile

End Function

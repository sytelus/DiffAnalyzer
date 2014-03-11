VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecAndWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SW_MINIMIZE = 6
Private Const SW_MAXIMIZE = 3

Private Const STARTF_USESHOWWINDOW = &H1
'local variable(s) to hold property value(s)
Private mvarWindowState As FormWindowStateConstants     'local copy

Private mbIsStillRunning As Boolean

Public Event ProcessEnded()
Public Event DoYourTask()
Public Event Error(ByVal vlErrorNumber As Long, ByVal vsErrorDescription As String, ByVal vsErrorSource As String, ByRef rbCancelErrorRaise As Boolean)

Public Property Let WindowState(ByVal vData As FormWindowStateConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WindowState = 5
    mvarWindowState = vData
End Property


Public Property Get WindowState() As FormWindowStateConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WindowState
    WindowState = mvarWindowState
End Property

Public Sub ExecAndWait(ByVal CmdLine As String, Optional ByVal vboolUseDoEvents As Boolean = False)
    
    On Error GoTo ERR_ExecAndWait
    
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ReturnValue As Integer
    
    Const mbRETURN_PROCESS_ENDED As Integer = 258
    
    start.dwFlags = start.dwFlags Or STARTF_USESHOWWINDOW
    Select Case mvarWindowState
        Case FormWindowStateConstants.vbMaximized
            start.wShowWindow = SW_MAXIMIZE
        Case FormWindowStateConstants.vbMinimized
            start.wShowWindow = SW_MINIMIZE
        Case FormWindowStateConstants.vbNormal
            'Do nothing
    End Select
        
    mbIsStillRunning = True
    
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    ' Start the shelled application:
    ReturnValue = CreateProcessA(0&, CmdLine$, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
       
    If ReturnValue = 0 Then
        Dim lAPIError As Long
        Dim sErrorMessage As String
        lAPIError = GetLastError
        sErrorMessage = MessageText(lAPIError)
        Err.Raise lAPIError, , "Failed to create the process " & CmdLine & " : " & sErrorMessage
    End If
    ' Wait for the shelled application to finish:
    Do
       ReturnValue = WaitForSingleObject(proc.hProcess, 0)
       If vboolUseDoEvents Then
        DoEvents
       End If
       RaiseEvent DoYourTask
       Loop Until ReturnValue <> 258
    ReturnValue = CloseHandle(proc.hProcess)
    mbIsStillRunning = False
    RaiseEvent ProcessEnded
    
Exit Sub
ERR_ExecAndWait:
    mbIsStillRunning = False
    Dim bCancelErrorRaise As Boolean
    bCancelErrorRaise = False
    RaiseEvent Error(Err.Number, Err.Description, Err.Source, bCancelErrorRaise)
    If Not bCancelErrorRaise Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Sub Class_Initialize()
    mbIsStillRunning = False
End Sub

Public Function IsStillRunning() As Boolean
    IsStillRunning = mbIsStillRunning
End Function

Private Function MessageText(lCode As Long) As String
       Dim sRtrnCode As String
       Dim lRet As Long
       sRtrnCode = Space$(256)
       lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lCode, 0&, _
                 sRtrnCode, 256&, 0&)
        If lRet > 0 Then
          MessageText = Left(sRtrnCode, lRet)
        Else
          MessageText = "Error not found."
        End If
End Function


Attribute VB_Name = "mDuk"
Public hDukLib As Long
Public libRefCnt As Long 'used when running in IDE...

Public Declare Function DukCreate Lib "Duk4VB.dll" () As Long
Public Declare Function AddFile Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal jsFile As String) As Long
Public Declare Function Eval Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal js As String) As Long
Public Declare Function DukPushNewJSClass Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal jsClassName As String, ByVal hInst As Long) As Long 'returns 0/-1
Public Declare Sub SetCallBacks Lib "Duk4VB.dll" (ByVal msgProc As Long, ByVal dbgCmdProc As Long, ByVal hostResolverProc As Long, ByVal lineInputfunc As Long, ByVal debugWritefunc As Long)
Public Declare Function DukOp Lib "Duk4VB.dll" (ByVal operation As opDuk, Optional ByVal ctx As Long = 0, Optional ByVal arg1 As Long, Optional ByVal sArg As String) As Long

'misc windows api..
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long

Enum cb_type
    cb_output = 0
    cb_Refresh = 1
    cb_Fatal = 2
    'cb_engine = 3
    cb_error = 4
    cb_ReleaseObj = 5
    cb_StringReturn = 6
    cb_debugger = 7
End Enum

Enum opDuk
    opd_PushUndef = 0
    opd_PushNum = 1
    opd_PushStr = 2
    opd_GetInt = 3
    opd_IsNullUndef = 4
    opd_GetString = 5
    opd_Destroy = 6
    opd_LastString = 7
    opd_ScriptTimeout = 8
    opd_debugAttach = 9
    opd_dbgCoOp = 10
End Enum

Enum Debug_Commands
    dc_NotSet = 0
    dc_run = 1
    dc_stepInto = 3
    dc_Stepout = 4
    dc_StepOver = 5
    dc_RunToLine = 6
    dc_Quit = 7
    dc_Manual = 8
End Enum

Global Const STATUS_NOTIFICATION = &H1
Global Const PRINT_NOTIFICATION = &H2
Global Const ALERT_NOTIFICATION = &H3
Global Const LOG_NOTIFICATION = &H4
Global Const BASIC_INFO_REQ = &H10
Global Const TRIGGER_STATUS_REQ = &H11
Global Const PAUSE_REQ = &H12
Global Const RESUME_REQ = &H13
Global Const STEP_INTO_REQ = &H14
Global Const STEP_OVER_REQ = &H15
Global Const STEP_OUT_REQ = &H16
Global Const LIST_BREAK_REQ = &H17
Global Const ADD_BREAK_REQ = &H18
Global Const DEL_BREAK_REQ = &H19
Global Const GET_VAR_REQ = &H1A
Global Const PUT_VAR_REQ = &H1B
Global Const GET_CALL_STACK_REQ = &H1C
Global Const GET_LOCALS_REQ = &H1D
Global Const EVAL_REQ = &H1E
Global Const DETACH_REQ = &H1F
Global Const DUMP_HEAP_REQ = &H20
    
Global Const DUK_DBG_MARKER_EOM = 0
Global Const DUK_DBG_MARKER_REQUEST = 1
Global Const DUK_DBG_MARKER_REPLY = 2
Global Const DUK_DBG_MARKER_ERROR = 3
Global Const DUK_DBG_MARKER_NOTIFY = 4

Global running As Boolean
Public LastStringReturn As String
Public readyToReturn As Boolean
Public ActiveDebuggerClass As CDukTape
Public dbg_cmd As Debug_Commands
Private dbg_response() As Byte
Private dbgBufOffset As Long

Dim mResponseBuffer As New CResponseBuffer

Function InitDukLib(Optional ByVal explicitPathToDll As String) As Boolean

    If Len(explicitPathToDll) = 0 Then
        explicitPathToDll = App.path
        If Not FileExists(explicitPathToDll & "\duk4vb.dll") Then
            explicitPathToDll = GetParentFolder(explicitPathToDll)
            If Not FileExists(explicitPathToDll & "\duk4vb.dll") Then
                explicitPathToDll = GetParentFolder(explicitPathToDll)
                If Not FileExists(explicitPathToDll & "\duk4vb.dll") Then explicitPathToDll = GetParentFolder(explicitPathToDll)
            End If
        End If
        If Not FileExists(explicitPathToDll & "\duk4vb.dll") Then
            explicitPathToDll = Empty
        Else
            explicitPathToDll = explicitPathToDll & "\duk4vb.dll"
        End If
    End If

    If FileExists(explicitPathToDll) Then
        hDukLib = LoadLibrary(explicitPathToDll) 'to ensure the ide finds the dll
        If hDukLib = 0 Then Exit Function
    End If
    
    'this can still work..but now its up to the runtime to find the dll..if not the app will terminate
    SetCallBacks AddressOf vb_stdout, _
                 AddressOf GetDebuggerCommand, _
                 AddressOf HostResolver, _
                 AddressOf VbLineInput, _
                 AddressOf DebugDataIncoming
                 
    InitDukLib = True
    
End Function

'this is used for script to host app object integration..
Public Function HostResolver(ByVal buf As Long, ByVal ctx As Long, ByVal argCnt As Long, ByVal hInst As Long) As Long
    Dim key As String
    Dim v1 As Variant
    
    On Error Resume Next
    'we could switch to numeric ids..but it would be harder to manage/debug when more complex..
    key = StringFromPointer(buf)
    
    'this is just a quick demo not the full setup see duk4vb project for a full COM relay using same structure
    If key = "list1.additem" Then
        If argCnt > 1 Then
            v1 = GetArgAsString(ctx, i + 3)
            Form1.List1.AddItem CStr(v1)
        End If
    End If
    
'    If key = "text2.text" Then
'        DukOp opd_PushStr, ctx, 0, Form1.Text2.text
'        HostResolver = 1
'    End If
            
End Function
    
Public Sub DebuggerCmd(cmd As Debug_Commands)
    
    Dim startPos As Long, endPos As Long
    
    Debug.Print String(70, "-")
    
'    With frmMain
'        .scivb.DeleteMarker .lastEIP, 1 'remove the yellow arrow
'        .scivb.DeleteMarker .lastEIP, 3 'remove the yellow line backcolor
'
'        'force a refresh of the specified line or it might not catch it..
'        startPos = .scivb.PositionFromLine(.lastEIP)
'        endPos = .scivb.PositionFromLine(.lastEIP + 1)
'        .scivb.DirectSCI.Colourise startPos, endPos
'
'    End With

    If Not mResponseBuffer.ConstructMessage(cmd) Then
        Debug.Print "Failed to construct message for " & cmd
    Else
        readyToReturn = True
    End If
    
End Sub

Function GetLastString() As String
    Dim rv As Long
    rv = DukOp(opd_LastString)
    If rv = 0 Then Exit Function
    GetLastString = StringFromPointer(rv)
End Function



Function GetArgAsString(ctx As Long, index As Long) As String
    
    'an invalid index here would trigger a script error and aborting the eval call..weird.. <---
    'as long as the native function is added with expected arg count, and you dont surpass it your ok
    'even if the js function ommitted args in its call, empty ones will just be retrieved as 'undefined'
    
    Dim ptr As Long
    ptr = DukOp(opd_GetString, ctx, index)
    
    If ptr <> 0 Then
        GetArgAsString = StringFromPointer(ptr)
    End If
    
End Function
 




'callback functions
'------------------------------
Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long)

    Dim msg As String
    
    If t = cb_Fatal Then
    
        MsgBox "A fatal error has occured in Duktape the application " & vbCrLf & _
               "is unstable now please save your work and exit." & vbCrLf & vbCrLf & _
               "The specific error message was: " & StringFromPointer(lpMsg), vbCritical, "Fatal Error"
        
        While Forms.Count > 0
            DoEvents
            Sleep 10
        Wend
        
        FreeLibrary hDukLib
        End
        
    End If
    
    If t = cb_Refresh Then
        DoEvents
        Sleep 3
        Exit Sub
    End If
    
    If lpMsg = 0 Then Exit Sub
    
    msg = StringFromPointer(lpMsg)
    
    Select Case t
        Case cb_StringReturn: LastStringReturn = msg
        'Case cb_ReleaseObj: ReleaseObj CLng(msg)
        Case cb_output, cb_error:  MsgBox msg, vbInformation, "Script Output"
    End Select
    
End Sub


Public Function VbLineInput(ByVal buf As Long, ByVal ctx As Long) As Long
    Dim b() As Byte
    Dim retVal As String
    VbLineInput = 0 'return value default..
    
    Dim text As String
    Dim def As String
    
    text = StringFromPointer(buf)
    def = GetArgAsString(ctx, 1)
    
    retVal = InputBox(text, "Script Basic Line Input", def)
    
    If Len(retVal) = 0 Then
        DukOp opd_PushUndef, ctx
        Exit Function
    Else
        DukOp opd_PushStr, ctx, 0, retVal
    End If
        
  
End Function

'debugger is requesting a command to operate on..vb blocks until user enters command..
Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
    
    Dim i As Long
    Dim cmd_length As Long
    Dim b() As Byte
    
        'frmMain.SyncUI
        
        If Not mResponseBuffer.isEmpty Then
            
            If mResponseBuffer.GetBuf(sz, b) Then
                CopyMemory ByVal buf, ByVal VarPtr(b(0)), sz
                GetDebuggerCommand = sz
            End If
            
            Exit Function
        End If
        
        
        'we block here until the UI sets the readyToReturn = true
        'this is not a CPU hog, and form remains responsive to user actions..
        readyToReturn = False
        While Not readyToReturn
            DoEvents
            Sleep 20
            i = i + 1
            If i = 500 Then
                If Not ActiveDebuggerClass Is Nothing Then
                    DukOp opd_dbgCoOp, ActiveDebuggerClass.Context
                End If
                i = 0
            End If
        Wend
        
        If Not mResponseBuffer.isEmpty Then
            
            If mResponseBuffer.GetBuf(sz, b) Then
                CopyMemory ByVal buf, ByVal VarPtr(b(0)), sz
                GetDebuggerCommand = sz
            End If
            
            Exit Function
        End If
        
        
End Function

'debugger is sending our interface data
Public Function DebugDataIncoming(ByVal buf As Long, ByVal sz As Long) As Long

    If buf = 0 Or sz = 0 Then Exit Function
    
    Dim b() As Byte
    ReDim b(sz)
    CopyMemory b(0), ByVal buf, sz
    Debug.Print bHexDump(b)
    
    DebugDataIncoming = sz
End Function

'example data coming in from start debug session setup..
'000000   31 20 31 30 32 30 31 20 76 31 2E 32 2E 31 20 75    1.10201.v1.2.1.u
'000010   6E 6B 6E 6F 77 6E 0A 00                            nknown..                                             ..
'000000   04 00                                              ..
'000000   81 00                                              ..
'000000   76 00                                              v.
'000000   64 3A 5C 5F 63 6F 64 65 5C 64 75 6B 34 76 62 5C    d:\_code\duk4vb\
'000010   6D 61 69 6E 2E 63 00                               main.c.
'000000   64 00                                              d.
'000000   65 76 61 6C 00                                     eval.
'000000   81 00                                              ..
'000000   80 00                                              ..
'000000   00 00                                              ..









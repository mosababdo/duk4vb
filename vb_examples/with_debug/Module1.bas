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
    opd_dbgManuallyTriggerGetVar = 11
End Enum

Enum Debug_Commands
    dc_NotSet = 0
    dc_break = 2
    dc_stepInto = 3
    dc_stepout = 4
    dc_StepOver = 5
    'dc_RunToLine = 6
    'dc_Quit = 7
    'dc_Manual = 8
    dc_Resume = 9
    dc_GetLocals = 10
    dc_GetVar = 11
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

Global Const DUK_DBG_CMD_STATUS = &H1
Global Const DUK_DBG_CMD_PRINT = &H2
Global Const DUK_DBG_CMD_ALERT = &H3
Global Const DUK_DBG_CMD_LOG = &H4

Global Const DUK_DBG_ERR_UNKNOWN = &H0
Global Const DUK_DBG_ERR_UNSUPPORTED = &H1
Global Const DUK_DBG_ERR_TOOMANY = &H2
Global Const DUK_DBG_ERR_NOTFOUND = &H3
Global Const DUK_VAR_NOT_FOUND = "DUK_VAR_NOT_FOUND"

Global running As Boolean
Public LastStringReturn As String
Public readyToReturn As Boolean
Public LastCommand As Debug_Commands
Dim varsLoaded As Boolean
Public replyReceived As Boolean

Public ActiveDebuggerClass As CDukTape
Public RespBuffer As New CResponseBuffer
Public RecvBuffer As New CWriteBuffer

Private variables As Collection
Private dbgStopNext As Boolean

Public Type stats
    state As Long
    filename As String
    curFunc As String
    lineNumber As Long
    pc As Long
End Type

Public status As stats
Public VarReturn As CVariable

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
'    If key = "list1.additem" Then
'        If argCnt > 1 Then
'            v1 = GetArgAsString(ctx, i + 3)
'            Form1.List1.AddItem CStr(v1)
'        End If
'    End If
    
'    If key = "text2.text" Then
'        DukOp opd_PushStr, ctx, 0, Form1.Text2.text
'        HostResolver = 1
'    End If
            
End Function

Sub RefreshVariables()
    
    Dim li As ListItem
    Dim v As CVariable
    
'    Set variables = New Collection
'
'    Form1.lvVars.ListItems.Clear
'    DebuggerCmd dc_GetLocals
'
'    For Each v In variables
'        Set li = Form1.lvVars.ListItems.Add(, , IIf(v.isGlobal, "Global", "Local"))
'        li.SubItems(1) = v.name
'        li.SubItems(2) = v.varType
'        li.SubItems(3) = v.Value
'        Set li.Tag = v
'    Next
    
End Sub

'this is messed up but it works...see notes in DukOp for opd_dbgManuallyTriggerGetVar
Function SyncronousGetVariableValue(name As String) As CVariable
    Set VarReturn = New CVariable
    VarReturn.name = name
    LastCommand = dc_GetVar
    replyReceived = False
    RespBuffer.ConstructMessage dc_GetVar, name, True   'build custom packet
    DukOp opd_dbgManuallyTriggerGetVar, ActiveDebuggerClass.Context
    Set SyncronousGetVariableValue = VarReturn
End Function

Public Sub DebuggerCmd(cmd As Debug_Commands, Optional arg1)
    
    If Not replyReceived Then Exit Sub ' we are still waiting for last commands response..
    replyReceived = False
    
    LastCommand = cmd
    If Not RespBuffer.ConstructMessage(cmd, arg1) Then
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



Function GetArgAsString(ctx As Long, Index As Long) As String
    
    'an invalid index here would trigger a script error and aborting the eval call..weird.. <---
    'as long as the native function is added with expected arg count, and you dont surpass it your ok
    'even if the js function ommitted args in its call, empty ones will just be retrieved as 'undefined'
    
    Dim ptr As Long
    ptr = DukOp(opd_GetString, ctx, Index)
    
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
        
        While Forms.count > 0
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
        Case cb_debugger:
                If msg = "Debugger-Detached" Then running = False
                
    End Select
    
End Sub


Public Function VbLineInput(ByVal buf As Long, ByVal ctx As Long) As Long
    Dim b() As Byte
    Dim retVal As String
    VbLineInput = 0 'return value default..
    
    Dim Text As String
    Dim def As String
    
    Text = StringFromPointer(buf)
    def = GetArgAsString(ctx, 1)
    
    retVal = InputBox(Text, "Script Basic Line Input", def)
    
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
        
topLine:
        If Not RespBuffer.isEmpty Then
            
            If RespBuffer.GetBuf(sz, b) Then
                CopyMemory ByVal buf, ByVal VarPtr(b(0)), sz
                GetDebuggerCommand = sz
            End If
            
            Exit Function
        End If
        
        If Not varsLoaded Then
            varsLoaded = True
            Set variables = New Collection
            Form1.lvVars.ListItems.Clear
            dbgStopNext = True
            DebuggerCmd dc_GetLocals
            GoTo topLine 'immediate send of response buffer..
        End If
        
        'we block here until the UI sets the readyToReturn = true
        'this is not a CPU hog, and form remains responsive to user actions..
        readyToReturn = False
        While Not readyToReturn
            DoEvents
            Sleep 20
            i = i + 1
            
            If running = False Then 'we have a detach
                Exit Function
            End If
            
            If i = 500 Then
                If Not ActiveDebuggerClass Is Nothing Then
                    DukOp opd_dbgCoOp, ActiveDebuggerClass.Context
                End If
                i = 0
            End If
            
        Wend
        
        If Not RespBuffer.isEmpty Then
            
            If RespBuffer.GetBuf(sz, b) Then
                CopyMemory ByVal buf, ByVal VarPtr(b(0)), sz
                GetDebuggerCommand = sz
            End If
            
            Exit Function
        End If
        
        
End Function

'debugger is sending our interface data, this happens in multiple stages until a single EOM byte is received (00)
Public Function DebugDataIncoming(ByVal buf As Long, ByVal sz As Long) As Long

    If buf = 0 Or sz = 0 Then Exit Function 'shouldnt happen...
    
    Dim b() As Byte
    ReDim b(sz - 1) 'b is 0 based,
    CopyMemory b(0), ByVal buf, sz
    'Debug.Print bHexDump(b)
     
    RecvBuffer.WriteBuf b()
    DebugDataIncoming = sz
    
   
    
End Function

'called by RecvBuff when a full message has been received..
Public Function DebuggerMessageReceived()
    
    Dim b As Byte
    Dim i As Long
    
    'If dbgStopNext Then Stop
        
    With RecvBuffer
        
        b = .ReadByte()
        
        If .firstMessage Then
            If b <> Asc(1) Then MsgBox "Bad debugger protocol version!"
            '<protocolversion> <SP (0x20)> <additional text, no LF> <LF (0x0a)>
            '1 <DUK_VERSION> <DUK_GIT_DESCRIBE> <target string> <LF>
            '1 10099 v1.0.0-254-g2459e88 duk command built from Duktape repo
            'todo: HandleInitMessage
            Exit Function
        End If
        
        Select Case b
            Case DUK_DBG_MARKER_NOTIFY: HandleNotify
            
            Case DUK_DBG_MARKER_REQUEST: .DebugDump ("Request")
            
            'ERR <int: error code> <str: error message or empty string> EOM
            '0x00    Unknown or unspecified error
            '0x01    Unsupported command
            '0x02    Too many (e.g. too many breakpoints, cannot add new)
            '0x03    Not found (e.g. invalid breakpoint index)
            Case DUK_DBG_MARKER_ERROR: .DebugDump ("Error")
            
            Case DUK_DBG_MARKER_REPLY:
            
                    If .BytesLeft = 0 Then
                        'Debug.Print "Success Reply"
                        Exit Function
                    End If
                    
                    'the reply is specific to the last command we issued..
                    Select Case LastCommand
                        Case dc_GetVar: HandleGetVar
                        'Case dc_GetLocals:
                        Case Else:
                            .DebugDump ("Reply last cmd: " & LastCommand)
                    End Select
                    
        End Select
        
    End With
    
End Function
                                            
                                            
Function HandleNotify()
    Dim fname As String, func As String
    Dim msg As Long, state As Long, lno As Long, pc As Long
    
    With RecvBuffer
    
        If .BytesLeft < 4 Then Exit Function '??
        msg = .ReadInt
        
        Select Case msg
                   'NFY <int: 1> <int: state> <str: filename> <str: funcname> <int: linenumber> <int: pc> EOM
              Case STATUS_NOTIFICATION:
                    status.state = .ReadInt
                    status.filename = .ReadString
                    status.curFunc = .ReadString
                    status.lineNumber = .ReadInt
                    status.pc = .ReadInt
                    'Debug.Print "File: " & fname & " func: " & func & " Line: " & lno
                    'we cant sync the UI until we get this message to show our line number were on...
                    varsLoaded = False
                    Form1.SyncUI
                    
                   'NFY <int: 2> <str: message> EOM - String output redirected from the print() function.
              Case PRINT_NOTIFICATION: .DebugDump "Print Notify"
              
                   'NFY <int: 3> <str: message> EOM - String output redirected from the alert() function.
              Case ALERT_NOTIFICATION: .DebugDump "Alert Notify"
              
                   'NFY <int: 4> <int: log level> <str: message> EOM - Logger output redirected from Duktape logger calls.
              Case LOG_NOTIFICATION:   .DebugDump "Log Notify"
        End Select
        
    End With
    
    
End Function


Function HandleGetVar()
    
    Dim b As Byte
    Dim found As Long
    
    'REP <int: 0/1, found> <tval: value> EOM
    found = RecvBuffer.ReadInt
    If found = 0 Then
        VarReturn.varType = DUK_VAR_NOT_FOUND
        Exit Function
    End If
        
    With VarReturn
        b = RecvBuffer.ReadByte
        Select Case b
                
            Case &H10: '0x10 <int32>    integer     4-byte integer, signed 32-bit integer in network order follows initial byte
                    .varType = "integer"
                    .Value = RecvBuffer.ReadInt(False)
            
            Case &H11: '0x11 <uint32> <data>    string  4-byte string, unsigned 32-bit string length in network order and string data follows initial byte
                    .varType = "string"
                    .Value = RecvBuffer.ReadString(False)
            
            Case &H16: .varType = "undefined"                 '0x16    undefined   Ecmascript "undefined"
            Case &H17: .varType = "null"                      '0x17    null    Ecmascript "null"
            Case &H18: .varType = "boolean": .Value = True    '0x18    true    Ecmascript "true"
            Case &H19:  .varType = "boolean": .Value = False  '0x19    false   Ecmascript "false"
            Case &H1A: .varType = "double":                   '0x1a <8 bytes>  number  IEEE double (network endian)
                       .Value = RecvBuffer.ReadDouble()
                       
            '0x1b <uint8> <uint8> <data>     object  Class number, pointer length, and pointer data (network endian)
            Case &H1B: .varType = "object": .Value = "[Object]"
            
            '0x1c <uint8> <data>     pointer     Pointer length, pointer data (network endian)
            '0x1d <uint16> <uint8> <data>    lightfunc   Lightfunc flags, pointer length, pointer data (network endian)
            '0x1e <uint8> <data>     heapptr     Pointer to a heap object (used by DumpHeap, network endian)
            
            '0x13 <uint32> <data>    buffer  4-byte buffer, unsigned 32-bit buffer length in network order and buffer data follows initial byte
            '0x14 <uint16> <data>    buffer  2-byte buffer, unsigned 16-bit buffer length in network order and buffer data follows initial byte
            
            
            'these ones we dont have to worry about..(not using packed protocol)
                '0x12 <uint16> <data>    string  2-byte string, unsigned 16-bit string length in network order and string data follows initial byte
                '0x60...0x7f <data>  string  String with length [0,31], string length is IB - 0x60, data follows
                '0x80...0xbf     integer     Integer [0,63], integer value is IB - 0x80
                '0xc0...0xff <uint8>     integer     Integer [0,16383], integer value is ((IB - 0xc0) << 8) + followup_byte
            
            'unused:
                '0x15    unused  Represents the internal "undefined unused" type which used to e.g. mark unused (unmapped) array entries
                '0x1f reserved
                '0x20...0x5f     reserved
            
            Case Else: .varType = "VarType: 0x" & Hex(b)
        End Select
        
    End With

End Function






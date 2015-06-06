Attribute VB_Name = "Module1"

Public Declare Sub DukDestroy Lib "Duk4VB.dll" ()
Public Declare Sub DukCreate Lib "Duk4VB.dll" ()
Public Declare Function AddFile Lib "Duk4VB.dll" (ByVal jsFile As String) As Long
Public Declare Function Eval Lib "Duk4VB.dll" (ByVal js As String) As Long
Public Declare Function LastString Lib "Duk4VB.dll" (ByVal buf As String, ByVal bufSz As Long) As Long

'safe from invalid indexes..invalid context will crash
Public Declare Function DukGetInt Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal index As Long) As Long
Public Declare Function DukGetString Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal index As Long) As Long 'returns string length..

Public Declare Sub DukPushNum Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal val As Long)
Public Declare Sub DukPushString Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal val As String)

Public Declare Function GetLastStringSize Lib "Duk4VB.dll" () As Long
Public Declare Sub SetCallBacks Lib "Duk4VB.dll" (ByVal msgProc As Long, ByVal dbgCmdProc As Long, ByVal hostResolverProc As Long, ByVal lineInputfunc As Long)

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Enum cb_type
    cb_output = 0
    'cb_dbgout = 1
    'cb_debugger = 2
    'cb_engine = 3
    cb_error = 4
    cb_refreshUI = 5
End Enum

Dim objs As New Collection

Function AddObject(o As Object, name As String) As Boolean
    On Error GoTo hell
    objs.Add o, name
    AddObject = True
    Exit Function
hell:
End Function

Function GetArgAsString(ctx As Long, index As Long) As String
    If DukGetString(ctx, index) > 0 Then GetArgAsString = GetLastString()
End Function
 
Function GetLastString() As String
    
    Dim rv As Long
    Dim tmp As String
    
    rv = GetLastStringSize()
    If rv < 0 Then Exit Function
    
    rv = rv + 2
    tmp = String(rv, " ")
    rv = LastString(tmp, rv)
    tmp = Mid(tmp, 1, rv)
    
    GetLastString = tmp
        
End Function

Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)

    Dim b() As Byte
    Dim msg As String
    
    'If shuttingDown Then Exit Sub
'
'    If t = cb_refreshUI Then
'        frmMain.Refresh
'        DoEvents
'        Sleep 10
'        Exit Sub
'    End If
    
    If lpMsg = 0 Or sz = 0 Then Exit Sub
    
    ReDim b(sz)
    CopyMemory b(0), ByVal lpMsg, sz
    msg = StrConv(b, vbUnicode)
    If Right(msg, 1) = Chr(0) Then msg = Left(msg, Len(msg) - 1)
    
    MsgBox msg
    
'    Select Case t
'        'Case cb_debugger: HandleDebugMessage msg
'        'Case cb_engine:   HandleEngineMessage msg
'        'Case cb_error:    ParseError msg
'        Case Else:
'
'                          If t = cb_dbgout Then msg = "DBG> " & msg
'
'                          With frmMain.txtOut
'                               .Text = .Text & Replace(msg, vbLf, vbCrLf)
'                               .Refresh
'                               DoEvents
'                          End With
'    End Select
    
End Sub

'this is used for script to host app object integration..
Public Function HostResolver(ByVal buf As Long, ByVal strlen As Long, ByVal ctx As Long, ByVal argCnt As Long) As Long


    Dim b() As Byte
    Dim name As String
    
    ReDim b(strlen)
    CopyMemory b(0), ByVal buf, strlen
    name = StrConv(b, vbUnicode)
    If Right(name, 1) = Chr(0) Then name = Left(name, Len(name) - 1)

    'MsgBox "Host resolver: " & name & " ctx:" & Hex(ctx) & " args: " & argCnt
    
    'MsgBox DukGetInt(ctx, 1)
 
    Dim rv As Long
    'rv = DukGetString(ctx, 2)
    'MsgBox "arg2: " & GetLastString()
    
    'Function OpenDialog(filt As FilterTypes, [initDir As String], [title As String], [pHwnd As Long]) As String
    '"call:cmndlg:OpenDialog:int:[string][string][int]:r_string"
    
    On Error Resume Next
    Dim o As Object, tmp, args(), retVal As Variant, i As Long
    
    tmp = Split(name, ":")
    Set o = objs(tmp(1))
    
    If o Is Nothing Then Exit Function
    
    If argCnt > 0 Then
        ReDim args(argCnt - 1)
        For i = 0 To argCnt - 1
            If InStr(1, tmp(i + 3), "string") > 0 Then
                args(i) = GetArgAsString(ctx, i + 1)
            ElseIf InStr(1, tmp(i + 3), "long") > 0 Then
                args(i) = DukGetInt(ctx, i + 1)
            End If
        Next
    End If
    
    Err.Clear
    'callbyname obj, method, type, args() as variant
    'retVal = CallByName(o, CStr(tmp(2)), VbMethod, args()) 'nope wont work this way.. :(
    
    Dim t As VbCallType
    If tmp(0) = "call" Then t = VbMethod
    If tmp(0) = "let" Then t = VbLet
    If tmp(0) = "get" Then t = VbGet
        
    If argCnt > 0 Then
        retVal = CallByNameEx(o, CStr(tmp(2)), t, args())
    Else
        retVal = CallByNameEx(o, CStr(tmp(2)), t)
    End If
    
    If InStr(1, tmp(UBound(tmp)), "string") > 0 Then
        DukPushString ctx, CStr(retVal)
    ElseIf InStr(1, tmp(UBound(tmp)), "long") > 0 Then
        DukPushNum ctx, CLng(retVal)
    End If
        
    
    
    'If Err.Number <> 0 Then MsgBox Err.Description Else MsgBox retVal
    
    
End Function

'http://www.vbforums.com/showthread.php?405366-RESOLVED-Using-CallByName-with-variable-number-of-arguments
Public Function CallByNameEx(Obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray As Variant)
    
        Dim oTLI As New TLIApplication
        Dim ProcID As Long
        Dim numArgs As Long
        Dim i As Long
        Dim v()
        
        On Error GoTo Handler
        
        'Set oTLI = CreateObject("TLI.TLIApplication")
        ProcID = oTLI.InvokeID(Obj, ProcName)
        
        If IsMissing(vArgsArray) Then
            CallByNameEx = oTLI.InvokeHook(Obj, ProcID, CallType)
        End If
        
        If IsArray(vArgsArray) Then
            numArgs = UBound(vArgsArray)
            ReDim v(numArgs)
            For i = 0 To numArgs
                v(i) = vArgsArray(numArgs - i)
            Next i
            CallByNameEx = oTLI.InvokeHookArray(Obj, ProcID, CallType, v)
        End If
        
    Exit Function
     
Handler:
        Debug.Print Err.Number, Err.Description
End Function
    
    


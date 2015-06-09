Attribute VB_Name = "mDuk"
Public hDukLib As Long
Public libRefCnt As Long 'used when running in IDE...

Public Declare Function DukCreate Lib "Duk4VB.dll" () As Long
Public Declare Function AddFile Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal jsFile As String) As Long
Public Declare Function Eval Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal js As String) As Long
Public Declare Function DukPushNewJSClass Lib "Duk4VB.dll" (ByVal ctx As Long, ByVal jsClassName As String, ByVal hInst As Long) As Long 'returns 0/-1
Public Declare Sub SetCallBacks Lib "Duk4VB.dll" (ByVal msgProc As Long, ByVal dbgCmdProc As Long, ByVal hostResolverProc As Long, ByVal lineInputfunc As Long)
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
End Enum


Dim objs As New Collection

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
    SetCallBacks AddressOf vb_stdout, 0, AddressOf HostResolver, AddressOf VbLineInput
    InitDukLib = True
    
End Function

Function GetLastString() As String
    Dim rv As Long
    rv = DukOp(opd_LastString)
    If rv = 0 Then Exit Function
    GetLastString = StringFromPointer(rv)
End Function

Function AddObject(o As Object, name As String) As Boolean
    On Error GoTo hell
    objs.Add o, name
    AddObject = True
    Exit Function
hell:
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
 
Function ReleaseObj(hInst As Long)
    On Error GoTo hell
    dbg "ReleaseObj: ", hInst
    Dim o As Object
    Set o = objs("obj:" & hInst)
    objs.Remove "obj:" & hInst
    Set o = Nothing
hell:
    If Err.Number <> 0 Then Debug.Print "Error in ReleaseObj(" & hInst & ")" & Err.Description
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
        Case cb_ReleaseObj: ReleaseObj CLng(msg)
        Case cb_output, cb_error:  MsgBox msg, vbInformation, "Script Output"
    End Select
    
End Sub


'this is used for script to host app object integration..
Public Function HostResolver(ByVal buf As Long, ByVal ctx As Long, ByVal argCnt As Long) As Long


    Dim b() As Byte
    Dim name As String
    
    name = StringFromPointer(buf)
    dbg "HostResolver: ", name, ctx, argCnt
    
    Dim rv As Long
    
    On Error Resume Next
    Dim o As Object, tmp, args(), retVal As Variant, i As Long, hInst As Long, oo As Object
    Dim firstUserArg As Long
    
    firstUserArg = 0
    tmp = Split(name, ":")
    If tmp(1) = "objptr" Then
        firstUserArg = 1
        hInst = DukOp(opd_GetInt, ctx, 2)
        For Each oo In objs
            If ObjPtr(oo) = hInst Then
                Set o = oo
                Exit For
            End If
        Next
    Else
        Set o = objs(tmp(1))
    End If
    
    If o Is Nothing Then
        dbg "Host resolver could not find object!"
        Exit Function
    End If
    
    If argCnt > 0 Then
        For i = firstUserArg To argCnt - 1
            'If DukIsNullOrUndef(ctx, i) = 1 Then
            '    Exit For
            'End If
            If InStr(1, tmp(i + 3), "string") > 0 Then
                 push args, GetArgAsString(ctx, i + 2)
            ElseIf InStr(1, tmp(i + 3), "long") > 0 Then
                push args, DukOp(opd_GetInt, ctx, i + 2)
            ElseIf InStr(1, tmp(i + 3), "bool") > 0 Then
                push args, CBool(GetArgAsString(ctx, i + 2))
            End If
        Next
        
        
    End If
    
    Err.Clear
    'callbyname obj, method, type, args() as variant
    'retVal = CallByName(o, CStr(tmp(2)), VbMethod, args()) 'nope wont work this way.. :(
    
    Dim t As VbCallType, isObj As Boolean
    
    If tmp(0) = "call" Then t = VbMethod
    If tmp(0) = "let" Then t = VbLet
    If tmp(0) = "get" Then t = VbGet
    If VBA.Left(tmp(UBound(tmp)), 5) = "r_obj" Then
        isObj = True
        tmp(UBound(tmp)) = Mid(tmp(UBound(tmp)), 6)
    End If
    
    If isObj Then
        Set retVal = CallByNameEx(o, CStr(tmp(2)), t, args(), isObj)
    Else
        retVal = CallByNameEx(o, CStr(tmp(2)), t, args(), isObj)
    End If
    
    HostResolver = 0 'are we setting a return value (doesnt seem to be critical)
    
    If InStr(1, tmp(UBound(tmp)), "string") > 0 Then
        dbg "returning string"
        DukOp opd_PushStr, ctx, 0, CStr(retVal)
        If t <> VbLet Then HostResolver = 1
    ElseIf InStr(1, tmp(UBound(tmp)), "long") > 0 Then
        dbg "returning long"
        DukOp opd_PushNum, ctx, CLng(retVal)
        If t <> VbLet Then HostResolver = 1
    End If
        
    If isObj Then
        dbg "returning new js class " & tmp(UBound(tmp))
        DukPushNewJSClass ctx, tmp(UBound(tmp)), ObjPtr(retVal)
        objs.Add retVal, "obj:" & ObjPtr(retVal)
        HostResolver = 1
    End If
    
    'If Err.Number <> 0 Then MsgBox Err.Description Else MsgBox retVal
    
    
End Function

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






'utility functions..
'--------------------------------------------------------------------


'http://www.vbforums.com/showthread.php?405366-RESOLVED-Using-CallByName-with-variable-number-of-arguments
Public Function CallByNameEx(Obj As Object, ProcName As String, CallType As VbCallType, Optional vArgsArray As Variant, Optional isObj As Boolean = False)
    
        Dim oTLI As New TLIApplication
        Dim ProcID As Long
        Dim numArgs As Long
        Dim i As Long
        Dim v()
        
        On Error GoTo Handler
        
        'Set oTLI = CreateObject("TLI.TLIApplication")
        ProcID = oTLI.InvokeID(Obj, ProcName)
        
        If Not IsArray(vArgsArray) Or AryIsEmpty(vArgsArray) Then
            dbg "CallByName: ", Obj, ProcName, isObj
            If isObj Then
                Set CallByNameEx = oTLI.InvokeHook(Obj, ProcID, CallType)
            Else
                CallByNameEx = oTLI.InvokeHook(Obj, ProcID, CallType)
            End If
        Else
            numArgs = UBound(vArgsArray)
            dbg "CallByName: ", Obj, ProcName, isObj, Join(vArgsArray, ", ")
            ReDim v(numArgs)
            For i = 0 To numArgs
                v(i) = vArgsArray(numArgs - i)
            Next i
            If isObj Then
                Set CallByNameEx = oTLI.InvokeHookArray(Obj, ProcID, CallType, v)
            Else
                CallByNameEx = oTLI.InvokeHookArray(Obj, ProcID, CallType, v)
            End If
        End If
        
    Exit Function
     
Handler:
        dbg "Error in CallByNameEx: ", Err.Number, Err.Description
End Function



Private Function StringFromPointer(buf As Long) As String
    Dim sz As Long
    Dim tmp As String
    Dim b() As Byte
    
    If buf = 0 Then Exit Function
       
    sz = lstrlen(buf)
    If sz = 0 Then Exit Function
    
    ReDim b(sz)
    CopyMemory b(0), ByVal buf, sz
    tmp = StrConv(b, vbUnicode)
    If Right(tmp, 1) = Chr(0) Then tmp = Left(tmp, Len(tmp) - 1)
    
    StringFromPointer = tmp
 
End Function

Sub dbg(prefix As String, ParamArray args())
    Dim a, tmp As String
    
    tmp = prefix
    
    For Each a In args
        If IsNumeric(a) Then
            tmp = tmp & Hex(a) & ", "
        ElseIf IsObject(a) Then
            tmp = tmp & "Obj: " & TypeName(a) & ", "
        Else
            tmp = tmp & a & ", "
        End If
    Next
    
    Form1.List1.AddItem tmp
    Debug.Print tmp
    
End Sub

Function GetParentFolder(path) As String
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function


Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

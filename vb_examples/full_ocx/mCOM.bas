Attribute VB_Name = "mCOM"
Option Explicit

'we should really move these into the CDuktape class so they are per held per instance..
Public comTypes As New Collection
Public objs As New Collection

'this we dont support..
'return types string() or arrays in general
'methods which require object arguments
'functions with more than 10 args

'todo: test property let/get , objrets


Function ParseObjectToCache(name As String, obj As Object, owner As CDukTape) As Boolean
    
    Dim cc As CCOMType
    
    If KeyExistsInCollection(comTypes, name) Then
        Set cc = comTypes(name)
        If cc.errors.count = 0 Then ParseObjectToCache = True
        Exit Function
    End If
        
    If Not obj Is Nothing Then objs.Add obj, name 'some types arent creatable/top level and are retvals
    
    Set cc = New CCOMType
    Set cc.owner = owner
    ParseObjectToCache = cc.LoadType(name)
    comTypes.Add cc, name
    
End Function

Function ReleaseObj(hInst As Long)
    On Error GoTo hell
    dbg "ReleaseObj: " & hInst
    Dim o As Object
    Set o = objs("obj:" & hInst)
    objs.Remove "obj:" & hInst
    Set o = Nothing
hell:
    If Err.Number <> 0 Then dbg "Error in ReleaseObj(" & hInst & ")" & Err.Description
End Function

Sub ResetComObjects()

    Dim o As Object
    For Each o In comTypes
        Set o = Nothing
    Next
    
    For Each o In objs
        Set o = Nothing
    Next
    
    Set comTypes = New Collection
    Set objs = New Collection
    
End Sub

'this is used for script to host app object integration..
Public Function cb_HostResolver(ByVal buf As Long, ByVal ctx As Long, ByVal argCnt As Long, ByVal hInst As Long) As Long
    
    Dim o As Object, tmp, args(), retVal As Variant, i As Long, oo As Object
    Dim firstUserArg As Long
    Dim rv As Long
    Dim b() As Byte
    Dim key As String
    Dim cc As CCOMType
    Dim meth As CMethod
    Dim pObjName As String
    Dim a As Long
    Dim t As VbCallType
    
    On Error Resume Next
    
    'this callback can be used by just the CDukTape class without the debugger..
    'If Not isControlActive() Then Exit Function
    
    key = StringFromPointer(buf)
    dbg "HostResolver: " & key & " ctx:" & ctx & " args: " & argCnt & " hInst: " & hInst

    a = InStr(key, ".") - 1
    If a > 0 Then pObjName = Mid(key, 1, a)
    
    If Not KeyExistsInCollection(comTypes, pObjName) Then Exit Function
    Set cc = comTypes(pObjName)
    
    If Not cc.GetMethod(key, meth) Then Exit Function
    
    If hInst <> 0 Then
        For Each oo In objs
            If ObjPtr(oo) = hInst Then
                Set o = oo
                Exit For
            End If
        Next
    Else
        Set o = objs(pObjName)
    End If
    
    If o Is Nothing Then
        dbg "Host resolver could not find object!"
        Exit Function
    End If
    
    'todo handle all arg types..
    If argCnt > 0 Then
        For i = 1 To argCnt
            If meth.ArgTypes(i) = "string" Then
                 push args, GetArgAsString(ctx, i + 2)
            ElseIf meth.ArgTypes(i) = "variant" Then
                 push args, CVar(GetArgAsString(ctx, i + 2))
            ElseIf meth.ArgTypes(i) = "long" Then
                push args, DukOp(opd_GetInt, ctx, i + 2)
            ElseIf meth.ArgTypes(i) = "integer" Then
                push args, CInt(DukOp(opd_GetInt, ctx, i + 2))
            ElseIf meth.ArgTypes(i) = "bool" Then
                push args, CBool(GetArgAsString(ctx, i + 2))
            End If
        Next
    End If
    
    Err.Clear
    
    If meth.retIsObj Then
        Set retVal = CallByNameEx(o, meth.name, meth.CallType, args(), meth.retIsObj)
    Else
        retVal = CallByNameEx(o, meth.name, meth.CallType, args(), meth.retIsObj)
    End If
    
    t = meth.CallType
    cb_HostResolver = IIf(meth.hasRet, 1, 0)
    
    'todo handle all ret types..
    If LCase(meth.retType) = "string" Or LCase(meth.retType) = "variant" Or LCase(meth.retType) = "boolean" Then
        dbg "returning string"
        DukOp opd_PushStr, ctx, 0, CStr(retVal)
        If t <> VbLet Then cb_HostResolver = 1
    ElseIf LCase(meth.retType) = "long" Then
        dbg "returning long"
        DukOp opd_PushNum, ctx, CLng(retVal)
        If t <> VbLet Then cb_HostResolver = 1
    End If
        
    If meth.retIsObj Then
        dbg "returning new js class " & meth.retType
        DukPushNewJSClass ctx, meth.retType & "Class", ObjPtr(retVal)
        objs.Add retVal, "obj:" & ObjPtr(retVal)
    End If
    
    'If Err.Number <> 0 Then MsgBox Err.Description Else MsgBox retVal
    
    
End Function


'listbox.additem ..even if the v(0) is String..its adding it as a strptr pointer..(must be taking as a long unless i wrap outter in cstr() fuck you..
'this is stupid..but tli.invokehook doesnt always work where the built in one does (listbox.additem)
'and it adds another external dependancy..so screw it..
Public Function CallByNameEx(obj As Object, ProcName As String, CallType As VbCallType, Optional v As Variant, Optional isObj As Boolean = False)

        Dim ProcID As Long
        Dim numArgs As Long

        On Error GoTo Handler
        
        'callbyName has some weird nuances..apparently v(0) as variant is
        'not the same as a as variant even if both contain a string..
        Dim a, b, c, d, e, f, g, h, i, j
    
        If Not IsArray(v) Or AryIsEmpty(v) Then
            dbg "CallByName: " & TypeName(obj) & ProcName & " " & isObj
            If isObj Then
                Set CallByNameEx = CallByName(obj, ProcName, CallType)
            Else
                CallByNameEx = CallByName(obj, ProcName, CallType)
            End If
        Else
            numArgs = UBound(v)
            
            If numArgs > 9 Then
                MsgBox "CallByNameEx does not support more than 10 args.. method: " & ProcName, vbCritical
            End If
            
            dbg "CallByName: " & TypeName(obj) & " " & ProcName & " " & isObj & " " & Join(v, ", ")
            
            If numArgs >= 0 Then a = v(0)
            If numArgs >= 1 Then b = v(1)
            If numArgs >= 2 Then c = v(2)
            If numArgs >= 3 Then d = v(3)
            If numArgs >= 4 Then e = v(4)
            If numArgs >= 5 Then f = v(5)
            If numArgs >= 6 Then g = v(6)
            If numArgs >= 7 Then h = v(7)
            If numArgs >= 8 Then i = v(8)
            If numArgs >= 9 Then j = v(9)
            
            If isObj Then
                Select Case numArgs
                    Case 0: Set CallByNameEx = CallByName(obj, ProcName, CallType, a)
                    Case 1: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b)
                    Case 2: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c)
                    Case 3: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d)
                    Case 4: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e)
                    Case 5: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f)
                    Case 6: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g)
                    Case 7: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h)
                    Case 8: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h, i)
                    Case 9: Set CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h, i, j)
                End Select
            Else
                Select Case numArgs
                    Case 0:  CallByNameEx = CallByName(obj, ProcName, CallType, a)
                    Case 1:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b)
                    Case 2:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c)
                    Case 3:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d)
                    Case 4:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e)
                    Case 5:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f)
                    Case 6:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g)
                    Case 7:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h)
                    Case 8:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h, i)
                    Case 9:  CallByNameEx = CallByName(obj, ProcName, CallType, a, b, c, d, e, f, g, h, i, j)
                End Select
            End If
        End If

    Exit Function

Handler:
        dbg "Error in CallByNameEx: " & " " & Err.Number & " " & Err.Description
End Function


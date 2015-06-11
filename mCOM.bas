Attribute VB_Name = "mCOM"
Global fso As New CFileSystem2

Public comTypes As New Collection
Public objs As New Collection

'this we dont support..
'return types string() or arrays in general
'arguments which require object arguments
'
'todo: test property let/get , objrets
'generator .let .get fix..

Function ParseObjectToCache(name As String, obj As Object) As Boolean
    Dim cc As CCOMType
    
    If KeyExistsInCollection(comTypes, name) Then
        Set cc = comTypes(name)
        If cc.errors.Count = 0 Then ParseObjectToCache = True
        Exit Function
    End If
        
    objs.Add obj, name
    Set cc = New CCOMType
    ParseObjectToCache = cc.LoadType(name)
    comTypes.Add cc, name
    
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


'this is used for script to host app object integration..
Public Function HostResolver(ByVal buf As Long, ByVal ctx As Long, ByVal argCnt As Long) As Long
    
    Dim o As Object, tmp, args(), retVal As Variant, i As Long, hInst As Long, oo As Object
    Dim firstUserArg As Long
    Dim rv As Long
    Dim b() As Byte
    Dim key As String
    Dim cc As CCOMType
    Dim meth As CMethod
    Dim pObjName As String
    Dim a As Long
    
    On Error Resume Next
    
    key = StringFromPointer(buf)
    dbg "HostResolver: ", key, ctx, argCnt

    a = InStr(key, ".") - 1
    If a > 0 Then pObjName = Mid(key, 1, a)
    
    If Not KeyExistsInCollection(comTypes, pObjName) Then Exit Function
    Set cc = comTypes(pObjName)
    
    If Not cc.GetMethod(key, meth) Then Exit Function
    
    
'    firstUserArg = 0
'    tmp = Split(name, ":")
'    If tmp(1) = "objptr" Then
'        firstUserArg = 1
'        hInst = DukOp(opd_GetInt, ctx, 2)
'        For Each oo In objs
'            If ObjPtr(oo) = hInst Then
'                Set o = oo
'                Exit For
'            End If
'        Next
'    Else
        Set o = objs(pObjName)
'    End If
    
    If o Is Nothing Then
        dbg "Host resolver could not find object!"
        Exit Function
    End If
    
    'todo handle all arg types..
    If argCnt > 0 Then
        For i = 1 To argCnt
            If meth.ArgTypes(i) = "string" Or meth.ArgTypes(i) = "variant" Then
                 push args, GetArgAsString(ctx, i + 1)
            ElseIf meth.ArgTypes(i) = "long" Then
                push args, DukOp(opd_GetInt, ctx, i + 1)
            ElseIf meth.ArgTypes(i) = "bool" Then
                push args, CBool(GetArgAsString(ctx, i + 1))
            End If
        Next
    End If
    
    Err.Clear
    
    If meth.retIsObj Then
        Set retVal = CallByNameEx(o, meth.name, meth.CallType, args(), meth.retIsObj)
    Else
        retVal = CallByNameEx(o, meth.name, meth.CallType, args(), meth.retIsObj)
    End If
    
    HostResolver = IIf(meth.hasRet, 1, 0)
    
    'todo handle all ret types..
    If meth.retType = "string" Or meth.retType = "variant" Then
        dbg "returning string"
        DukOp opd_PushStr, ctx, 0, CStr(retVal)
        If t <> VbLet Then HostResolver = 1
    ElseIf meth.retType = "long" Then
        dbg "returning long"
        DukOp opd_PushNum, ctx, CLng(retVal)
        If t <> VbLet Then HostResolver = 1
    End If
        
    If meth.retIsObj Then
        dbg "returning new js class " & tmp(UBound(tmp))
        DukPushNewJSClass ctx, meth.retType & "Class", ObjPtr(retVal)
        objs.Add retVal, "obj:" & ObjPtr(retVal)
    End If
    
    'If Err.Number <> 0 Then MsgBox Err.Description Else MsgBox retVal
    
    
End Function


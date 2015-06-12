VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public parentName As String
Public ctype As String
Public Name As String
Public ArgNames  As New Collection
Public ArgTypes As New Collection
Public hasArgs As Boolean
Public hasRet As Boolean
Public retType As String
Public raw As String

Function DescribeSelf() As String
    Dim tmp()
    
    With Me
        push tmp, "raw: " & .raw
        push tmp, vbTab & "cttpe: " & .ctype
        push tmp, vbTab & "name: " & .Name
        If .ArgNames.Count > 0 Then push tmp, vbTab & "ArgNames: " & c2s(.ArgNames)
        If .ArgTypes.Count > 0 Then push tmp, vbTab & "ArgTypes: " & c2s(.ArgTypes)
        push tmp, vbTab & "hasArgs: " & .hasArgs
        push tmp, vbTab & "hasRet: " & .hasRet
        push tmp, vbTab & "retIsObj: " & .retIsObj
        If Len(.retType) > 0 Then push tmp, vbTab & "retType: " & .retType
    End With
    
    push tmp, vbCrLf
    DescribeSelf = Join(tmp, vbCrLf)
    
End Function

Property Get retIsObj() As Boolean
    Const intristic = "byte,string,integer,long,double,boolean,single,currency,date"
    If Len(retType) = 0 Then Exit Property
    If Not AnyOfTheseInstr(retType, intristic) Then retIsObj = True
End Property

Private Function AnyOfTheseInstr(sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    For i = 0 To UBound(tmp)
        tmp(i) = Trim(tmp(i))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function

Private Function c2s(c As Collection) As String
    Dim x, y
    If c.Count = 0 Then Exit Function
    For Each x In c
        y = y & x & ", "
    Next
    y = Mid(y, 1, Len(y) - 2)
    c2s = y
End Function

'property let whatever as long
'property let something
'Property Get Drives() As IDriveCollection
'Function GetFolder(ByVal FolderPath As String) As IFolder
'Function GetSpecialFolder ( ByVal SpecialFolder As __MIDL___MIDL_itf_scrrun_0001_0000_0002 )  As IFolder
'Sub DeleteFile(ByVal FileSpec As String, ByVal Force As Boolean)

Function parse(parentObjName As String, ByVal proto As String) As Boolean
    Dim a As Long
    Dim postfix As String
    Dim args, arg, tmp
    
    parentName = parentObjName
    proto = Replace(proto, vbCrLf, Empty)
    proto = Replace(proto, vbTab, Empty)
    proto = Replace(proto, "[", "")
    proto = Replace(proto, "]", "")
    proto = Replace(proto, "  ", " ")
    proto = Trim(proto)
    raw = proto
    
    If Len(proto) = 0 Then Exit Function
    
    'a = InStr(proto, " ")
    'If a < 1 Then Exit Function
   '
    Dim FirstWord As String
    
    'FirstWord = Mid(proto, 1, a)
    
    Dim start As Long
    start = 1
    word = LCase(nextSeg(proto, start))
    
    If word = "function" Then
        hasRet = True
        ctype = "call"
    End If
    
    If word = "sub" Then
        ctype = "call"
    End If
    
    If word = "property" Then
        word = LCase(nextSeg(proto, start)) 'next word
        If word = "set" Then Exit Function 'not supported..
        If word <> "get" And word <> "let" Then
            Debug.Print "unknown property type?: " & word
            Exit Function
        Else
            ctype = word
            If word = "get" Then hasRet = True
        End If
    End If
    
    If InStr(proto, "(") < 1 Then
        word = Mid(proto, start)
        start = start + Len(word)
        tmp = Split(word, " As ", , vbTextCompare)
        Name = Trim(tmp(0))
        If UBound(tmp) = 0 Then
            retType = "variant"
        Else
            retType = Trim(LCase(tmp(1)))
        End If
        If ctype = "let" Then
            ArgNames.Add "v"
            ArgTypes.Add retType
            retType = Empty
            hasArgs = True
        End If
        parse = True
        Exit Function
    Else
        word = nextSeg(proto, start, "(") 'next word
    End If
    
    Name = word
    
    b = InStrRev(proto, ")")
    If b > 0 And b < Len(proto) Then
        postfix = Mid(proto, b + 1)
        postfix = Trim(Replace(postfix, "As", Empty, , , vbTextCompare))
        retType = postfix
        If Len(retType) Then hasRet = True
    End If
    
    a = InStr(proto, "(")
    If a > 0 And a < b Then
        a = a + 1
        If b > a Then
            args = Mid(proto, a, b - a)
            args = Replace(args, "byref", Empty, , , vbTextCompare)
            args = Trim(Replace(args, "byval", Empty, , , vbTextCompare))
            If Len(args) > 0 Then hasArgs = True
            If InStr(args, ",") > 0 Then
                args = Split(args, ",")
                For Each arg In args
                    tmp = Split(arg, " As ", , vbTextCompare)
                    ArgNames.Add Trim(tmp(0))
                    If UBound(tmp) = 0 Then
                        ArgTypes.Add "variant"
                    Else
                        ArgTypes.Add Trim(LCase(tmp(1)))
                    End If
                Next
            Else
                tmp = Split(args, " As ", , vbTextCompare)
                ArgNames.Add Trim(tmp(0))
                If UBound(tmp) = 0 Then
                    ArgTypes.Add "variant"
                Else
                    ArgTypes.Add Trim(LCase(tmp(1)))
                End If
            End If
        End If
    End If
    
    parse = True
    
End Function

Private Function nextSeg(s As String, Optional ByRef start As Long = 1, Optional marker As String = " ") As String
    Dim a As Long
    a = InStr(start, s, marker, vbTextCompare)
   
    If a > 0 Then
        nextSeg = Trim(Mid(s, start, a - start))
    End If
    
    start = a + 1
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub



'var fso2 = {
'    OpenTextFile : function(FileName,IOMode,Create,Format){
'        return resolver("call:fso2:OpenTextFile:string:[long]:[bool]:[long]:r_objTextStreamClass", arguments.length, FileName,IOMode,Create,Format);
'    }
'
'}
'
'function TextStreamClass(){
'    this.hInst=0;
'    this.ReadAll = function(){
'        return resolver("call:objptr:ReadAll:r_string", arguments.length, this.hInst);
'    }
'}
'
'function TextBoxClass(){
'    this.hInst=0;
'}
'
'TextBoxClass.prototype = {
'    get Text (){
'        return resolver("get:objptr:Text:r_string", arguments.length, this.hInst);
'    },
'    set Text (val){
'        resolver("let:objptr:Text", arguments.length, this.hInst, val);
'    }
'};

Function GenerateJS(Optional requiresHInst As Long = 0)
    Dim t As String
    Dim hi As String
    
    hi = 0
    If requiresHInst Then hi = " this.hInst"
    
    't = vbTab & "//" & Me.raw & vbCrLf
    
    If ctype = "call" Then
        t = t & vbTab & "this." & Name & " = function(" & c2s(ArgNames) & "){" & vbCrLf & _
            vbTab & vbTab & "return resolver('" & parentName & "." & Name & "', arguments.length," & hi & IIf(ArgNames.Count = 0, "", ", ") & c2s(ArgNames) & ");" & vbCrLf & _
            vbTab & "}" & vbCrLf
            
    ElseIf ctype = "let" Then
        t = t & vbTab & "set " & Name & "(val){" & vbCrLf & _
            vbTab & vbTab & "return resolver('" & parentName & "." & Name & ".let', 1," & hi & ", val);" & vbCrLf & _
            vbTab & "}," & vbCrLf
            
    ElseIf ctype = "get" Then
        t = t & vbTab & "get " & Name & "(){" & vbCrLf & _
            vbTab & vbTab & "return resolver('" & parentName & "." & Name & ".get', 0," & hi & ");" & vbCrLf & _
            vbTab & "}," & vbCrLf
            
    End If
    
    GenerateJS = t
        
End Function


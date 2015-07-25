VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCOMType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public methods As New Collection
Public dependancies As New Collection
Public errors As New Collection
Public js_file As String

Function GetMethod(key As String, outMeth As CMethod) As Boolean
    If KeyExistsInCollection(methods, key) Then
        Set outMeth = methods(key)
        GetMethod = True
    End If
End Function

'load the prototypes for a specific object, and all of its dependancies
Function LoadType(modname As String) As Boolean
    
    Dim key As String
    Dim m As CMethod
    Dim pd As String
    Dim protoEnd As Long, dat, d
    
    name = modname
    Set errors = New Collection
    
    pd = App.path & "\COM\"
    If Not FolderExists(pd) Then
        pd = GetParentFolder(App.path) & "\COM\"
        If Not FolderExists(pd) Then
            pd = GetParentFolder(App.path) & "\with_COM\COM\"
            If Not FolderExists(pd) Then
                errors.Add "COM sub folder does not exist.." & pd
                Exit Function
            End If
        End If
    End If
        
    If Not FileExists(pd & name & ".js") Then
        errors.Add "No js bindings for " & name
        Exit Function
    End If
    
    js_file = pd & name & ".js"
    
    dat = ReadFile(pd & name & ".js")
    protoEnd = InStr(dat, "*/")
    
    If protoEnd < 1 Then
        errors.Add "proto type block missing in js bindings for " & name
        Exit Function
    End If
    
    dat = Mid(dat, 1, protoEnd)
    dat = Replace(dat, vbTab, Empty)
    dat = Split(dat, vbCrLf)
    
    For Each d In dat
        d = Trim(d)
        If Len(d) > 0 And _
            VBA.Left(d, 1) <> "/" And _
            VBA.Left(d, 1) <> "#" And _
            VBA.Left(d, 1) <> ";" _
        Then 'its not a comment line..
            'Debug.Print "Parsing: " & d
            Set m = New CMethod
            If Not m.parse(name, d) Then
                errors.Add "Failed to parse: " & d 'lets build and show all errors at once for sanity sake..
            Else
                If m.retIsObj Then
                    If m.retType = "object" Then
                        errors.Add "Generic object types not supported: " & d
                    Else
                        dependancies.Add m.retType 'specific type we must find
                    End If
                End If
                key = m.MethodKey
                If KeyExistsInCollection(methods, key) Then
                    errors.Add "methods key already exists: " & key
                Else
                    methods.Add m, key
                End If
            End If
        End If
    Next
        
    If errors.count > 0 Then Exit Function
    LoadType = True
    
End Function






VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDukTape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private hDuk As Long
Private mTimeout As Long
Private mRunning As Boolean

'check for errors either manually after or receive through event..
'see ocx build for a version which raises hard errors like msscript control (easier to trap sometimes)
Public hadError As Boolean
Public LastError As String
Public LastErrorLine As Long

Event Error(ByVal line As Long, ByVal desc As String)

Property Get Running() As Boolean
    Running = mRunning
End Property

Property Let Timeout(v As Long)
    mTimeout = v
    DukOp opd_ScriptTimeout, 0, v
End Property

Property Get Timeout() As Long
    Timeout = mTimeout
End Property

Property Get Context() As Long
    Context = hDuk
End Property

Function AddObject(obj As Object, name As String, Optional ByRef outErrList As Collection) As Boolean
    
    Dim d
    Dim cc As CCOMType
    Dim result As Boolean
    Dim dd As CCOMType
    
    result = ParseObjectToCache(name, obj, Me)
    Set cc = comTypes(name) 'always added even with errors
       
    If Not result Then
       Set outErrList = cc.errors
       Exit Function
    End If
        
    If Not Me.AddFile(cc.js_file) Then
        outErrList.Add "Error loading js File: " & cc.js_file & " " & LastError
        Exit Function
    End If
        
    For Each d In cc.dependancies
        If Not comTypeExists(d) Then
            If Not AddObject(Nothing, CStr(d), outErrList) Then
                Exit Function
            End If
        End If
    Next
    
    AddObject = True
    
End Function

Private Sub Class_Initialize()
    If mDuk.hDukLib = 0 Then
        If Not InitDukLib() Then
            hadError = True
            LastError = "Library not initilized"
            RaiseEvent Error(LastErrorLine, LastError)
            Exit Sub
        End If
    End If
    hDuk = DukCreate()
    mDuk.libRefCnt = mDuk.libRefCnt + 1
End Sub

Private Sub Class_Terminate()
    If hDuk <> 0 Then DukOp opd_Destroy, hDuk
    mDuk.libRefCnt = mDuk.libRefCnt - 1
    If IsIde() And mDuk.libRefCnt = 0 Then
        'in the IDE we release it as soon as not needed..that way IDE doesnt
        'ever hang onto the dll after debugging. If it does we cant recompile the dll
        'without closing VB IDE which is annoying...
        'FreeLibrary mDuk.hDukLib
    End If
End Sub

Private Sub ClearErr()
    hadError = False
    LastError = Empty
    LastErrorLine = 0
End Sub

Function AddFile(fPath) As Boolean
    Dim rv As Long
    Dim tmp As String
    
    If hDuk = 0 Then Exit Function
    ClearErr
    
    Debug.Print "AddFile: " & fPath
    If InStr(1, fPath, "textstream", vbTextCompare) > 0 Then Stop
    
    mRunning = True
    rv = mDuk.AddFile(hDuk, CStr(fPath))
    mRunning = False
    
    If rv = 0 Then
        AddFile = True
        Exit Function
    End If
    
    On Error Resume Next
    hadError = True
    LastError = GetLastString()
    
    rv = InStr(LastError, "(line")
    If rv > 0 Then
        tmp = Mid(LastError, rv + 5)
        'LastError = Mid(LastError, 1, rv - 1)
        tmp = Left(tmp, Len(tmp) - 1)
        tmp = Trim(tmp)
        LastErrorLine = CLng(tmp)
    End If
    
    RaiseEvent Error(LastErrorLine, LastError)
    
End Function

Function Eval(js)
    Dim rv As Long
    Dim tmp As String
    
    If hDuk = 0 Then Exit Function
    ClearErr
    
    mRunning = True
    rv = mDuk.Eval(hDuk, CStr(js))
    mRunning = False
    
    If rv = 0 Then
        Eval = GetLastString()
        Exit Function
    End If
    
    On Error Resume Next
    hadError = True
    LastError = GetLastString()
   
    rv = InStr(LastError, "(line")
    If rv > 0 Then
        tmp = Mid(LastError, rv + 5)
        'LastError = Mid(LastError, 1, rv - 1)
        tmp = Left(tmp, Len(tmp) - 1)
        tmp = Trim(tmp)
        LastErrorLine = CLng(tmp)
    End If
    
    RaiseEvent Error(LastErrorLine, LastError)
    
End Function

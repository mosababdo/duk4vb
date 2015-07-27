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
Attribute VB_Exposed = True
'Author: David Zimmer <dzzie@yahoo.com>
'Site: Sandsprite.com
'License: http://opensource.org/licenses/MIT

Option Explicit
Private hDuk As Long
Private mTimeout As Long
Private mRunning As Boolean
Private mDebugging As Boolean

'check for errors either manually after or receive through event..
Public hadError As Boolean
Public LastError As String
Public LastErrorLine As Long
Public userCOMDir As String

Event Error(ByVal line As Long, ByVal desc As String)
Event dbgOut(msg As String)

Friend Sub doDbgOut(msg As String)
    RaiseEvent dbgOut(msg)
End Sub

Property Get running() As Boolean
    running = mRunning
End Property

Public Property Get isDebugging() As Boolean
    isDebugging = mDebugging
End Property

'only visible within our ocx because it request special integration
'and whole point of the ocx is to hide complexity and make it easy to use.
Friend Function DebugAttach(Optional attach As Boolean = True) As Boolean
    If attach Then
        If mDebugging Then Exit Function
        'If Not ActiveDebuggerClass Is Nothing Then Exit Function
        If mTimeout <> 0 Then DukOp opd_ScriptTimeout, 0, 0
        If DukOp(opd_debugAttach, hDuk, 1) = 0 Then
            DebugAttach = True
            mDebugging = True
            'Set ActiveDebuggerClass = Me
        Else
            If mTimeout <> 0 Then DukOp opd_ScriptTimeout, 0, mTimeout
        End If
    Else
        If Not mDebugging Then Exit Function
        'If ObjPtr(ActiveDebuggerClass) <> ObjPtr(Me) Then Exit Function
        If mTimeout <> 0 Then DukOp opd_ScriptTimeout, 0, mTimeout
        If DukOp(opd_debugAttach, hDuk, 0) = 0 Then
            DebugAttach = True
            mDebugging = False
            'Set ActiveDebuggerClass = Nothing
        End If
    End If
End Function

Property Let Timeout(v As Long)
    mTimeout = v
    DukOp opd_ScriptTimeout, 0, v
End Property

Property Get Timeout() As Long
    Timeout = mTimeout
End Property

Property Get context() As Long
    context = hDuk
End Property

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
        'FreeLibrary mDuk.hDukLib '<--to dangerous to use after initial dev of dll is done..
    End If
    'ResetComObjects 'because these are global we can not reset them in a single class terminate..
End Sub

Function Reset() As Boolean
    
    'invalid while running!
    If mDebugging Or mRunning Then
        Reset = False
        Exit Function
    End If
    
    Class_Terminate
    Class_Initialize
    Reset = True
    
End Function

Private Sub ClearErr()
    hadError = False
    LastError = Empty
    LastErrorLine = 0
End Sub

Function AddObject(obj As Object, name As String, Optional ByRef outErrList As Collection) As Boolean
    
    Dim d
    Dim cc As CCOMType
    Dim result As Boolean
        
    If outErrList Is Nothing Then Set outErrList = New Collection
    
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
        If Not AddObject(Nothing, CStr(d), outErrList) Then
            Exit Function
        End If
    Next
    
    AddObject = True
    
End Function

Function AddFile(fpath) As Boolean
    Dim rv As Long
    Dim tmp As String
    
    If hDuk = 0 Then Exit Function
    ClearErr

    mRunning = True
    Set ActiveDukTapeClass = Me
    rv = mDuk.AddFile(hDuk, CStr(fpath))
    mRunning = False
    Set ActiveDukTapeClass = Nothing
    
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
    Set ActiveDukTapeClass = Me
    rv = mDuk.Eval(hDuk, CStr(js))
    Set ActiveDukTapeClass = Nothing
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

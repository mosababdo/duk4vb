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

'check for errors either manually after or receive through event..
Public hadError As Boolean
Public LastError As String
Public LastErrorLine As Long

Event Error(ByVal line As Long, ByVal desc As String)

Property Get Context() As Long
    Context = hDuk
End Property

Private Sub Class_Initialize()
    If mDuk.hDukLib = 0 Then
        If Not InitDukLib() Then Exit Sub
    End If
    hDuk = DukCreate()
End Sub

Private Sub Class_Terminate()
    If hDuk <> 0 Then DukOp opd_Destroy, hDuk
End Sub

Function AddFile(fPath) As Boolean
    Dim rv As Long
    Dim tmp As String
    
    hadError = False
    LastError = Empty
    LastErrorLine = 0

    If hDuk = 0 Then
        hadError = True
        LastError = "Library not initilized"
        RaiseEvent Error(LastErrorLine, LastError)
        Exit Function
    End If
    
    rv = mDuk.AddFile(hDuk, fPath)
    
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
    
    hadError = False
    LastError = Empty
    LastErrorLine = 0

    If hDuk = 0 Then
        hadError = True
        LastError = "Library not initilized"
        RaiseEvent Error(LastErrorLine, LastError)
        Exit Function
    End If

    rv = mDuk.Eval(hDuk, js)
    
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResponseBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: David Zimmer <dzzie@yahoo.com>
'Site: Sandsprite.com
'License: http://opensource.org/licenses/MIT

'This class handles responses back to the duktape debugger.
'it includes primitives to build up the binary protocol response
'it is a module level class in mDuk.bas since it has to work across call backs
'once an entire packet is built, it holds the whole thing in an internal buffer
'it allows the duktape engine to sip off of its data a few bytes at a time
'since this is how the engine is implemented. In order to achieve this
'(and keep my sanity) it implements a file pointer like mechanism where
'an internal pointer keeps incrementing as data is accessed so the next
'sip is always the next available data. There is no error handling because I
'want to know immediatly if there is a flaw in the logic.

Private buf() As Byte
Public Position As Long
Public length As Long
Public isEmpty  As Boolean

Function GetBuf(sz As Long, outBuf() As Byte) As Boolean
        
    Dim endAt As Long
    Dim ub As Long
    ub = UBound(buf)
    
    Erase outBuf
    endAt = Position + sz - 1
    If endAt > length Then
        Exit Function
    End If
    
    For i = 1 To sz
        bpush outBuf, buf(Position)
        Position = Position + 1
        If Position > length Then
            Exit For
        End If
    Next
    
    'Debug.Print bHexDump(outBuf)
    If Position = length Then isEmpty = True
    GetBuf = True
    
End Function

Sub Clear()
    Erase buf
    Position = 0
    length = 0
    isEmpty = True
End Sub

Sub AddByte(b)
    bpush buf, CByte(b)
    length = length + 1
    isEmpty = False
End Sub

Sub AddBuf(b() As Byte)
    Dim i As Long
    For i = 0 To UBound(b)
        AddByte b(i)
    Next
End Sub

Sub AddIntVal(ByVal i As Long)
    Dim b(4) As Byte
    i = htonl(i)
    b(0) = &H10 'prefix specifies a 4 byte int follows..
    CopyMemory ByVal VarPtr(b(1)), ByVal VarPtr(i), 4
    'Debug.Print bHexDump(b)
    For i = 0 To UBound(b)
        AddByte b(i)
    Next
End Sub

Sub AddRawInt(ByVal i As Long)
    Dim b(3) As Byte
    i = htonl(i)
    CopyMemory ByVal VarPtr(b(0)), ByVal VarPtr(i), 4
    'Debug.Print bHexDump(b)
    For i = 0 To UBound(b)
        AddByte b(i)
    Next
End Sub

Sub AddString(s As String)
     Dim tmp() As Byte
     tmp = StrConv(s, vbFromUnicode, LANG_US)
     AddByte &H11
     AddRawInt Len(s)
     AddBuf tmp
End Sub



'for a discussion on isNormalCall see: (last section after "The answer is yes!")
'  http://sandsprite.com/blogs/index.php?uid=11&pid=353
'
Function ConstructMessage(d As Debug_Commands, Optional arg1, Optional arg2) As Boolean
    
    Clear
    Dim b As Byte
    
    Select Case d
        Case dc_stepInto: b = STEP_INTO_REQ
        Case dc_StepOver: b = STEP_OVER_REQ
        Case dc_stepout: b = STEP_OUT_REQ
        Case dc_break: b = PAUSE_REQ
        Case dc_Resume: b = RESUME_REQ
        Case dc_GetLocals: b = GET_LOCALS_REQ
        Case dc_GetCallStack: b = GET_CALL_STACK_REQ
    End Select
    
    If b <> 0 Then
        AddByte DUK_DBG_MARKER_REQUEST
        AddByte &H80 + b
        AddByte DUK_DBG_MARKER_EOM
        ConstructMessage = True
        Exit Function
    End If
    
    'all of the ones below here require arguments...
    If d = dc_GetVar Then
        'REQ <int: 0x1a> <str: varname> EOM
        AddByte DUK_DBG_MARKER_REQUEST
        AddByte &H80 + &H1A                   'DUK_DBG_CMD_GETVAR
        AddString CStr(arg1)
        AddByte DUK_DBG_MARKER_EOM
        ConstructMessage = True
        Exit Function
    End If
    
    If d = dc_SetBreakpoint Then
        'REQ <int: 0x18> <str: fileName> <int: line> EOM
        AddByte DUK_DBG_MARKER_REQUEST
        AddByte &H80 + &H18
        AddString CStr(arg1) 'filename
        AddIntVal CLng(arg2) 'line number
        AddByte DUK_DBG_MARKER_EOM
        ConstructMessage = True
        Exit Function
    End If
    
    If d = dc_delBreakpoint Then
        'REQ <int: 0x19> <int: index> EOM
        AddByte DUK_DBG_MARKER_REQUEST
        AddByte &H80 + &H19
        AddIntVal CLng(arg1) 'index
        AddByte DUK_DBG_MARKER_EOM
        ConstructMessage = True
        Exit Function
    End If
 
    If d = dc_Eval Then
        'REQ <int: 0x1e> <str: expression> EOM
        AddByte DUK_DBG_MARKER_REQUEST
        AddByte &H80 + &H1E
        AddString CStr(arg1)
        AddByte DUK_DBG_MARKER_EOM
        ConstructMessage = True
        Exit Function
    End If
    
    
End Function


Private Sub Class_Initialize()
    Clear
End Sub

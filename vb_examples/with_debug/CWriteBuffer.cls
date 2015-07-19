VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWriteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'when duktape sends us info it does so in multiple packets..here we buffer it until an end of message
'marker is received (flagging on sz=1, val=0)
'
'this class also implements a filepointer like read ability where the position is incremented by
'the readsize so you can keep calling .readint or .readbyte etc and it will auto walk through the buffer.
'internal prefixes like the flags for this is an int or string and string length are auto accounted for.
'also note this build of duktape has been modified, originally it was designed to encode small strings and
'ints into a compacted format for debug protocols over things like serial port transfers or network.
'since those dont matter here in a full system with direct callbacks..I disabled that for simplicity.
'search the source for DUK_DBG_USE_PACKED_PROTOCOL to see the small mod

Private buf() As Byte
Private mPosition As Long
Public DisplayDebugOutput As Boolean
Public firstMessage As Boolean

Property Get BytesLeft() As Long
    BytesLeft = length - Position
End Property

Property Get length() As Long
    length = UBound(buf) + 1
End Property

Property Get Position() As Long
    Position = mPosition
End Property

Function ReadByte() As Byte
    ReadByte = buf(mPosition)
    mPosition = mPosition + 1
End Function

Function ReadInt(Optional withPrefix As Boolean = True) As Long
    Dim r As Long, b As Byte
    If withPrefix Then
        b = ReadByte
        If b <> &H10 Then Debug.Print "Error in ReadInt unknown prefix!"
    End If
    CopyMemory ByVal VarPtr(r), ByVal VarPtr(buf(mPosition)), 4
    mPosition = mPosition + 4
    ReadInt = htonl(r)
End Function

Function ReadString(Optional withPrefix As Boolean = True) As String
    Dim r As String, b As Byte, sz As Long, m() As Byte
    If withPrefix Then
        b = ReadByte
        If b <> &H11 Then Debug.Print "Error in ReadString unknown prefix!"
    End If
    sz = Me.ReadInt(False)
    For i = 1 To sz
        bpush m, Me.ReadByte
    Next
    ReadString = StrConv(m, vbUnicode, LANG_US)
End Function

Function ReadDouble() As Double
    Dim d As Double
    If Me.BytesLeft < 8 Then Exit Function
    CopyMemory ByVal VarPtr(d), ByVal VarPtr(buf(mPosition)), 8
    mPosition = mPosition + 8
    ReadDouble = hto64(d)
End Function

Sub WriteBuf(b() As Byte)
 
    
    'Const eom = "-END-OF-MESSAGE-"
    
    If UBound(b) = 0 And b(0) = 0 Then
        mPosition = 0
        Call DebuggerMessageReceived
        firstMessage = False
        replyReceived = True
        Erase buf
        Exit Sub
    End If
    
    For i = 0 To UBound(b)
         bpush buf, b(i)
    Next
    
End Sub

'we use a property here since VB6 IDE will auto evaluate its value (ie run this code)
'just when you mouse over it at debug time..so you can wrap in if false to only
'see on demand..on mouse hover..fancy..thanks dillatante for the idea from your hexdump class :)
Function DebugDump(Optional prefix As String)
    dbg "MessageReceived: " & prefix
    dbg bHexDump(buf)
End Function

Property Get Hint()
    Hint = bHexDump(buf)
End Property

Private Sub Class_Initialize()
    DisplayDebugOutput = True
    firstMessage = True
End Sub

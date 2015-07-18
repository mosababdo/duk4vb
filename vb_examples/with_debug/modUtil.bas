Attribute VB_Name = "modUtil"
Global Const LANG_US = &H409
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long


Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
     ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const WM_ACTIVATE As Long = &H6

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_CANCEL = &H3
Private Const VK_CONTROL = &H11

'jcis
'http://www.vbforums.com/showthread.php?672465-RESOLVED-How-to-clear-Immediate-Window-in-IDE
Public Sub ClearImmediateWindow()
Dim lWinVB As Long, lWinEmmediate As Long
    
      Debug.Print String(30, vbCrLf)
      
'    'DO NOT SET A BREAKPOINT IN HERE IT WILL STEAL ACTIVATION AND SCREW YOUR SOURCE!
'
'    keybd_event VK_CANCEL, 0, 0, 0  ' (This is Control-Break)
'
'    lWinVB = FindWindow("wndclass_desked_gsk", vbNullString)
'    'Last param depends on languages, use your inmediate window caption:
'    lWinEmmediate = FindWindowEx(lWinVB, ByVal 0&, "VbaWindow", "Immediate")
'
'    PostMessage lWinEmmediate, WM_ACTIVATE, 1, 0&
'
'    keybd_event VK_CONTROL, 0, 0, 0 'Select All
'    keybd_event vbKeyA, 0, 0, 0 'Select All
'    keybd_event vbKeyDelete, 0, 0, 0 'Clear
'
'    keybd_event vbKeyA, 0, KEYEVENTF_KEYUP, 0
'    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
'    keybd_event vbKeyF5, 0, 0, 0   'Continue execution
'    keybd_event vbKeyF5, 0, KEYEVENTF_KEYUP, 0
    
End Sub

Function bHexDump(b() As Byte) As String
    Dim s As String
    s = StrConv(b(), vbUnicode, LANG_US)
    bHexDump = HexDump(s)
End Function

Function HexDump(ByVal str, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    
    offset = 0
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function


Function StringFromPointer(buf As Long) As String
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
    
    'Form1.List1.AddItem tmp
    Debug.Print tmp
    
End Sub

Function GetParentFolder(path) As String
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

'Function FileExists(path) As Boolean
'  On Error Resume Next
'  If Len(path) = 0 Then Exit Function
'  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
'End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init:     ReDim ary(0): ary(0) = Value
End Sub

Sub papush(ary, ParamArray Value()) 'this modifies parent ary object
    
    Dim i As Long
    
    If AryIsEmpty(ary) Then
        ReDim ary(0)
    Else
        push ary, ""
    End If
    
    i = UBound(ary)
    
    For Each v In Value
        ary(i) = ary(i) & " " & v
    Next
    
End Sub

Sub bpush(ary, Value As Byte, Optional freshStart As Boolean = False)   'this modifies parent ary object
    On Error GoTo Init
    If freshStart Then Erase ary
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init:     ReDim ary(0): ary(0) = Value
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

Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error Resume Next
    Dim t
    t = c(val)
    If Err.Number = 0 Then
        KeyExistsInCollection = True
        Exit Function
    Else
        Err.Clear
        Set t = c(val) 'maybe its an object collection?
        KeyExistsInCollection = (Err.Number = 0)
    End If
End Function

Function c2s(c As Collection) As String
    Dim x, y
    If c.count = 0 Then Exit Function
    For Each x In c
        y = y & x & ", "
    Next
    y = Mid(y, 1, Len(y) - 2)
    c2s = y
End Function

Function FolderExists(path As String) As Boolean
  On Error GoTo hell
  Dim tmp As String
  tmp = path & "\"
  If Len(tmp) = 1 Then Exit Function
  If Dir(tmp, vbDirectory) <> "" Then FolderExists = True
  Exit Function
hell:
    FolderExists = False
End Function

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function ReadFile(fileName) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open fileName For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(fileName), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Function hto64(d As Double) As Double
    Dim a As Long
    Dim b As Long
    Dim dd As Double
    
    CopyMemory ByVal VarPtr(a), ByVal VarPtr(d), 4
    CopyMemory ByVal VarPtr(b), ByVal VarPtr(d) + 4, 4
    
    a = htonl(a)
    b = htonl(b)
    
    CopyMemory ByVal VarPtr(dd), ByVal VarPtr(b), 4
    CopyMemory ByVal VarPtr(dd) + 4, ByVal VarPtr(a), 4
    
    hto64 = dd
    
End Function
 

Attribute VB_Name = "modUtil"
'Author: David Zimmer <dzzie@yahoo.com>
'Site: Sandsprite.com
'License: http://opensource.org/licenses/MIT

Option Explicit
Global Const LANG_US = &H409
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
     ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const WM_ACTIVATE As Long = &H6

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_CANCEL = &H3
Private Const VK_CONTROL = &H11

Function bHexDump(b() As Byte) As String
    Dim s As String
    s = StrConv(b(), vbUnicode, LANG_US)
    bHexDump = HexDump(s)
End Function

Function HexDump(ByVal str, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String, i As Long, h, tt, x
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

Sub dbg(msg As String)
    
    Dim tmp As String
    Dim disp As String
    Dim li As ListItem
    
    tmp = Replace(msg, vbCr, Empty)
    tmp = Replace(tmp, vbLf, Chr(5))
    tmp = Replace(tmp, Chr(5), vbCrLf)
    
    disp = Replace(msg, vbCr, "\r")
    disp = Replace(disp, vbLf, "\n")
    disp = Replace(disp, vbTab, "\t")
    
    On Error Resume Next
    If Not ActiveDukTapeClass Is Nothing Then
        ActiveDukTapeClass.doDbgOut tmp
    ElseIf Not ActiveUserControl Is Nothing Then
        ActiveUserControl.duk_dbgOut tmp
    Else
        Debug.Print "No active control/class? dbg: " & msg
    End If
    
End Sub

Function GetParentFolder(path) As String
    On Error Resume Next
    Dim tmp, ub
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

Sub push(ary, value) 'this modifies parent ary object
Dim x
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Sub papush(ary, ParamArray value()) 'this modifies parent ary object
    
    Dim i As Long, v
    
    If AryIsEmpty(ary) Then
        ReDim ary(0)
    Else
        push ary, ""
    End If
    
    i = UBound(ary)
    
    For Each v In value
        ary(i) = ary(i) & " " & v
    Next
    
End Sub

Sub bpush(ary, value As Byte, Optional freshStart As Boolean = False)   'this modifies parent ary object
    On Error GoTo init
    Dim x
    If freshStart Then Erase ary
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub




Function IsIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function


Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
  Dim i
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function AnyOfTheseInstr(sIn, sCmp) As Boolean
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
    Dim x, Y
    If c.Count = 0 Then Exit Function
    For Each x In c
        Y = Y & x & ", "
    Next
    Y = Mid(Y, 1, Len(Y) - 2)
    c2s = Y
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

Function FileNameFromPath(fullpath) As String
Dim tmp
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    End If
End Function

Function ReadFile(FileName) As Variant
  Dim f As Long
  Dim temp As Variant
  f = FreeFile
  temp = ""
   Open FileName For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(FileName), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path As String, it As String)

    If FileExists(path) Then Kill path
        
    Dim f As Long
    Dim b() As Byte
    
    b() = StrConv(it, vbFromUnicode, LANG_US)
    
    f = FreeFile
    Open path For Binary As #f
    Put f, , b()
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
 
Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function GetFreeFileName(ByVal folder, Optional ByVal extension = ".txt") As String
    
    If Not FolderExists(CStr(folder)) Then Exit Function
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    If Left(extension, 1) <> "." Then extension = "." & extension
    
    Dim tmp As String
    Do
      tmp = folder & RandomNum() & extension
    Loop Until Not FileExists(tmp)
    
    GetFreeFileName = tmp
End Function



Function CountOccurances(it, find) As Integer
    Dim tmp() As String
    If InStr(1, it, find, vbTextCompare) < 1 Then CountOccurances = 0: Exit Function
    tmp = Split(it, find, , vbTextCompare)
    CountOccurances = UBound(tmp)
End Function

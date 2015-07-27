VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScriptFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:   dzzie@yahoo.com
'Site:     http://sandsprite.com


Public lv As Object


Dim alertCount As Long
Dim showingAlerts As Boolean
Dim alertCycles As Long



Function GetListviewData(Index As Long)
    On Error Resume Next
    GetListviewData = lv.ListItems(Index).Tag
End Function

Function ResetAlertCount()
    alertCount = 0
    alertCycles = 0
    showingAlerts = True
    Form2.sc.Timeout = 10000
End Function

Function DebugLog(msg, Optional data)
    On Error Resume Next
    Dim li As ListItem
    Set li = Form2.lv2.ListItems.Add(, , msg)
    If LCase(msg) = "exploit" Then li.ForeColor = vbRed
    If Len(data) > 0 Then li.SubItems(1) = data
End Function
 

 

    
Public Sub alert(x)
    On Error Resume Next
    Err.Clear
    alertCount = alertCount + 1
    If showingAlerts And alertCount > 20 Then
        If MsgBox("Disable alert? (chance " & alertCycles & ")", vbYesNo) = vbYes Then
            showingAlerts = False
        End If
        alertCount = 0
        alertCycles = alertCycles + 1
        'If alertCycles > 5 Then
        '    MsgBox "More than 100 alerts have been shown setting timeout to 1"
        '    set Form2.sc =
        'End If
    End If
    If showingAlerts Then MsgBox x, , "Alert " & alertCount & " (can disable at 20)"
    If Err.Number <> 0 Then MsgBox TypeName(x)
End Sub

Public Function eval(x)
    t x
    DebugLog "eval", x
End Function

Public Function t(x)
    On Error Resume Next
    'If Len(Form2.txtOut.Text) > 0 Then
    '    DebugLog "txtOut", Form2.txtOut.Text
    'End If
    Form2.txtOut.Text = x
    DebugLog "txtOut", Form2.txtOut.Text
End Function

Public Function Save2Clipboard(x)
    
    If InStr(x, Chr(0)) > 0 Then
        MsgBox "Can not copy binary data", vbInformation
        Exit Function
    End If
    
    Clipboard.Clear
    Clipboard.SetText x

End Function


Public Function GetClipboard()
    GetClipboard = Clipboard.GetText
End Function

Public Sub writeFile(path, data)
    Dim f As Long
    Dim b() As Byte
    
    If MsgBox("You are about to write " & Hex(Len(data)) & " bytes to " & vbCrLf & path & vbCrLf & vbCrLf & "Allow?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    f = FreeFile
    b() = StrConv(data, vbFromUnicode, LANG_US)
    Open path For Binary As f
    Put f, , b()
    Close f
    
   If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
   
End Sub

Public Function HexDump(str, Optional hexOnly = 0)
    On Error Resume Next
    HexDump = fso.HexDump(str, hexOnly)
End Function

Public Function ReadFile(path)
    On Error Resume Next
    ReadFile = fso.ReadFile(path)
End Function

Public Function unescape(x)
    On Error Resume Next
    unescape = modEscapes.unescape(x)
End Function

Public Function pound_unescape(x)
    On Error Resume Next
    pound_unescape = modEscapes.pound_unescape(x)
End Function

Public Function HexString2Bytes(str)
    On Error Resume Next
    HexString2Bytes = modEscapes.HexStringUnescape(str)
        
End Function

Public Function EscapeHexString(hexstr)
    
    EscapeHexString = modEscapes.EscapeHexString(hexstr)
    
End Function

 


Private Sub Class_Initialize()
    On Error Resume Next
    Set Collab = Me
    Set lv = Form2.lv
    Set media = Me
End Sub

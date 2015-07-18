Attribute VB_Name = "modBreakpoints"

'nuance: if you remove a breakpoint of the current line of execution, it just runs

Public breakpoints As New Collection

Function isExecutableLine(sourceText As String) As Boolean
    Dim tmp As String
    On Error Resume Next

    tmp = LCase(sourceText)
    tmp = Trim(Replace(tmp, vbTab, Empty))

    If Len(tmp) = 0 Then GoTo fail
    If Left(tmp, 1) = "/" Then GoTo fail 'is comment
    If Left(tmp, 8) = "function" Then GoTo fail  'functio/sub start lines are hit more than you expect, once as it skips over it, so we block it as bp cause confusing..
    
    isExecutableLine = True
Exit Function
fail: isExecutableLine = False
End Function

Public Function BreakPointExists(fileName As String, lineNo As Long, Optional b As CBreakpoint, Optional colIndex As Long) As Boolean

    On Error Resume Next
    colIndex = 1
    For Each b In breakpoints
        If b.lineNo = lineNo And b.fileName = fileName Then
            BreakPointExists = True
            Exit Function
        End If
        colIndex = colIndex + 1
    Next
    
    colIndex = 0
    
End Function

Public Sub ToggleBreakPoint(fileName As String, lineNo As Long, sourceText As String)
  
    If BreakPointExists(fileName, lineNo) Then
        RemoveBreakpoint fileName, lineNo
    Else
        SetBreakpoint fileName, lineNo, sourceText
    End If

End Sub

'file name is case sensitive!
'sooo we need a live context to actually set breakpoints..but we can store them
'at design time, and then on initial debugger startup make sure to cycle through set breakpoints to initial set..
Public Function SetBreakpoint(ByVal fileName As String, lineNo As Long, ByVal sourceText As String) As Boolean
    Dim b As CBreakpoint
    
    If BreakPointExists(fileName, lineNo) Then
        SetBreakpoint = True
        Exit Function
    End If
    
    If Not isExecutableLine(sourceText) Then Exit Function 'just covers some basics for convience..
    
    Set b = New CBreakpoint
    
    With b
        .fileName = fileName
        .lineNo = lineNo
        .sourceText = sourceText
    End With
    
    If running Then
        If Not SyncronousSetBreakPoint(b) Then
            Debug.Print "Failed to set breakpoint: " & b.Stats
            Exit Function
        End If
    End If

    breakpoints.Add b
    If Form1.curFile = fileName Then Form1.scivb.SetMarker lineNo
    SetBreakpoint = True
    
    
End Function

Public Sub RemoveBreakpoint(fileName As String, lineNo As Long)
    Dim b As CBreakpoint
    Dim colIndex As Long
    
    If Not BreakPointExists(fileName, lineNo, b, colIndex) Then Exit Sub
    
    If running Then
        If Not SyncDelBreakPoint(b) Then
            Debug.Print "Failed to delete bp from duktape?: " & b.Stats
            Exit Sub
        End If
    End If
        
    If Form1.curFile = fileName Then Form1.scivb.DeleteMarker lineNo
    breakpoints.Remove colIndex
    
End Sub
'
'Sub ClearUIBreakpoints()
'    Dim b As CBreakpoint
'    For Each b In breakpoints
'        frmMain.scivb.DeleteMarker b.lineNo
'    Next
'End Sub
'
Sub RemoveAllBreakpoints()
    Dim b As CBreakpoint
    For Each b In breakpoints
        RemoveBreakpoint b.fileName, b.lineNo
    Next
End Sub

'called on debugger startup when first message received..
'assumes only single source file ansd is still current..! todo:
Sub InitDebuggerBpx()
    Dim b As CBreakpoint
    For Each b In breakpoints
        If Form1.curFile = b.fileName Then
            If b.sourceText = Form1.scivb.GetLineText(b.lineNo) Then
                 'we cant do a sync call here..so need full protocol (to early in startup?)
                 Set tmpBreakPoint = b
                 DebuggerCmd dc_SetBreakpoint, b.fileName, b.lineNo
                 If Len(b.errText) > 0 Then Debug.Print b.Stats
            End If
        Else
            If Not SyncronousSetBreakPoint(b) Then
                Debug.Print "Failed to set breakpoint: " & b.Stats
            End If
        End If
    Next
End Sub



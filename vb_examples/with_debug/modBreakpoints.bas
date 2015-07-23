Attribute VB_Name = "modBreakpoints"
'Author: David Zimmer <dzzie@yahoo.com>
'Site: Sandsprite.com
'License: http://opensource.org/licenses/MIT

Option Explicit
Public breakpoints As New Collection
 
'i am not going to bother with these tests anymore default behavior is acceptable..
Function isExecutableLine(sourceText As String) As Boolean
    Dim tmp As String
    On Error Resume Next
    
    isExecutableLine = True
    Exit Function
    
'    tmp = LCase(sourceText)
'    tmp = Trim(Replace(tmp, vbTab, Empty))
'    tmp = Replace(tmp, vbCr, Empty)
'    tmp = Replace(tmp, vbLf, Empty)
'
'    'bp on an empty line will stop on next line
'    'bp on a function (){ start will break on next line
'    'bp on a comment or multiline comment will break at line
'    'bp on a function close brace will set, but never hit. <--
'    'bp on close brace of a if state breaks on next line
'    'bp on a single line function with multiple statements will hit once (step into only steps once and all bypass)
'
'    If Len(tmp) = 0 Then GoTo fail
'    If tmp = "}" Then GoTo fail  'is end function/end if
'    If Left(tmp, 1) = "/" Then GoTo fail 'is comment
'
'    isExecutableLine = True
'Exit Function
'fail: isExecutableLine = False
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
    
    If BreakPointExists(fileName, lineNo, b) Then
        b.sourceText = sourceText 'just in case it changed..
        SetBreakpoint = True
        Exit Function
    End If
    
    'If Not isExecutableLine(sourceText) Then 'just covers some basics for convience..
    '    doOutput "Can not set breakpoint here, not an executable line"
    '    Exit Function
    'End If
    
    Set b = New CBreakpoint
    
    With b
        .fileName = fileName
        .lineNo = lineNo
        .sourceText = sourceText
    End With
    
    If running Then
        If Not SyncSetBreakPoint(b) Then
            dbg "Failed to set breakpoint: " & b.Stats
            Exit Function
        Else
            b.isSet = True
        End If
    End If

    breakpoints.Add b
    If Form1.curFile = fileName Then Form1.scivb.SetMarker lineNo
    SetBreakpoint = True
    
    
End Function

Public Sub RemoveBreakpoint(fileName As String, lineNo As Long)
    Dim b As CBreakpoint
    Dim colIndex As Long
    Dim cur_b As CBreakpoint
    
    If Not BreakPointExists(fileName, lineNo, b, colIndex) Then Exit Sub
    
    If running Then
        If Not SyncDelBreakPoint(b) Then
            dbg "Failed to delete bp from duktape?: " & b.Stats
            Exit Sub
        End If
        
        'we have to compact our duktape bp indexes - technically we should call relist...
        'note we specifically dont check filename its a flat array currently
        For Each cur_b In breakpoints
            If cur_b.index > b.index Then
                cur_b.index = cur_b.index - 1
            End If
        Next
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
        If Not b.isSet Then
            If Form1.curFile = b.fileName Then
                If b.sourceText = Form1.scivb.GetLineText(b.lineNo) Then
                     If Not SyncSetBreakPoint(b) Then
                        dbg "InitDebuggerBpx: Failed to set bp" & b.Stats
                     End If
                     If Len(b.errText) > 0 Then dbg b.Stats
                End If
            Else
                If Not SyncSetBreakPoint(b) Then
                    dbg "Failed to set breakpoint: " & b.Stats
                End If
            End If
        End If
    Next
End Sub



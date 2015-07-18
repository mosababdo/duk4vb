VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#78.0#0"; "scivb_lite.ocx"
Begin VB.Form Form1 
   Caption         =   "Simple DukTape JS Example"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   11205
      TabIndex        =   9
      Top             =   270
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10170
      TabIndex        =   8
      Text            =   "v1"
      Top             =   270
      Width           =   915
   End
   Begin VB.TextBox txtOut 
      Height          =   1185
      Left            =   5670
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6705
      Width           =   3165
   End
   Begin MSComctlLib.ListView lvErrors 
      Height          =   1050
      Left            =   3645
      TabIndex        =   4
      Top             =   6795
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1852
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Error"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvVars 
      Height          =   1050
      Left            =   10890
      TabIndex        =   5
      Top             =   6840
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1852
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "scope"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvCallStack 
      Height          =   1185
      Left            =   8955
      TabIndex        =   6
      Top             =   6795
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2090
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Function"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer tmrHideCallTip 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   8010
      Top             =   135
   End
   Begin MSComctlLib.ImageList ilToolbars_Disabled 
      Left            =   9225
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":010C
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0218
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0324
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0430
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":053C
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0648
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0754
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0860
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":096A
            Key             =   "Toggle Breakpoint"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   8550
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A76
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B82
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C8C
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D96
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0EA0
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FAA
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10B4
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11BE
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12C8
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13D2
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarDebug 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Start Debugger"
            Object.ToolTipText     =   "Start Debugger"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Break"
            Object.ToolTipText     =   "Break"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Toggle Breakpoint"
            Object.ToolTipText     =   "Toggle Breakpoint"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear All Breakpoints"
            Object.ToolTipText     =   "Clear All Breakpoiunts"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step In"
            Object.ToolTipText     =   "Step In"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Over"
            Object.ToolTipText     =   "Step Over"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Out"
            Object.ToolTipText     =   "Step Out"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run to Cursor"
            Object.ToolTipText     =   "Run to Cursor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin SCIVB_LITE.SciSimple scivb 
      Height          =   5865
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   10345
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   3120
      Left            =   90
      TabIndex        =   3
      Top             =   6570
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   5503
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Output"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Errors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Variables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CallStack"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Idle"
      Height          =   375
      Left            =   4095
      TabIndex        =   1
      Top             =   270
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim duk As CDukTape
Dim WithEvents sciext As CSciExtender
Attribute sciext.VB_VarHelpID = -1
 
Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
Const SC_MARK_BACKGROUND = 22
'http://www.scintilla.org/aprilw/SciLexer.bas
 
Public lastEIP As Long
Private forceStop As Boolean
Public curFile As String

Public Sub SyncUI()
    
    Dim curline As Long
    ClearLastLineMarkers
    
    curline = status.lineNumber - 1
    scivb.SetMarker curline, 1
    scivb.SetMarker curline, 3
    lastEIP = curline
    
    scivb.GotoLine curline
    scivb.SetFocus
    
    RefreshVariables
    'RefreshCallStack
    
End Sub
 


Private Sub ClearLastLineMarkers()
    Dim startPos As Long, endPos As Long

    scivb.DeleteMarker lastEIP, 1 'remove the yellow arrow
    scivb.DeleteMarker lastEIP, 3 'remove the yellow line backcolor

    'force a refresh of the specified line or it might not catch it..
    startPos = scivb.PositionFromLine(lastEIP)
    endPos = scivb.PositionFromLine(lastEIP + 1)
    scivb.DirectSCI.Colourise startPos, endPos
    
End Sub

Private Sub Command1_Click()
    'Dim v As CVariable
    'Set v = SyncronousGetVariableValue(Text1.Text)
    'MsgBox v.Value & " (" & v.varType & ")"
    
    DebuggerCmd dc_SetBreakpoint, "c:\test.js", 1
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not duk Is Nothing Then
        If duk.isDebugging Then duk.DebugAttach False
        Set duk = Nothing
    End If
End Sub

Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Run":               If running Then DebuggerCmd dc_Resume Else ExecuteScript
        Case "Start Debugger":    If running Then DebuggerCmd dc_Resume Else ExecuteScript True
        Case "Step In":           DebuggerCmd dc_stepInto
        Case "Step Over":         DebuggerCmd dc_StepOver
        Case "Step Out":          DebuggerCmd dc_stepout
'        Case "Run to Cursor":     RunToLine scivb.CurrentLine + 1
'        Case "Toggle Breakpoint": ToggleBreakPoint
'        Case "Clear All Breakpoints": RemoveAllBreakpoints
        Case "Break":                 DebuggerCmd dc_break
        Case "Stop":
                        duk.Timeout = 1
                        forceStop = True
                        DebuggerCmd dc_stepInto
                        
        
    End Select
    
End Sub

Private Sub ExecuteScript(Optional withDebugger As Boolean)
 
    Dim rv
   
    
    'List1.Clear
    
    running = True
    SetToolBarIcons
    lblStatus = "Status: " & IIf(withDebugger, "Debugging...", "Running...")
    
    forceStop = False
    Set duk = New CDukTape
    Set RecvBuffer = New CWriteBuffer
    
    If withDebugger Then
        duk.Timeout = 0
        duk.DebugAttach
    Else
        duk.Timeout = 7000 'set to 0 to disabled
    End If
     
    WriteFile curFile, scivb.Text
    rv = duk.AddFile(curFile)
    
    If Not duk Is Nothing Then 'form closing?
         If withDebugger Then duk.DebugAttach False
        
         If duk.hadError Then
             If Not forceStop Then
                MsgBox "Error: " & duk.LastError
             End If
         Else
             If Len(rv) > 0 And rv <> "undefined" Then MsgBox "eval returned: " & rv
         End If
         
         Set duk = Nothing
         ClearLastLineMarkers
         lblStatus = "Status: Idle" 'these would call form_load again if closing down..
         running = False
         SetToolBarIcons
    
    End If
    
End Sub

Private Sub SetToolBarIcons()
    Dim b As Button
    
    Set tbarDebug.ImageList = Nothing
    Set tbarDebug.ImageList = IIf(running, ilToolbar, ilToolbars_Disabled)
    
    For Each b In tbarDebug.Buttons
        If Len(b.key) > 0 Then
            b.Image = b.key
            b.ToolTipText = b.key
            If b.key <> "Run" And b.key <> "Start Debugger" Then
                b.Enabled = running
            End If
        End If
    Next
    
End Sub

Private Sub Form_Load()

    SetToolBarIcons
    lvVars.Visible = False
    lvCallStack.Visible = False
    lvErrors.Visible = False

    
    scivb.DirectSCI.HideSelection False
    scivb.DirectSCI.MarkerDefine 2, SC_MARK_CIRCLE
    scivb.DirectSCI.MarkerSetFore 2, vbRed 'set breakpoint color
    scivb.DirectSCI.MarkerSetBack 2, vbRed

    scivb.DirectSCI.MarkerDefine 1, SC_MARK_ARROW
    scivb.DirectSCI.MarkerSetFore 1, vbBlack 'current eip
    scivb.DirectSCI.MarkerSetBack 1, vbYellow

    scivb.DirectSCI.MarkerDefine 3, SC_MARK_BACKGROUND
    scivb.DirectSCI.MarkerSetFore 3, vbBlack 'current eip
    scivb.DirectSCI.MarkerSetBack 3, vbYellow

    scivb.DirectSCI.AutoCSetIgnoreCase True
    scivb.DisplayCallTips = True
    Call scivb.LoadCallTips(App.path & "\dependancies\calltips.txt")
    scivb.ReadOnly = False

    Set sciext = New CSciExtender
    sciext.Init scivb
    
    'scivb.Text = Replace(Replace("function b(c){\n\treturn c++\n}\na=0;\na = b(a)\na=b(a)", "\n", vbCrLf), "\t", vbTab)
    scivb.LoadFile App.path & "\test.js"
    curFile = App.path & "\test.js"
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With scivb
        .Width = Me.Width - .Left - 200
        ts.Width = .Width
        txtOut.Width = .Width - 200
        ts.Top = Me.Height - ts.Height - 800
        .Height = Me.Height - .Top - ts.Height - 1000
        With txtOut
            .Move ts.Left + 100, ts.Top + 150, ts.Width - 200, ts.Height - 500
            lvVars.Move .Left, .Top, .Width, .Height
            lvCallStack.Move .Left, .Top, .Width, .Height
            lvErrors.Move .Left, .Top, .Width, .Height
        End With
        SetLastColumnWidth lvCallStack
        SetLastColumnWidth lvVars
        SetLastColumnWidth lvErrors
    End With
End Sub

Private Sub SetLastColumnWidth(lv As ListView)
    lv.ColumnHeaders(lv.ColumnHeaders.count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.count).Left - 100
End Sub

Private Sub ts_Click()
    Dim i As Long
    i = ts.SelectedItem.index
    txtOut.Visible = IIf(i = 1, True, False)
    lvErrors.Visible = IIf(i = 2, True, False)
    lvVars.Visible = IIf(i = 3, True, False)
    lvCallStack.Visible = IIf(i = 4, True, False)
End Sub

''we use a timer for this to give them a chance to click on the calltip to edit the variable..
Private Sub tmrHideCallTip_Timer()
    If sciext.isMouseOverCallTip() Then Exit Sub
    tmrHideCallTip.Enabled = False
    scivb.StopCallTip
    Set selVariable = Nothing
End Sub
 
Private Sub sciext_MarginClick(lline As Long, Position As Long, margin As Long, modifiers As Long)
    'Debug.Print "MarginClick: line,pos,margin,modifiers", lLine, Position, margin, modifiers
    ToggleBreakPoint curFile, lline, scivb.GetLineText(lline)
End Sub

Private Sub sciext_MouseDwellEnd(lline As Long, Position As Long)
   If running Then tmrHideCallTip.Enabled = True
End Sub

Private Sub sciext_MouseDwellStart(lline As Long, Position As Long)
    'Debug.Print "MouseDwell: " & lLine & " CurWord: " & sciext.WordUnderMouse(Position)

    Dim txt As String
    Dim curWord As String
    Dim cv As CVariable
    
    If running Then
         curWord = sciext.WordUnderMouse(Position)
         If Len(curWord) = 0 Then Exit Sub
         Set cv = SyncronousGetVariableValue(curWord)
         If cv.varType <> DUK_VAR_NOT_FOUND Then
            scivb.SelStart = Position 'so call tip shows right under it..
            scivb.SelLength = 0
            txt = cv.Value
            If Len(txt) = 0 Then
                txt = cv.varType
            ElseIf Len(txt) > 25 Then
                txt = Mid(txt, 1, 20) & "..."
            End If
            If cv.varType = "string" Then txt = """" & txt & """"
            scivb.ShowCallTip curWord & " = " & txt
         End If
        
    End If


End Sub

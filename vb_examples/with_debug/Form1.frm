VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple DukTape JS Example"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   180
      TabIndex        =   1
      Top             =   5400
      Width           =   9780
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   675
      Width           =   9825
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
      TabIndex        =   2
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
   Begin VB.Label lblStatus 
      Caption         =   "Status: Idle"
      Height          =   375
      Left            =   4095
      TabIndex        =   3
      Top             =   270
      Width           =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Run":               If running Then DebuggerCmd dc_run Else ExecuteScript
        Case "Start Debugger":    If running Then DebuggerCmd dc_run Else ExecuteScript True
'        Case "Stop":              DebuggerCmd dc_
        Case "Step In":           DebuggerCmd dc_stepInto
        Case "Step Over":         DebuggerCmd dc_StepOver
        Case "Step Out":          DebuggerCmd dc_Stepout
'        Case "Run to Cursor":     RunToLine scivb.CurrentLine + 1
'        Case "Toggle Breakpoint": ToggleBreakPoint
'        Case "Clear All Breakpoints": RemoveAllBreakpoints
'        Case "Break":                 dbg_Break hDebugObject
        
    End Select
    
End Sub

Private Sub ExecuteScript(Optional withDebugger As Boolean)
 
    Dim rv
    Dim duk As CDukTape
    
    List1.Clear
    
    running = True
    SetToolBarIcons
    lblStatus = "Status: " & IIf(withDebugger, "Debugging...", "Running...")
    
    Set duk = New CDukTape
    
    If withDebugger Then
        duk.Timeout = 0
        duk.DebugAttach
    Else
        duk.Timeout = 7000 'set to 0 to disabled
    End If
      
    WriteFile "c:\test.js", Text1.text
    rv = duk.AddFile("c:\test.js")
    
    If withDebugger Then duk.DebugAttach False
   
    If duk.hadError Then
        MsgBox "Error: " & duk.LastError
    Else
        If Len(rv) > 0 And rv <> "undefined" Then MsgBox "eval returned: " & rv
    End If
    
    Set duk = Nothing
    
    lblStatus = "Status: Idle"
    running = False
    SetToolBarIcons
    
    
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
    
    Text1 = Replace(Replace("function b(c){\n\treturn c++\n}\na=0;\na = b(a)\na=b(a)", "\n", vbCrLf), "\t", vbTab)
            
End Sub
 

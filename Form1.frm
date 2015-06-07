VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   45
      TabIndex        =   1
      Top             =   4770
      Width           =   12705
   End
   Begin VB.TextBox Text1 
      Height          =   4605
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   12750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Dim rv As Long
    Dim hDukLib As Long
    Dim dlg As New clsCmnDlg2
    Dim fso As New CFileSystem2
    Dim tmp As String
    Dim fso2 As New Scripting.FileSystemObject
    
    Me.Visible = True
    
    hDukLib = LoadLibrary(App.Path & "\duk4vb.dll") 'to ensure the ide finds the dll
    
    If hDukLib = 0 Then
        MsgBox "Could not load duk4vb.dll?", vbCritical
        Exit Sub
    End If
    
    SetCallBacks AddressOf vb_stdout, 0, AddressOf HostResolver, AddressOf VbLineInput
    DukCreate
    AddObject dlg, "cmndlg"
    AddObject fso, "fso"
    AddObject Me, "form"
    AddObject fso2, "fso2"
 
    'CallByNameEx dlg, "OpenDialog", VbMethod, Array(0, "title", "c:\", 4)
    
     
     'Call CallByName(fso2, "OpenTextFile", VbMethod, "c:\\lastGraph.txt", 1, True)'works
    
    'Dim ts As Object
    'Set ts = CallByNameEx(fso2, "OpenTextFile", VbMethod, Array("c:\\lastGraph.txt", 1, True, 0), True) 'works..
    'MsgBox CallByNameEx(ts, "ReadAll", VbMethod)
    'End
    
    rv = AddFile(App.Path & "\test.js")
    If rv <> 0 Then
        MsgBox "Addfile Error: " & GetLastString()
    End If
    
    'fso2.OpenTextFile("'c:\\lastGraph.txt'",ForReading ,True)
    
    '+getting obj ref..
    '+returning js object
    '+calling a method on returned js object
    '-getting all ???? for ReadAll() output from com object?
    'rv = Eval("var ts = fso2.OpenTextFile('c:\\lastGraph.txt',1,true,0); v = ts.ReadAll();alert(v)") 'works!
    'rv = Eval("var ts = fso2.OpenTextFile('c:\\lastGraph.txt',1); v = ts.ReadAll();alert(v)") 'works (default args)
    
    'Eval "form.Text1.Text = 'test'" 'works!
    
    'Text1.text = "this is my message in a vb textbox!"
    'Eval "alert(form.Text1.Text)" 'works..
    
    'rv = Eval("prompt('text')") 'works
    'rv = Eval("1+2") 'works
    'Eval "alert(1+2)" 'works
    'Eval "a='testing';alert(a[0]);" 'works
    'rv = Eval("pth = cmndlg.ShowOpen(4,'title','c:\\',0); alert(fso.ReadFile(pth))") 'works
    'Eval "form.caption = 'test!'; alert(form.ReadFile('c:\\lastGraph.txt'));"
    'Eval "form.caption = 'test!';alert(form.caption)"
     
    If rv < 0 Then
        Text1.text = "Error: " & GetLastString()
    Else
        If GetLastStringSize() > 0 Then
            Text1.text = GetLastString()
        End If
    End If
    
    DukDestroy
    FreeLibrary hDukLib 'so the ide doesnt hang on to the dll and we can recompile it..
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width
End Sub

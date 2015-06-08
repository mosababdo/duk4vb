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
    
    Dim rv
    Dim duk As CDukTape
    
    Dim dlg As New clsCmnDlg2
    Dim fso As New CFileSystem2
    Dim fso2 As New Scripting.FileSystemObject
    
    Me.Visible = True
    
    'If Not InitDukLib(App.path & "\duk4vb.dll") Then
    '    MsgBox "Could not load duk4vb.dll?", vbCritical
    '    Exit Sub
    'End If
    
    Set duk = New CDukTape 'note after init/loadlib for trick to work..
    
    AddObject dlg, "cmndlg"
    AddObject fso, "fso"
    AddObject Me, "form"
    AddObject fso2, "fso2"
 
    If Not duk.AddFile(App.path & "\test.js") Then GoTo finished
       
    'rv = duk.Eval("var ts = fso2.OpenTextFile('c:\\lastGraph.txt',1,true,0); v = ts.ReadAll();alert(v)") 'works
    
    'rv = duk.Eval("var ts = fso2.OpenTextFile('c:\\lastGraph.txt',1); v = ts.ReadAll();alert(v)") 'works (default args)
    
    'duk.Eval "form.Text1.Text = 'test'" 'works
    
    Text1.text = "this is my message in a vb textbox!"
    rv = duk.Eval("form.Text1.Text + ' read back in from javascript!'")  'works..
    
    'rv = duk.Eval("prompt('text')") 'works
    'rv = duk.Eval("1+2") 'works
    'rv = duk.Eval("alert(1+2)" 'works
    'Eval hDuk,"a='testing';alert(a[0]);" 'works
    'rv = duk.Eval(hDuk,"pth = cmndlg.ShowOpen(4,'title','c:\\',0); alert(fso.ReadFile(pth))") 'works
    'duk.Eval "form.caption = 'test!'; alert(form.ReadFile('c:\\lastGraph.txt'));"
    'duk.Eval "form.caption = 'test!';alert(form.caption)"
     
finished:

    If duk.hadError Then
        Text1.text = "Error: " & duk.LastError
    Else
        If Len(rv) Then Text1.text = rv
    End If
    
    Set duk = Nothing
    FreeLibrary mDuk.hDukLib 'so the ide doesnt hang on to the dll and we can recompile it..
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
    
    hDukLib = LoadLibrary(App.Path & "\duk4vb.dll") 'to ensure the ide finds the dll
    
    If hDukLib = 0 Then
        MsgBox "Could not load duk4vb.dll?", vbCritical
        Exit Sub
    End If
    
    SetCallBacks AddressOf vb_stdout, 0, AddressOf HostResolver, 0
    DukCreate
    AddObject dlg, "cmndlg"
    AddObject fso, "fso"
    AddObject Me, "form"
 
    'CallByNameEx dlg, "OpenDialog", VbMethod, Array(0, "title", "c:\", 4)
    
    rv = AddFile(App.Path & "\test.js")
    If rv <> 0 Then
        MsgBox "Addfile Error: " & GetLastString()
    End If
    
    'rv = Eval("1+2") 'works
    'rv = Eval("alert(1+2)") 'works
    'rv = Eval("pth = cmndlg.ShowOpen(4,'title','c:\\',0); alert(fso.ReadFile(pth))") 'works
     'rv = Eval("form.caption = 'test!'; alert(form.ReadFile('c:\\lastGraph.txt'));")
     rv = Eval("form.caption = 'test!';alert(form.caption)")
'    If rv < 0 Then
'        MsgBox "Error: " & GetLastString()
'    Else
'        If GetLastStringSize() > 0 Then
'            MsgBox "result: " & GetLastString()
'        End If
'    End If
    
    DukDestroy
    FreeLibrary hDukLib 'so the ide doesnt hang on to the dll and we can recompile it..
    
End Sub

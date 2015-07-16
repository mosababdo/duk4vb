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
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   135
      TabIndex        =   3
      Top             =   4815
      Width           =   5190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11430
      TabIndex        =   2
      Top             =   4545
      Width           =   1275
   End
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
      Height          =   1950
      Left            =   45
      TabIndex        =   1
      Top             =   6435
      Width           =   12705
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

Dim dlg As Object 'New vbdevkit.clsCmnDlg2
Dim fso As Object 'new vbdevkit.Cfilesystem2


Private Sub Command1_Click()
    Form_Load
End Sub

Private Sub LoadVbDevKitTests()
    On Error Resume Next
    Set dlg = CreateObject("vbdevkit.clsCmnDlg2")
    Set fso = CreateObject("vbdevkit.Cfilesystem2")
End Sub

Private Sub Form_Load()
    
         
    Dim rv
    Dim js
    Dim duk As CDukTape
    Dim errs As New Collection
    
    Dim ms_fso As New Scripting.FileSystemObject
    
    LoadVbDevKitTests
    
    Me.Visible = True
    Text1.text = "this is my message in a vb textbox!"
      
    Set duk = New CDukTape
    
    If Not dlg Is Nothing Then
        If Not duk.AddObject(dlg, "dlg", errs) Then
            Text1 = c2s(errs)
            Exit Sub
        End If
    End If

    If Not fso Is Nothing Then
        If Not duk.AddObject(fso, "fso2", errs) Then
            Text1 = c2s(errs)
            Exit Sub
        End If
    End If

    If Not duk.AddObject(Me, "form", errs) Then
        Text1 = c2s(errs)
        Exit Sub
    End If
    
    If Not duk.AddObject(ms_fso, "fso", errs) Then
        Text1 = c2s(errs)
        Exit Sub
    End If


'test cases all currently working
'    js = "1+2"
'    js = "alert(1+2)"
'    js = "while(1){;}"                 'timeout test
'    js = "prompt('text')"
'    js = "a='testing';alert(a[0]);"
'    js = "fso2.ReadFile('c:\\lastGraph.txt')"

'------------- vbdevkit tests ---------------------
'    js = "alert(dlg.OpenDialog(4))"
'    js = "pth = dlg.OpenDialog(4,'title','c:\\',0); fso2.ReadFile(pth)"
'--------------------------------------------------

'    js = "form.Text1.Text = 'test'"
'    js = "form.Text1.Text + ' read back in from javascript!'"
'    js = "form.caption = 'test!';alert(form.caption)"
    js = "for(i=0;i<10;i++)form.List2.AddItem('item:'+i);alert('clearing!');form.List2.Clear()"
'    js = "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1,true,0);v = ts.ReadAll(); v"         'value of v is returned from eval..
'    js = "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1); v = ts.ReadAll();alert(v)"         '(default args test)
                  
    duk.Timeout = 7000 'set to 0 to disabled
    Me.Caption = "Running..."
    rv = duk.Eval(js)
    Me.Caption = "Complete."
   
    If duk.hadError Then
        Text1.text = "Error: " & duk.LastError
    Else
        If Len(rv) Then Text1.text = "eval returned: " & rv
    End If
    
    Set duk = Nothing
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width
End Sub

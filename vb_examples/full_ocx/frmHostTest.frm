VERSION 5.00
Object = "{047848A0-21DD-421D-951E-B4B1F3E1718D}#77.0#0"; "dukDbg.ocx"
Begin VB.Form frmHostTest 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   16740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   12240
      TabIndex        =   12
      Top             =   6300
      Width           =   1860
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   12060
      TabIndex        =   11
      Text            =   "a=0;if(a){a++;}else{a++;}a=0;"
      Top             =   6750
      Width           =   2535
   End
   Begin VB.ListBox lstCallStack 
      Height          =   3180
      Left            =   12015
      TabIndex        =   9
      Top             =   2880
      Width           =   4470
   End
   Begin VB.CommandButton cmdMultiInst 
      Caption         =   "MultiInst"
      Height          =   375
      Left            =   11925
      TabIndex        =   8
      Top             =   1935
      Width           =   1545
   End
   Begin VB.TextBox txtManual 
      Height          =   285
      Left            =   11880
      TabIndex        =   7
      Text            =   "1+2"
      Top             =   1170
      Width           =   2220
   End
   Begin VB.CommandButton cmdJustDuk 
      Caption         =   "JustDukCls"
      Height          =   375
      Left            =   11925
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cboTest 
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   5355
      Width           =   11310
   End
   Begin VB.TextBox txtOut 
      Height          =   1545
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "frmHostTest.frx":0000
      Top             =   7380
      Width           =   5730
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   5760
      TabIndex        =   3
      Top             =   6030
      Width           =   5685
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   315
      TabIndex        =   2
      Top             =   6885
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   315
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6255
      Width           =   5010
   End
   Begin dukDbg.ucDukDbg ucDukDbg1 
      Height          =   5010
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   8837
   End
   Begin VB.Label Label1 
      Caption         =   "Callstack"
      Height          =   285
      Left            =   12015
      TabIndex        =   10
      Top             =   2520
      Width           =   1185
   End
End
Attribute VB_Name = "frmHostTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New Scripting.FileSystemObject

Private Sub cboTest_Click()
    ucDukDbg1.Text = cboTest.Text
End Sub

Private Sub cmdJustDuk_Click()
    Dim duk As New CDukTape
    Dim rv
    
    On Error Resume Next
    
    rv = duk.Eval(txtManual.Text)
    If duk.hadError Then
        MsgBox "Error Line: " & duk.LastErrorLine & " Description:" & duk.LastError
    Else
        If rv <> "undefined" Then MsgBox rv
    End If
    
End Sub

Private Sub cmdMultiInst_Click()
    Dim f As New frmHostTest
    f.Show
End Sub

 

Private Sub Command1_Click()
        
    On Error Resume Next
    
    Dim duk As New CDukTape
    Dim c As Collection
    
    duk.userCOMDir = App.Path
    
    If Not duk.AddObject(Text2, "fartbox") Then
        MsgBox "Error Adding Object: " & duk.LastError
        Exit Sub
    End If
    
    duk.Eval "v=fartbox.Text;alert(v)"
    
    If duk.hadError Then
        MsgBox "Error: " & duk.LastError
    End If
    
End Sub

Private Sub Form_Load()
    
    List1.AddItem "Message Log"
    
     'Exit Sub
    
    'for multi instance count tests..
    If VB.Forms.Count = 1 Then
        With ucDukDbg1
            .AddObject fso, "fso"
            .AddObject Me, "form"
            .LoadFile App.Path & "\test.js"
            .AddIntellisense "fso", "BuildPath GetDriveName GetParentFolderName GetFileName GetBaseName GetExtensionName GetAbsolutePathName GetTempName DriveExists FileExists FolderExists DeleteFile DeleteFolder MoveFile MoveFolder CopyFile CopyFolder CreateTextFile OpenTextFile GetStandardStream GetFileVersion"
            .AddIntellisense "form", "caption,Text1,List2"
            .AddIntellisense "List2", "AddItem,Clear,Enabled,ListCount"
            .AddIntellisense "Text1", "Text"
            .LoadCallTips App.Path & "\calltips.txt"
        End With
    End If
    
    cboTest.AddItem "while(1){;}"                  'timeout test/break
    cboTest.AddItem "print(prompt('text'))"
    cboTest.AddItem "form.Text1.Text = 'test'"
    cboTest.AddItem "print(form.Text1.Text + ' read back in from javascript!')"
    cboTest.AddItem "form.caption = 'test!';alert(form.caption)"
    cboTest.AddItem "for(i=0;i<10;i++)form.List2.AddItem('item:'+i);alert('clearing!');form.List2.Clear()"
    cboTest.AddItem "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1,true,0);v = ts.ReadAll(); alert(v)"          'value of v is returned from eval..
    cboTest.AddItem "var ts = fso.OpenTextFile('c:\\lastGraph.txt',1); v = ts.ReadAll();alert(v)"          '(default args test)
    
    'if you want to access the embedded scintilla control
    Dim sci As SciSimple
    Set sci = ucDukDbg1.sci


End Sub



Private Sub ucDukDbg1_dbgOut(msg As String)

    List1.AddItem "dbgout: " & msg
    If msg = "cls" Then txtOut.Text = Empty
    
End Sub

Private Sub ucDukDbg1_dukErr(line As Long, msg As String)
    List1.AddItem "dukErr: " & line & " : " & msg
End Sub

Private Sub ucDukDbg1_StateChanged(state As dukDbg.dbgStates)

    If state = dsStarted Then
        List1.Clear
        List2.Clear
        txtOut.Text = Empty
    End If
    
    If state = dsPaused Then
        lstCallStack.Clear
        Dim c As Collection
        Dim cc As cCallStack
        
        Set c = ucDukDbg1.GetCallStack()
        For Each cc In c
            lstCallStack.AddItem cc.lineNo & " " & cc.func & " " & cc.fpath
        Next
    End If
    
End Sub

Private Sub ucDukDbg1_printOut(msg As String)
 
    Dim leng As Long
    Dim tmp As String
    Dim includeCRLF As Boolean
    
    tmp = Replace(msg, vbCr, Empty)
    tmp = Replace(tmp, vbLf, Chr(5))
    tmp = Replace(tmp, Chr(5), vbCrLf)
    leng = Len(txtOut.Text)
    
    If leng > 0 And Right(tmp, 2) <> vbCrLf Then includeCRLF = True
    
    txtOut.SelLength = 0
    txtOut = txtOut & IIf(includeCRLF, vbCrLf, "") & tmp
    txtOut.SelStart = leng + 2
    
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15705
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10740
   ScaleWidth      =   15705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHinst 
      Caption         =   "requires hInst"
      Height          =   330
      Left            =   5040
      TabIndex        =   7
      Top             =   3915
      Width           =   3120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "save"
      Height          =   420
      Left            =   14355
      TabIndex        =   6
      Top             =   9315
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load File"
      Height          =   375
      Left            =   11610
      TabIndex        =   5
      Top             =   3870
      Width           =   2085
   End
   Begin VB.TextBox txtClassName 
      Height          =   420
      Left            =   2430
      TabIndex        =   4
      Top             =   3870
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "parse"
      Height          =   375
      Left            =   14085
      TabIndex        =   2
      Top             =   3870
      Width           =   1275
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   4455
      Width           =   15450
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
      Height          =   3615
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   135
      Width           =   15450
   End
   Begin VB.Label Label1 
      Caption         =   "class name"
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   3915
      Width           =   1950
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dlg As New clsCmnDlg2
Dim fso As New CFileSystem2
Dim loadedFile As String

Private Sub Command1_Click()

     If Len(txtClassName) = 0 Then
        MsgBox "Class name must be filled in", vbInformation
        Exit Sub
     End If
     
     Dim tmp() As String
    
    'push tmp, "Sub DeleteFile(ByVal FileSpec As String, ByVal Force As Boolean)"

    tmp = Split(Text1, vbCrLf)
    
    Dim o()
    Dim props()
    Dim funcs()
    Dim protos()
    Dim js As String
    
    push protos, "/*"
    Dim m As CMethod
    For Each t In tmp
        t = Trim(t)
        If Len(t) > 0 And _
            VBA.Left(t, 1) <> "/" And _
            VBA.Left(t, 1) <> "#" And _
            VBA.Left(t, 1) <> ";" _
        Then
            Set m = New CMethod
            If Len(t) > 0 Then
                If Not m.parse(txtClassName, t) Then
                    push o, "Failed: " & t
                Else
                    push protos, vbTab & t
                    push o, m.DescribeSelf()
                    If m.ctype = "call" Then
                        push funcs, m.GenerateJS(chkHinst.value)
                    Else
                        push props, m.GenerateJS(chkHinst.value)
                    End If
                End If
            End If
        End If
    Next
    push protos, "*/" & vbCrLf & vbCrLf
    
    js = Join(protos, vbCrLf)
    js = js & "function " & txtClassName & "Class(){" & vbCrLf & vbCrLf
    If chkHinst.value = 1 Then js = js & vbTab & "this.hInst = 0"

    js = js & vbCrLf & vbCrLf & Join(funcs, vbCrLf) & vbCrLf & "}"
    
    If Not AryIsEmpty(props) Then
        a = InStrRev(props(UBound(props)), ",")
        props(UBound(props)) = Mid(props(UBound(props)), 1, a - 1)
        js = js & vbCrLf & vbCrLf & txtClassName & "Class.prototype = {" & vbCrLf & Join(props, vbCrLf) & vbCrLf & "}"
    End If
    
    js = js & vbCrLf & vbCrLf & "var " & txtClassName & " = new " & txtClassName & "Class()" & vbCrLf
    Text1 = Join(o, vbCrLf)
    Text2 = js
End Sub

Private Sub Command2_Click()
    loadedFile = dlg.OpenDialog(AllFiles)
    If Len(loadedFile) = 0 Then Exit Sub
    Text1 = fso.ReadFile(loadedFile)
    txtClassName = fso.GetBaseName(loadedFile)
End Sub

Private Sub Command3_Click()
Dim pd As String

    If Len(loadedFile) = 0 Then
        pd = dlg.FolderDialog()
    Else
        pd = fso.GetParentFolder(loadedFile)
    End If
    
    If Not fso.FolderExists(pd) Then Exit Sub
       
    tmp = pd & "\" & txtClassName & ".js"
    fso.WriteFile CStr(tmp), Text2.Text
    MsgBox "saved to: " & tmp
End Sub

'Function OpenTextFile(
'    ByVal FileName As String,
'    ByVal IOMode As IOMode,
'    ByVal Create As Boolean,
'    ByVal Format As Tristate
') As ITextStream

'var fso2 = {
'    OpenTextFile : function(FileName,IOMode,Create,Format){
'        return resolver("call:fso2:OpenTextFile:string:[long]:[bool]:[long]:r_objTextStreamClass", arguments.length, FileName,IOMode,Create,Format);
'    }
'
'}
'
'function TextStreamClass(){
'    this.hInst=0;
'    this.ReadAll = function(){
'        return resolver("call:objptr:ReadAll:r_string", arguments.length, this.hInst);
'    }
'}
'
'function TextBoxClass(){
'    this.hInst=0;
'}
'
'TextBoxClass.prototype = {
'    get Text (){
'        return resolver("get:objptr:Text:r_string", arguments.length, this.hInst);
'    },
'    set Text (val){
'        resolver("let:objptr:Text", arguments.length, this.hInst, val);
'    }
'};


Private Sub Form_Load()
        
   
    
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple DukTape JS Example"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   6390
      TabIndex        =   4
      Text            =   "test string from vb!"
      Top             =   4905
      Width           =   2040
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   135
      TabIndex        =   2
      Top             =   4815
      Width           =   5190
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Execute JS"
      Height          =   375
      Left            =   8685
      TabIndex        =   1
      Top             =   4905
      Width           =   1275
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
      Width           =   9825
   End
   Begin VB.Label Label1 
      Caption         =   "TxtData"
      Height          =   240
      Left            =   5625
      TabIndex        =   3
      Top             =   4950
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Command1_Click()
 
    Dim rv
    Dim duk As CDukTape
    
    List1.Clear
    
    Set duk = New CDukTape
    duk.Timeout = 7000 'set to 0 to disabled
    Me.Caption = "Running..."
    rv = duk.Eval(Text1.text)
    Me.Caption = "Complete."
   
    If duk.hadError Then
        MsgBox "Error: " & duk.LastError
    Else
        If Len(rv) > 0 And rv <> "undefined" Then MsgBox "eval returned: " & rv
    End If
    
    Set duk = Nothing
    
End Sub

Private Sub Form_Load()

    Text1 = vbCrLf & _
            "function vbClass(){" & vbCrLf & _
            "    this.additem = function(x){" & vbCrLf & _
            "         resolver(""list1.additem"",2,0,x)" & vbCrLf & _
            "    }" & vbCrLf & _
            "" & vbCrLf & _
            "    this.getTextVal = function(){" & vbCrLf & _
            "         return resolver(""text2.text"",0,0)" & vbCrLf & _
            "    }" & vbCrLf & _
            "}" & vbCrLf & _
            "" & vbCrLf & _
            "vb = new vbClass();" & vbCrLf & _
            "" & vbCrLf & _
            "for(i=0;i<10;i++) vb.additem(""str_""+i)" & vbCrLf & _
            "" & vbCrLf & _
            "alert( vb.getTextVal() )"
            
End Sub
 

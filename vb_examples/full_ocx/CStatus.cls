VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'this was annoying to generate by hand!

Private m_state As Long
Private m_fileName As String
Private m_curFunc As String
Private m_lineNumber As Long
Private m_pc As Long
Private m_lastLineNo As Long
Private m_stepToLine As Long
Private m_callStackLoaded As Boolean

Property Get state() As Long
       state = m_state
End Property

Friend Property Let state(v As Long)
   m_state = v
End Property

Property Get FileName() As String
     FileName = m_fileName
End Property

Friend Property Let FileName(v As String)
    m_fileName = v
End Property

Property Get curFunc() As String
      curFunc = m_curFunc
End Property

Friend Property Let curFunc(v As String)
    m_curFunc = v
End Property

Property Get lineNumber() As Long
     lineNumber = m_lineNumber
End Property

Friend Property Let lineNumber(v As Long)
    m_lineNumber = v
End Property

Property Get lastLineNo() As Long
      lastLineNo = m_lastLineNo
End Property

Friend Property Let lastLineNo(v As Long)
   m_lastLineNo = v
End Property

Property Get pc() As Long
     pc = m_pc
End Property

Friend Property Let pc(v As Long)
    m_pc = v
End Property

Friend Property Get StepToLine() As Long
    StepToLine = m_stepToLine
End Property

Friend Property Let StepToLine(v As Long)
    m_stepToLine = v
End Property

Friend Property Get callStackLoaded() As Boolean
    callStackLoaded = m_callStackLoaded
End Property

Friend Property Let callStackLoaded(v As Boolean)
    m_callStackLoaded = v
End Property


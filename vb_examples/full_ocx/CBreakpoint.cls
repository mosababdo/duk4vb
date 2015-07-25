VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBreakpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public owner As Long
Public lineNo As Long
Public fileName As String
Public sourceText As String
Public index As Long
Public errText As String
Public isSet As Boolean

Property Get Stats() As String
    
    Dim tmp() As String
    
    papush tmp, "BreakPoint:", index, " Line: ", lineNo
    papush tmp, "   File: ", fileName
    papush tmp, "   Error: ", errText
    papush tmp, "   Source: ", sourceText
    
    Stats = Join(tmp, vbCrLf)
    
End Property

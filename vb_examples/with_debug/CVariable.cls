VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public isGlobal As Boolean
Private m_name As String
Private m_value As Variant
Public varType As String
'Public Index As Long 'for arrays this is array index, for local variables this is its offset it local var table..
'Public pAryElement As Long

Property Let name(v As String)
    m_name = v
End Property

Property Get name() As String
    name = m_name
End Property

Property Let Value(v As String)
    m_value = v
End Property

Property Get Value() As String
    Value = m_value
End Property

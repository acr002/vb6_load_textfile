VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c13CT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: ]
  Option Explicit

' 2019.07.10(êÖ).new
'***********************************************
Private m_num As Long
Private m_buf As String
'***********************************************

Public Property Let buf(ByVal vData As String)
      m_buf = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get buf() As String
      buf = m_buf
End Property
'-----------------------------------------------------------------------------

Public Property Let num(ByVal vData As Long)
      m_num = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get num() As Long
      num = m_num
End Property
'-----------------------------------------------------------------------------


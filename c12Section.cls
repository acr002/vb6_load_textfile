VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c12Section"
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
Private m_key As String
Private m_q_type As String
Private m_cts As c13CTs
Private m_title As String
'***********************************************

Public Property Let title(ByVal vData As String)
      m_title = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get title() As String
      title = m_title
End Property
'-----------------------------------------------------------------------------

Public Property Set cts(ByVal vData As c13CTs)
      Set m_cts = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get cts() As c13CTs
      Set cts = m_cts
End Property
'-----------------------------------------------------------------------------

Public Property Let q_type(ByVal vData As String)
      m_q_type = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get q_type() As String
      q_type = m_q_type
End Property
'-----------------------------------------------------------------------------

Public Property Let key(ByVal vData As String)
      m_key = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get key() As String
      key = m_key
End Property
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
  set m_cts = New c13CTs
End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Terminate()
  set m_cts = Nothing
End Sub
'-----------------------------------------------------------------------------


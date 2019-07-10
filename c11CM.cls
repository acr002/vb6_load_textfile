VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c11CM"
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
Private m_sections As c12Sections
Private m_jobcode As String
'***********************************************

Public Property Let jobcode(ByVal vData As String)
      m_jobcode = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get jobcode() As String
      jobcode = m_jobcode
End Property
'-----------------------------------------------------------------------------

Public Property Set sections(ByVal vData As c12Sections)
      Set m_sections = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get sections() As c12Sections
      Set sections = m_sections
End Property
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
  set m_sections = New c12Sections
End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Terminate()
  set m_sections = Nothing
End Sub
'-----------------------------------------------------------------------------


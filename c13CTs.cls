VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c13CTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"c13CT"
Attribute VB_Ext_KEY = "Member0" ,"c13CT"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: ]
  Option Explicit

' 2019.07.10(êÖ).new
'***********************************************
Private mCol As Collection
'***********************************************

Public Function find(ByVal a_key As String) As c13CT
  If exist(a_key) Then
    Set find = item(a_key)
  Else
    Set find = Add(a_key)
  End If
End Function
'-----------------------------------------------------------------------------

Public Function exist(ByVal a_key As String) As Boolean
  Dim bb As c13CT
  Dim has_target As Boolean
  For Each bb In mCol
    If bb.key = a_key Then
      has_target = True
      Exit For
    End If
  Next bb
  exist = has_target
End Function
'-----------------------------------------------------------------------------

Public Function fetch(ByVal a_key As String) As c13CT
  Dim bb As c13CT
  For Each bb In mCol
    If bb.key = a_key Then
      Set fetch = bb
      Exit For
    End If
  Next bb
End Function
'-----------------------------------------------------------------------------

Public Function Add(ByVal a_key As String) As c13CT
  Dim obj_new As c13CT
  Set obj_new = New c13CT
  obj_new.key = a_key
  mCol.Add obj_new, a_key
  Set Add = obj_new
  Set obj_new = Nothing
End Function
'-----------------------------------------------------------------------------

' num As Long moduleë§Ç≈ê›íuÇµÇƒÇ≠ÇæÇ≥Ç¢
' buf As String             moduleë§Ç≈ê›íuÇµÇƒÇ≠ÇæÇ≥Ç¢
' Optional sKey As String   moduleë§Ç≈ê›íuÇµÇƒÇ≠ÇæÇ≥Ç¢
Public Function Add(num As Long, buf As String, Optional sKey As String) As c13CT
      Dim objNewMember As c13CT
      Set objNewMember = New c13CT
      objNewMember.num = num
      objNewMember.buf = buf
      If Len(sKey) = 0 Then
          mCol.Add objNewMember
      Else
          mCol.Add objNewMember, sKey
      End If
      Set Add = objNewMember
      Set objNewMember = Nothing
End Function
'-----------------------------------------------------------------------------

Public Property Get Item(vntIndexKey As Variant) As c13CT
  Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property
'-----------------------------------------------------------------------------

Public Property Get Count() As Long
      Count = mCol.Count
End Property
'-----------------------------------------------------------------------------

Public Sub Remove(vntIndexKey As Variant)
      mCol.Remove vntIndexKey
  End Sub
'-----------------------------------------------------------------------------

Public Property Get NewEnum() As IUnknown
  Attribute NewEnum.VB_UserMemId = -4
  Attribute NewEnum.VB_MemberFlags = "40"
      Set NewEnum = mCol.[_NewEnum]
End Property
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
      Set mCol = New Collection
  End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Terminate()
      Set mCol = Nothing
  End Sub
'-----------------------------------------------------------------------------


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Arr() As CBElements 'speichert Listen (=CBElements) von CBElement-Objekten
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Function Add(Obj As CBElements) As CBElements
    Set Add = Obj
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2)
    Set m_Arr(m_Count) = Add
    m_Count = m_Count + 1
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Item(ByVal Index As Long) As CBElements
    If Index >= m_Count Then Exit Property
    Set Item = m_Arr(Index)
End Property

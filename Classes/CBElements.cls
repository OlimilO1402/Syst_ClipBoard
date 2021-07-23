VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Date As Date
Private m_List As Collection 'Of CBElement

Private Sub Class_Initialize()
    Set m_List = New Collection
    m_Date = Now
End Sub

Public Sub Clear()
    Class_Initialize
End Sub

Public Function Add(aCBElement) As CBElement
    Set Add = aCBElement
    m_List.Add Add, CStr(Add.CBFormat)
End Function

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(ByVal aCBFormat As ClipboardFormat) As CBElement
    Set Item = m_List.Item(CStr(aCBFormat))
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Property Get MyDate() As Date
    Date = m_Date
End Property

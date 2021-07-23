VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CBFormat As ClipboardFormat
Private m_Arr()    As Byte

Friend Sub New_(ByVal aCBFormat As ClipboardFormat, BytArr() As Byte)
    m_CBFormat = aCBFormat:    m_Arr = BytArr
End Sub

Public Property Get CBFormat() As ClipboardFormat
    CBFormat = m_CBFormat
End Property

Public Property Get ByteSize() As Long
    ByteSize = UBound(m_Arr) '+ 1
End Property

Public Function ToStr() As String
    ToStr = MCBFormat.CBFormat_ToStr(m_CBFormat) & " " & ByteSize
End Function

Public Function ToRTF() As String
    '
End Function

Public Function ToText() As String
    '
End Function

Public Function ToHtml() As String
    '
End Function

Public Function ToXml() As String
    '
End Function

Public Function ToPicture() As StdPicture
    '
End Function

Public Function ToWave()
    '
End Function

Public Function ToEwmf() 'As Variant
    '
End Function

Attribute VB_Name = "MNew"
Option Explicit

Public Function CBElement(aCBFormat As ClipboardFormat, BytArr() As Byte) As CBElement
    Set CBElement = New CBElement: CBElement.New_ aCBFormat, BytArr
End Function

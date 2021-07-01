Attribute VB_Name = "Mewmf"
Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Declare Function PlayMetaFile Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByVal hMF As Long) As Long

Public Type METARECORD
    rdSize     As Long
    rdFunction As Integer
    rdParm(1)  As Integer
End Type
Public Declare Function PlayMetaFileRecord Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByRef lpHandletable As Long, _
     ByRef lpMetaRecord As METARECORD, _
     ByVal nHandles As Long) As Long

Public Declare Function PlayEnhMetaFile Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByVal hemf As Long, _
     ByRef lpRect As RECT) As Long


Public Type ENHMETARECORD
    iType    As Long
    nSize    As Long
    dParm(1) As Long
End Type
Public Declare Function PlayEnhMetaFileRecord Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByRef lpHandletable As Long, _
     ByRef lpEnhMetaRecord As ENHMETARECORD, _
     ByVal nHandles As Long) As Long


Public Type ENHMETAHEADER
    iType          As Long
    nSize          As Long
    rclBounds      As RECT
    rclFrame       As RECT
    dSignature     As Long
    nVersion       As Long
    nBytes         As Long
    nRecords       As Long
    nHandles       As Integer
    sReserved      As Integer
    nDescription   As Long
    offDescription As Long
    nPalEntries    As Long
    szlDevice      As Size
    szlMillimeters As Size
End Type
Public Declare Function GetEnhMetaFileHeader Lib "gdi32.dll" ( _
     ByVal hemf As Long, _
     ByVal cbBuffer As Long, _
     ByRef lpemh As ENHMETAHEADER) As Long

Public Declare Function SetClipboardViewer Lib _
    "user32" (ByVal hwnd As Long) As Long


Public Const LANG_NEUTRAL = &H0
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function ErrMessage(ByVal lerr As Long) As String
    
    Dim buffer As String: buffer = String(1024, Chr(0))
    Dim e      As Long:        e = GetLastError
    Dim rh     As Long:       rh = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, e, LANG_NEUTRAL, buffer, 200, ByVal 0&)
    ErrMessage = Replace(buffer, Chr(0), "")

End Function

'Public Const CF_TEXT             As Long = 1        ' Handle zu String
'Public Const CF_BITMAP           As Long = 2        ' Handle eines Bitmaps
'Public Const CF_METAFILEPICT     As Long = 3        ' Handle zu einem Metafile Bild
'Public Const CF_SYLK             As Long = 4
'Public Const CF_DIF              As Long = 5        ' "Software Arts' Data Interchange Format"
'Public Const CF_TIFF             As Long = 6        ' Handle zu einem Tiff-Bitmap
'Public Const CF_OEMTEXT          As Long = 7        ' Handle zu einem OEM-String
'Public Const CF_DIB              As Long = 8        ' Handle zu einer %BITMAPINFO%-Struktur
'Public Const CF_BOTTOMUP_DIB     As Long = CF_DIB
'Public Const CF_PALETTE          As Long = 9        ' Handle zu einer Palette
'Public Const CF_PENDATA          As Long = 10       ' sind Daten zu einem Microsoft Pen Extensions
'Public Const CF_RIFF             As Long = 11       ' Handle zu einer Audiodatei
'Public Const CF_WAVE             As Long = 12       ' Handle zu Wavedatei
'Public Const CF_UNICODETEXT      As Long = 13       ' Handle zu einem Unicode-String
'Public Const CF_ENHMETAFILE      As Long = 14       ' Handle zu einer Enhanced Metadatei
'Public Const CF_HDROP            As Long = 15       ' Liste von Dateihandles
'Public Const CF_LOCALE           As Long = 16       ' Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde
'Public Const CF_DIBV5            As Long = 17       ' Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)
'
'Public Const CF_JPEG             As Long = 19
'Public Const CF_TOPDOWN_DIB      As Long = 20
'
'Public Const CF_MULTI_TIFF       As Long = 22
'
'Public Const CF_OWNERDISPLAY     As Long = &H80&     ' benutzerdefinierter Anzeigetyp
'Public Const CF_PRIVATEFIRST     As Long = &H200&    ' privates Handle
'Public Const CF_PRIVATELAST      As Long = &H2FF&    ' privates Handle
'Public Const CF_GDIOBJFIRST      As Long = &H300&
'Public Const CF_GDIOBJLAST       As Long = &H3FF&
'
'Public Const CF_RTF              As Long = &HC09A&   ' 49306 Richt Text Format
'Public Const CF_HTML             As Long = &HC108&   ' 49416 HTML Format
'Public Const CF_XML              As Long = &HC308&   ' 49928 XML Spreadheet
'
'
'Public Const CF_DataObject       As Long = &HC009&  ' DataObject
'Public Const CF_FileName         As Long = &HC006&  ' Dateiname
'Public Const CF_FileNameW        As Long = &HC007&  ' Dateiname
'
'' da gibt es noch mehr Konstanten...
'
'' BitBlt dwRop-Konstante
'Public Const SRCCOPY             As Long = &HCC0020
'

Public Function Hex2(s As String) As String
    Hex2 = IIf(Len(s) = 1, "0", "") & s
End Function
Public Function Hex4(l As Long) As String
    Hex4 = Hex(l): If Len(Hex4) < 4 Then Hex4 = String(4 - Len(Hex4), "0") & Hex4
End Function


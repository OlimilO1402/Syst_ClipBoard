VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISubclassedWindow
'siehe auch:
'https://msdn.microsoft.com/de-de/library/windows/desktop/ms649016(v=vs.85).aspx#_win32_Removing_a_Window_from_the_Clipboard_Viewer_Chain
'https://msdn.microsoft.com/de-de/library/windows/desktop/ms649016(v=vs.85).aspx#_win32_Removing_a_Window_from_the_Clipboard_Viewer_Chain
'http://www.vbarchiv.net/workshop/workshop_25-die-zwischenablage-im-griff.html
Private Declare Function OpenClipboard Lib _
    "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumClipboardFormats Lib _
    "user32" (ByVal wFormat As Long) As Long
Private Declare Function CountClipboardFormats Lib _
    "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib _
    "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardData Lib _
    "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib _
    "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardFormatNameA Lib _
    "user32" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib _
    "user32" () As Long
Private Declare Function CloseClipboard Lib _
    "user32" () As Long
    
Private Const WM_CHANGECBCHAIN   As Long = &H30D
Private Const WM_DRAWCLIPBOARD As Long = &H308
Private Declare Function SetClipboardViewer Lib _
    "user32" (ByVal hWndNewViewer As Long) As Long
    
Private Declare Function ChangeClipboardChain Lib _
    "user32" (ByVal hWndRemove As Long, ByVal hWndNewNext As Long) As Long    'BOOL



Private Const WM_CLIPBOARDUPDATE As Long = &H31D
Private Declare Function AddClipboardFormatListener Lib _
    "user32" (ByVal hwnd As Long) As Long     'BOOL

Private Declare Function LcidToRfc1766A Lib _
    "mlang" (ByVal Locale As Long, ByVal pszRfc1766 As String, ByVal nChar As Long) As Long
                        
Private Declare Sub RtlMoveMemory Lib _
    "kernel32" (ByVal pDst As Long, ByVal pSrc As Long, ByVal blen As Long)

Private Declare Function GlobalAlloc Lib _
    "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib _
    "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib _
    "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib _
    "kernel32" (ByVal hMem As Long) As Long
'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366574(v=vs.85).aspx
'GlobalAlloc: Allocates the specified number of bytes from the heap.
'Note:
'The global functions have greater overhead and provide fewer features than other memory management functions.
'New applications should use the heap functions unless documentation states that a global function should be used.
'For more information, see Global and Local Functions.
'
'Global and Local Functions:
'https://msdn.microsoft.com/de-de/library/windows/desktop/aa366596(v=vs.85).aspx
'For example, some Windows functions allocate memory that must be freed with LocalFree, and the global
'functions are still used with Dynamic Data Exchange (DDE), the clipboard functions, and OLE data objects.

'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const GMEM_MOVEABLE As Long = &H2&
Public Enum ClipboardFormat
                           ' Das Objekt in der Zwischenablage ist ein . . .
    CF_TEXT = 1      ' = vbCFText   ' Handle auf einen unformatierten Ansi-Text (UTF-8).
    CF_BITMAP = 2    ' = vbCFBitmap ' Handle auf eine Windows Bitmap-Grafik   (Bitmap   = .BMP-Datei)
    CF_METAFILEPICT = 3             ' Handle auf eine Windows Metafile-Grafik (Metafile = .WMF-Datei)
    CF_SYLK = 4                     ' Handle auf einen Microsoft Symbolic-Link
    CF_DIF = 5                      ' "Software Arts' Data Interchange Format"
    CF_TIFF = 6                     ' Handle zu einem Tiff-Bitmap
    CF_OEMTEXT = 7                  ' Handle zu einem OEM-String
    CF_DIB = 8       ' = vbCFDIB    ' Handle zu einer %BITMAPINFO%-Struktur (Geräteunabhängige Bitmap)
    CF_BOTTOMUP_DIB = CF_DIB        '
    CF_PALETTE = 9  ' = vbCFPalette ' Handle zu einer Palette
    CF_PENDATA = 10                 ' sind Daten zu einem Microsoft Pen Extensions
    CF_RIFF = 11                    ' Handle zu einer Audiodatei
    CF_WAVE = 12                    ' Handle zu Wavedatei
    CF_UNICODETEXT = 13             ' Handle zu einem Unicode-String (UTF-16)
    CF_ENHMETAFILE = 14 ' = vbCFEMetafile ' Handle zu einer Enhanced Metadatei
    CF_HDROP = 15    ' = vbCFFiles  ' Liste von Dateihandles im Zusammenhang mit Drag-And-Drop
    CF_LOCALE = 16                  ' Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde
    CF_DIBV5 = 17                   ' Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)
    
    CF_JPEG = 19
    CF_TOPDOWN_DIB = 20
    
    CF_MULTI_TIFF = 22
    
    CF_OWNERDISPLAY = &H80&         '   128 ' benutzerdefinierter Anzeigetyp
    CF_DSPTEXT = &H81&              '   129 ' Text, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPBITMAP = &H82&            '   130 ' Bitmap, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_DSPMETAFILEPICT = &H83&      '   131 ' Zwischendatei, das Anzeigeformat ist mit einem eigenen Format verbunden
    CF_PRIVATEFIRST = &H200&        '   512 ' privates Handle
    CF_PRIVATELAST = &H2FF&         '   767 ' privates Handle
    CF_GDIOBJFIRST = &H300&         '   768 ' Dient in der Zwischenablage dem Setzen von privaten Datenformate
    CF_GDIOBJLAST = &H3FF&          '  1023 ' Dient in der Zwischenablage dem Setzen von privaten Datenformaten
    CF_FileName = &HC006&           ' 49158 ' Dateiname
    CF_FileNameW = &HC007&          ' 49159 ' Dateiname
    CF_DataObject = &HC009&         ' 49161 ' DataObject
    'von mir selber hinzugefügt, herausgefunden durch Excel-Zelle oder Word-Text in Zwischenablage:
    CF_XRTF = &HC09A&                ' 49306 ' Richt Text Format (Excel)
    CF_WRTF = &HC09D&                ' 49309 ' Richt Text Format (Word)
    
    CF_HTML_xls1 = &HC108&               ' 49416 ' HTML Format
    CF_HTML_xls2 = &HC12F&               ' 49455 ' HTML Format
    'da will mich bei Microsoft wohl jemand ärgern!!
    'jetzt hat sich plötzlich die Konstante geändert schon sehr merkwürdig
    'man braucht eine Funktion die zum Text "HTML" die Konstante raussucht!
    
    CF_PICTURE = &HC20A&             ' 49674 ' Handle auf ein Objekt vom Datentyp Picture
    CF_OBJECT = &HC215&              ' 49685 ' Handle auf ein beliebiges Objekt
    CF_XML = &HC308&                 ' 49928 ' XML Spreadheet
    
    CF_RTF = &HFFFFBF01 ' = vbCFRTF  ' -16639 Rich Text Format (.RTF-Datei).
    CF_Link = &HFFFFBF00             ' -16640 ' Informationen zur DDE-Verbindung.

' da gibt es noch mehr Konstanten...
End Enum

' BitBlt dwRop-Konstante
Private Const SRCCOPY            As Long = &HCC0020

Private m_CBFormats() As Long
Private m_hWnd As Long
Private m_hWndNextVw As Long
Public Event Changed(sender As cClipBoard)

Private Sub Class_Initialize()
    ReDim m_CBFormats(0)
End Sub

Public Sub Clear()
    OpenClipboard m_hWnd
    EmptyClipboard
    CloseClipboard
    ReDim m_CBFormats(0)
End Sub
Public Property Get IsClear() As Boolean
    IsClear = UBound(m_CBFormats) = 0
End Property
Public Sub ClearCBFormats()
    ReDim m_CBFormats(0)
End Sub

Public Property Get OwnerHwnd() As Long
    OwnerHwnd = m_hWnd
End Property
Public Property Let OwnerHwnd(ByVal Value As Long)
    m_hWnd = Value
End Property

Public Sub OwnerAsListener()
    m_hWndNextVw = SetClipboardViewer(m_hWnd)
End Sub
'Public Property Get hwnd() As Long
'    hwnd = m_hWnd
'End Property
'Public Property Let hwnd(aHWnd As Long)
'    m_hWnd = aHWnd
'End Property

Public Sub ReadClipboardFormats()
    If OpenClipboard(m_hWnd) Then
        Dim n As Long: n = CountClipboardFormats
        ReDim m_CBFormats(0 To n)
        Dim cbf As Long
        Dim i As Long
        Do
            i = i + 1
            ' Das nächste Format aus der Zwischenablage ermitteln
            cbf = EnumClipboardFormats(cbf)
            If cbf = 0 Then Exit Do
            m_CBFormats(i) = cbf
        Loop
        CloseClipboard
    End If
End Sub

Public Property Get CBFormats() As Long()
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    CBFormats = m_CBFormats
End Property


Public Function GetCBFormatForName(aName As String) As Long
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats
    Dim i As Long
    For i = 0 To UBound(m_CBFormats)
        Dim n1 As String: n1 = CBFormat_ToStr(m_CBFormats(i))
        Dim n2 As String: n2 = "*" & aName & "*"
        If n1 Like n2 Then
            GetCBFormatForName = m_CBFormats(i)
            Exit Function
        End If
    Next
End Function

Public Property Get HasFormat(aCBFormat As ClipboardFormat) As Boolean
    HasFormat = IsClipboardFormatAvailable(aCBFormat)
'    Dim i As Long
'    For i = 0 To UBound(m_CBFormats)
'        HasCBFormat = m_CBFormats(i) = aCBFormat
'        If HasCBFormat Then Exit Property
'    Next
End Property

Public Property Get ObjData(ByVal aCBFormat As Long) As Object
    If OpenClipboard(m_hWnd) Then
        Dim pMem As Long:         pMem = GetClipboardData(aCBFormat)
        If pMem Then
            If aCBFormat = ClipboardFormat.CF_BITMAP Then
                'Dim hCopy As Long
                'hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                Set ObjData = GetIPictureFromPtr(pMem)
            Else
            'Dim bLen As Long:     bLen = GlobalSize(pMem) '+ 2
            'If bLen Then
                'Dim pLok As Long: pLok = GlobalLock(pMem)
                'If pLok Then
                    Dim obj As Object
                    RtlMoveMemory VarPtr(obj), ByVal VarPtr(pMem), 4 'bLen
                    Set ObjData = obj
                    'GlobalUnlock pMem
                'End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Property Get ObjHandle(ByVal aCBFormat As Long) As Long 'Object
    If OpenClipboard(m_hWnd) Then
        Dim pMem As Long:         pMem = GetClipboardData(aCBFormat)
        If pMem Then
            ObjHandle = pMem
'            Dim bLen As Long:     bLen = GlobalSize(pMem) '+ 2
'            If bLen Then
'                Dim pLok As Long: pLok = GlobalLock(pMem)
'                If pLok Then
'                    'If aCBFormat = CF_LOCALE Then
'                        Dim hObj As Long
'                        RtlMoveMemory VarPtr(hObj), pLok, bLen
'                        ObjHandle = hObj
'                        'StrData = CStr(localeID) & " " & "&H" & Hex4(localeID) & " " & LcidToStr(localeID)
'                    Else
'                        Dim slen As Long: slen = IIf(aCBFormat = CF_UNICODETEXT, (bLen \ 2) - 1, bLen)
'                        StrData = Space$(slen)
'                        RtlMoveMemory StrPtr(StrData), pLok, bLen
'                        If aCBFormat <> CF_UNICODETEXT Then
'                            StrData = StrConv(StrData, vbUnicode)
'                        End If
'                    End If
'                    GlobalUnlock pMem
'                End If
'            End If
        End If
        CloseClipboard
    End If
End Property
Public Property Get StrData(ByVal aCBFormat As Long) As String
    If OpenClipboard(m_hWnd) Then
        Dim pMem As Long:         pMem = GetClipboardData(aCBFormat)
        If pMem Then
            Dim blen As Long:     blen = GlobalSize(pMem) '+ 2
            If blen Then
                Dim pLok As Long: pLok = GlobalLock(pMem)
                If pLok Then
                    If aCBFormat = CF_LOCALE Then
                        Dim localeID As Long
                        RtlMoveMemory VarPtr(localeID), pLok, blen
                        StrData = CStr(localeID) & " " & "&H" & Hex4(localeID) & " " & LcidToStr(localeID)
                    Else
                        Dim slen As Long: slen = IIf(aCBFormat = CF_UNICODETEXT, (blen \ 2) - 1, blen)
                        StrData = Space$(slen)
                        RtlMoveMemory StrPtr(StrData), pLok, blen
                        If aCBFormat <> CF_UNICODETEXT Then
                            StrData = StrConv(StrData, vbUnicode)
                        End If
                    End If
                    GlobalUnlock pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property
Private Function Hex4(l As Long) As String
    Hex4 = Hex(l): If Len(Hex4) < 4 Then Hex4 = String$(4 - Len(Hex4), "0") & Hex4
End Function
Private Function LcidToStr(ByVal aLCID As Long) As String
    Dim s As String: s = String$(6, vbNullChar)
    If LcidToRfc1766A(aLCID, s, 6) = 0 Then
        Dim i As Long
        For i = 0 To 1
            If Len(s) > 3 + i Then
                Mid$(s, 4 + i, 1) = UCase$(Mid$(s, 4 + i, 1))
            End If
        Next
    End If
    LcidToStr = s
End Function
Public Property Let StrData(ByVal aCBFormat As Long, Value As String)
    If OpenClipboard(m_hWnd) Then
        Dim blen As Long:         blen = LenB(Value) + 2
        If blen > 2 Then
            Dim pMem As Long:     pMem = GlobalAlloc(GMEM_MOVEABLE, blen)
            If pMem Then
                Dim pLok As Long: pLok = GlobalLock(pMem)
                If pLok Then
                    RtlMoveMemory ByVal pLok, ByVal StrPtr(Value), blen
                    GlobalUnlock pMem
                    SetClipboardData aCBFormat, pMem
                End If
            End If
        End If
        CloseClipboard
    End If
End Property

Public Function CBFormat_ToStr(aCBFormat As Long) As String
    Dim s  As String: s = CLng(aCBFormat) & ", &H" & Hex(aCBFormat)
    Dim s2 As String: s2 = Space(256)
    Dim rv As Long: rv = GetClipboardFormatNameA(aCBFormat, s2, 256)
    s2 = Trim(s2)
    If Len(s2) Then
        s = s & "(api): " & s2
    Else
        Select Case aCBFormat
        Case 0:
        Case CF_TEXT:         s = s & ": Handle auf einen unformatierten Ansi-Text (UTF-8)"
        Case CF_BITMAP:       s = s & ": Handle auf eine Windows Bitmap-Grafik"
        Case CF_METAFILEPICT: s = s & ": Handle auf eine Windows Metafile-Grafik"
        Case CF_SYLK:         s = s & ": Handle auf einen Microsoft Symbolic-Link"
        Case CF_DIF:          s = s & ": Software Arts' Data Interchange Format"
        Case CF_TIFF:         s = s & ": Handle auf ein Tiff-Bitmap"
        Case CF_OEMTEXT:      s = s & ": Handle auf einen OEM-String"
        Case CF_DIB:          s = s & ": Handle auf eine %BITMAPINFO%-Struktur"
        Case CF_BOTTOMUP_DIB: s = s & ": CF_DIB"
        Case CF_PALETTE:      s = s & ": Handle auf eine Palette"
        Case CF_PENDATA:      s = s & ": Handle auf Daten zu Microsoft Pen Extensions"
        Case CF_RIFF:         s = s & ": Handle auf eine Audiodatei im RIFF-Wave-format"
        Case CF_WAVE:         s = s & ": Handle auf eine Wavedatei"
        Case CF_UNICODETEXT:  s = s & ": Handle auf einen Unicode-String (UTF-16)"
        Case CF_ENHMETAFILE:  s = s & ": Handle auf eine Enhanced-Metadatei"
        Case CF_HDROP:        s = s & ": Liste von Dateihandles im Zusammenhang mit Drag-And-Drop"
        Case CF_LOCALE:       s = s & ": Sprach-ID, die für Text-Strings in der Zwischenablage benutzt wurde"
        Case CF_DIBV5:        s = s & ": Handle zu einer %BITMAPV5HEADER%-Struktur (Win 2000/XP)"
    
        Case CF_JPEG:         s = s & ": CF_JPEG"
        Case CF_TOPDOWN_DIB:  s = s & ": CF_TOPDOWN_DIB"
    
        Case CF_MULTI_TIFF:   s = s & ": CF_MULTI_TIFF"
    
        Case CF_OWNERDISPLAY: s = s & ": benutzerdefinierter Anzeigetyp"
        Case CF_DSPTEXT:      s = s & ": Text, das Anzeigeformat ist mit einem eigenen Format verbunden"
        Case CF_DSPBITMAP:    s = s & ": Bitmap, das Anzeigeformat ist mit einem eigenen Format verbunden"
        Case CF_DSPMETAFILEPICT: s = s & ": Zwischendatei, das Anzeigeformat ist mit einem eigenen Format verbunden"
        Case CF_PRIVATEFIRST: s = s & ": CF_PRIVATEFIRST privates Handle"
        Case CF_PRIVATELAST:  s = s & ": CF_PRIVATELAST privates Handle"
        Case CF_GDIOBJFIRST:  s = s & ": CF_GDIOBJFIRST"
        Case CF_GDIOBJLAST:   s = s & ": CF_GDIOBJLAST"
    
        Case CF_DataObject:   s = s & ": CF_DataObject DataObject"
        Case CF_FileName:     s = s & ": CF_FileName   Dateiname"
        Case CF_FileNameW:    s = s & ": CF_FileNameW  Dateiname"
        Case Else:            s = s & ": unbekanntes Format"
        End Select
    End If
    CBFormat_ToStr = s
End Function

'eine Excel-Zelle in die Zwischenablage kopiert liefert folgende Formate:
'lngFormat = 49161
'lngFormat = 14
'lngFormat = 3
'lngFormat = 2
'lngFormat = 49820
'lngFormat = 49804
'lngFormat = 49806
'lngFormat = 4
'lngFormat = 5
'lngFormat = 49819
'lngFormat = 49439
'lngFormat = 13
'lngFormat = 1
'lngFormat = 49807
'lngFormat = 49327
'lngFormat = 49163
'lngFormat = 49156
'lngFormat = 49155
'lngFormat = 49166
'lngFormat = 49165
'lngFormat = 49167
'lngFormat = 49785
'lngFormat = 129
'lngFormat = 49154
'lngFormat = 49171
'lngFormat = 16
'lngFormat = 7
'lngFormat = 8
'lngFormat = 17

'oder sortiert:
'lngFormat = 1
'lngFormat = 2
'lngFormat = 3
'lngFormat = 4
'lngFormat = 5
'lngFormat = 7
'lngFormat = 8
'lngFormat = 13
'lngFormat = 14
'lngFormat = 16
'lngFormat = 17
'lngFormat = 129

'lngFormat = 49154
'lngFormat = 49155
'lngFormat = 49156
'lngFormat = 49161
'lngFormat = 49163
'lngFormat = 49165
'lngFormat = 49166
'lngFormat = 49167
'lngFormat = 49171

'lngFormat = 49327
'lngFormat = 49439
'lngFormat = 49785

'lngFormat = 49804
'lngFormat = 49806
'lngFormat = 49807

'lngFormat = 49819
'lngFormat = 49820


Private Function ISubclassedWindow_HandleMessage(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
    Select Case uMsg
    Case WM_CHANGECBCHAIN:
        '
    Case WM_DRAWCLIPBOARD:
        RaiseEvent Changed(Me)
    End Select
End Function

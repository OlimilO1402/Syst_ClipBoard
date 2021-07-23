Attribute VB_Name = "MCBFormat"
Option Explicit

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
    
    CF_HTML_xls0 = &HC0E2&           ' 49378 ' Excel(2016) HTML Format
    CF_HTML_xls1 = &HC108&           ' 49416 ' Excel(2019) HTML Format
    CF_HTML_xls2 = &HC12F&           ' 49455 ' Excel(2019) HTML Format
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

#If VBA7 Then
    Private Declare PtrSafe Function GetClipboardFormatNameW Lib "user32" (ByVal wFormat As Long, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
#Else
    Private Declare Function GetClipboardFormatNameW Lib "user32" (ByVal wFormat As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
#End If

Public Function CBFormat_ToStr(aCBFormat As Long) As String
    Dim s  As String: s = CLng(aCBFormat) & ", &H" & Hex(aCBFormat)
    Dim s2 As String: s2 = Space(256)
    Dim rv As Long: rv = GetClipboardFormatNameW(aCBFormat, StrPtr(s2), 512)
    s2 = Trim(s2)
    If Len(s2) Then
        If Right(s2, 1) = vbNullChar Then s2 = Left(s2, Len(s2) - 1)
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


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
'Implements ISubclassedWindow
'siehe auch:
'https://msdn.microsoft.com/de-de/library/windows/desktop/ms649016(v=vs.85).aspx#_win32_Removing_a_Window_from_the_Clipboard_Viewer_Chain
'https://msdn.microsoft.com/de-de/library/windows/desktop/ms649016(v=vs.85).aspx#_win32_Removing_a_Window_from_the_Clipboard_Viewer_Chain
'http://www.vbarchiv.net/workshop/workshop_25-die-zwischenablage-im-griff.html

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function CountClipboardFormats Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardFormatNameA Lib "user32" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal bLen As Long)

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function SetClipboardViewer Lib "user32" (ByVal hWndNewViewer As LongPtr) As Long
Private Declare PtrSafe Function ChangeClipboardChain Lib "user32" (ByVal hWndRemove As LongPtr, ByVal hWndNewNext As LongPtr) As Long        'BOOL
Private Declare PtrSafe Function AddClipboardFormatListener Lib "user32" (ByVal hWnd As LongPtr) As Long         'BOOL
Private Declare PtrSafe Function LcidToRfc1766A Lib "mlang" (ByVal Locale As Long, ByVal pszRfc1766 As String, ByVal nChar As Long) As Long
#Else
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardFormatNameA Lib "user32" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
   
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal pDst As Long, ByVal pSrc As Long, ByVal blen As Long)

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWndNewViewer As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWndRemove As Long, ByVal hWndNewNext As Long) As Long        'BOOL
Private Declare Function AddClipboardFormatListener Lib "user32" (ByVal hWnd As Long) As Long         'BOOL
Private Declare Function LcidToRfc1766A Lib "mlang" (ByVal Locale As Long, ByVal pszRfc1766 As String, ByVal nChar As Long) As Long
#End If
    
Private Const WM_CHANGECBCHAIN   As Long = &H30D
Private Const WM_DRAWCLIPBOARD   As Long = &H308
Private Const WM_CLIPBOARDUPDATE As Long = &H31D

Private Const GMEM_MOVEABLE      As Long = &H2&
' BitBlt dwRop-Konstante
Private Const SRCCOPY            As Long = &HCC0020

                        
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

Private m_hWnd        As LongPtr
Private m_CBFormats() As Long
Private m_hWndNextVw  As LongPtr
Public Event Changed(Sender As cClipBoard)

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

Public Property Get OwnerHwnd() As LongPtr
    OwnerHwnd = m_hWnd
End Property
Public Property Let OwnerHwnd(ByVal Value As LongPtr)
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

Public Property Get ObjHandle(ByVal aCBFormat As Long) As LongPtr 'Object
    If OpenClipboard(m_hWnd) Then
        Dim pMem As LongPtr:       pMem = GetClipboardData(aCBFormat)
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

Private Function Hex4(L As Long) As String
    Hex4 = Hex(L): If Len(Hex4) < 4 Then Hex4 = String$(4 - Len(Hex4), "0") & Hex4
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

Public Property Get cbElements() As cbElements
    If UBound(m_CBFormats) < 1 Then ReadClipboardFormats

    Set cbElements = New cbElements
    Dim aCBFormat As ClipboardFormat
    If OpenClipboard(m_hWnd) Then
        Dim i As Long
        For i = 0 To UBound(m_CBFormats)
            aCBFormat = m_CBFormats(i)
            Dim pMem As Long:         pMem = GetClipboardData(aCBFormat)
            If pMem Then
                Dim blen As Long:     blen = GlobalSize(pMem) '+ 2
                If blen Then
                    Dim pLok As Long: pLok = GlobalLock(pMem)
                    If pLok Then
                        ReDim barr(blen) As Byte
                        RtlMoveMemory VarPtr(barr(0)), pLok, blen
                        GlobalUnlock pMem
                        cbElements.Add MNew.CBElement(aCBFormat, barr)
                    End If
                End If
            End If
        Next
    End If
    CloseClipboard
End Property

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


'Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
'    Select Case uMsg
'    Case WM_CHANGECBCHAIN:
'        '
'    Case WM_DRAWCLIPBOARD:
'        RaiseEvent Changed(Me)
'    End Select
'End Function


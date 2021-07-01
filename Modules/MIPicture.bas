Attribute VB_Name = "MIPicture"
Option Explicit

' ACHTUNG : Einbindung der Excel-Objektbibliotek erforderlich (Extras-Verweise)
'           Einbindung der OLE-Automation erforderlich  (Extras-Verweise)
' getestet unter AC97 und AC2000

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Declare a UDT to store the bitmap information
Private Type PictureDescription
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

'''Windows API Function Declarations

'Does the clipboard contain a bitmap/metafile?
'Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long

'Open the clipboard to read
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

'Get a pointer to the bitmap/metafile
'Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long

'Close the clipboard
'Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function StringFromGUID2 Lib _
    "ole32" (rguid As GUID, ByVal lpsz As String, ByVal cchMax As Long) As Long

'Convert the handle into an OLE IPicture interface.
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PictureDescription, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long

'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

'The API format types we're interested in
'Const CF_BITMAP = 2
'Const CF_PALETTE = 9
'Const CF_ENHMETAFILE = 14

Const IMAGE_BITMAP     As Long = 0
Const LR_COPYRETURNORG As Long = &H4

'Function PastePicture(Optional lXlPicType As Long = xlPicture) As IPicture
'
'    'Some pointers
'    'Dim hPal As Long
'
'    'Convert the type of picture requested from the xl constant to the API constant
'    Dim lPicType  As Long:  lPicType = IIf(lXlPicType = xlBitmap, CF_BITMAP, CF_ENHMETAFILE)
'
'
'    'Check if the clipboard contains the required format
'    Dim hPicAvail As Long: hPicAvail = IsClipboardFormatAvailable(lPicType)
'
'    If hPicAvail <> 0 Then
'        'Get access to the clipboard
'        Dim h As Long: h = OpenClipboard(0&)
'
'        If h > 0 Then
'            'Get a handle to the image data
'            Dim hPtr As Long: hPtr = GetClipboardData(lPicType)
'
'            'Create our own copy of the image on the clipboard, in the appropriate format.
'            Dim hCopy As Long
'            If lPicType = CF_BITMAP Then
'                hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
'            Else
'                hCopy = CopyEnhMetaFile(hPtr, vbNullString)
'            End If
'
'            'Release the clipboard to other programs
'            h = CloseClipboard
'
'            'If we got a handle to the image, convert it into a Picture object and return it
'            If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, lPicType)
'        End If
'    End If
'
'End Function

Public Function GetIPictureFromPtr(pMem As Long) As IPictureDisp
    If pMem = 0 Then Exit Function
    Dim hPic    As Long:                  hPic = CopyImage(pMem, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    Dim IDispat As GUID:               IDispat = GetGUID_IDispatch
    Dim PicDesc As PictureDescription: PicDesc = GetPicDescBmp(hPic)
    Dim IPic    As IPicture
    Dim hr      As Long:                    hr = OleCreatePictureIndirect(PicDesc, IDispat, True, IPic)
    Set GetIPictureFromPtr = IPic
End Function
Private Function GetGUID_IDispatch() As GUID
    With GetGUID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA:
        .Data4(4) = &H0:  .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB:
    End With
    Dim s As String: s = Space(80)
    Dim hr As Long
    hr = StringFromGUID2(GetGUID_IDispatch, s, 80)
    Debug.Print StrConv(Trim(s), vbFromUnicode)
End Function
Private Function GetPicDescBmp(ByVal hPic As Long) As PictureDescription
    Const PICTYPE_BITMAP      As Long = 1
    With GetPicDescBmp
        .Size = LenB(GetPicDescBmp)
        .Type = PICTYPE_BITMAP
        .hPic = hPic
        .hPal = 0
    End With
End Function
Private Function GetPicDescEwmf(ByVal hPic As Long) As PictureDescription
    Const PICTYPE_ENHMETAFILE As Long = 4
    With GetPicDescEwmf
        .Size = LenB(GetPicDescEwmf)
        .Type = PICTYPE_ENHMETAFILE
        .hPic = hPic
        .hPal = 0
    End With
End Function


'Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
'
'    ' IPicture requires a reference to "OLE Automation"
'    Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
'
'    'OLE Picture types
'    Const PICTYPE_BITMAP      As Long = 1
'    Const PICTYPE_ENHMETAFILE As Long = 4
'
'    ' Create the Interface GUID (for the IPicture interface)
'    With IID_IDispatch
'        .Data1 = &H7BF80980
'        .Data2 = &HBF32
'        .Data3 = &H101A
'        .Data4(0) = &H8B
'        .Data4(1) = &HBB
'        .Data4(2) = &H0
'        .Data4(3) = &HAA
'        .Data4(4) = &H0
'        .Data4(5) = &H30
'        .Data4(6) = &HC
'        .Data4(7) = &HAB
'    End With
'
'    ' Fill uPicInfo with necessary parts.
'    With uPicInfo
'        .Size = Len(uPicInfo)                                                   ' Length of structure.
'        .Type = IIf(lPicType = CF_BITMAP, PICTYPE_BITMAP, PICTYPE_ENHMETAFILE)  ' Type of Picture
'        .hPic = hPic                                                            ' Handle to image.
'        .hPal = IIf(lPicType = CF_BITMAP, hPal, 0)                              ' Handle to palette (if bitmap).
'    End With
'
'    ' Create the Picture object.
'    r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)
'
'    ' If an error occured, show the description
'    If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)
'
'    ' Return the new Picture object.
'    Set CreatePicture = IPic
'
'End Function
Private Function fnOLEError(lErrNum As Long) As String
    
    'OLECreatePictureIndirect return values
    Const E_ABORT = &H80004004
    Const E_ACCESSDENIED = &H80070005
    Const E_FAIL = &H80004005
    Const E_HANDLE = &H80070006
    Const E_INVALIDARG = &H80070057
    Const E_NOINTERFACE = &H80004002
    Const E_NOTIMPL = &H80004001
    Const E_OUTOFMEMORY = &H8007000E
    Const E_POINTER = &H80004003
    Const E_UNEXPECTED = &H8000FFFF
    Const S_OK = &H0
    
    Select Case lErrNum
    Case E_ABORT:        fnOLEError = " Aborted"
    Case E_ACCESSDENIED: fnOLEError = " Access Denied"
    Case E_FAIL:         fnOLEError = " General Failure"
    Case E_HANDLE:       fnOLEError = " Bad/Missing Handle"
    Case E_INVALIDARG:   fnOLEError = " Invalid Argument"
    Case E_NOINTERFACE:  fnOLEError = " No Interface"
    Case E_NOTIMPL:      fnOLEError = " Not Implemented"
    Case E_OUTOFMEMORY:  fnOLEError = " Out of Memory"
    Case E_POINTER:      fnOLEError = " Invalid Pointer"
    Case E_UNEXPECTED:   fnOLEError = " Unknown Error"
    Case S_OK:           fnOLEError = " Success!"
    End Select

End Function

'Function GrafikZwischenablage2Datei(DatName As String) As Boolean
'    Dim lPicType As Long:    lPicType = xlBitmap
'
'    Dim oPic     As Variant: Set oPic = PastePicture(lPicType)
'    If oPic Is Nothing Then Exit Function
'    SavePicture oPic, DatName
'    GrafikZwischenablage2Datei = True
'End Function
'
'Sub mytest()
'    Dim ZielDat As String: ZielDat = "d:\tar\mytest.bmp"
'    If GrafikZwischenablage2Datei(ZielDat) Then
'        MsgBox ZielDat & " erzeugt!"
'    Else
'        MsgBox "kein bild in der zwischenablage"
'    End If
'End Sub
'

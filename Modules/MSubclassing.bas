Attribute VB_Name = "MSubclassing"
'--- Anfang Modul "basSubclassing" alias basSubclassing.bas ---
Option Explicit

Public Enum EnumSubclassID
    escidFrmMain = 1
    'escidFrmMainCmdOk
    '...
End Enum

Private Declare Function SetWindowSubclass Lib _
    "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Public Declare Function DefSubclassProc Lib _
    "comctl32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveWindowSubclass Lib _
    "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long

Private Declare Sub RtlZeroMemory Lib _
    "kernel32" (ByVal pDest As Long, ByVal sz As Long)
Public Declare Sub RtlMoveMemory Lib _
    "kernel32" (ByVal pDest As Long, ByVal pSrc As Long, ByVal blen As Long)


Public Function SubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Dim bCallDefProc As Boolean: bCallDefProc = True
    Dim lRet As Long
    
Try: On Error GoTo Catch
    
    If dwRefData Then
        Dim oClient As ISubclassedWindow
        Set oClient = GetObjectFromPointer(dwRefData)
        If Not (oClient Is Nothing) Then
            lRet = oClient.HandleMessage(hWnd, uMsg, wParam, lParam, uIdSubclass, bCallDefProc)
        End If
    End If
    
Catch:
    If Err Then
        Debug.Print "Error in SubclassProc: ", Err.Number, Err.Description
    End If
Finally:
    On Error Resume Next
    If bCallDefProc Then
        lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
    SubclassProc = lRet
End Function

Public Function SubclassWindow(ByVal hWnd As Long, oClient As ISubclassedWindow, ByVal eSubclassID As EnumSubclassID) As Boolean
    Dim bRet As Boolean
Try: On Error GoTo Catch
    bRet = SetWindowSubclass(hWnd, AddressOf MSubclassing.SubclassProc, eSubclassID, ObjPtr(oClient)) <> 0
Catch:
    If Err Then
        Debug.Print "Error in SubclassWindow: ", Err.Number, Err.Description
        bRet = False
    End If
Finally:
    SubclassWindow = bRet
End Function

Public Function UnSubclassWindow(ByVal hWnd As Long, ByVal eSubclassID As EnumSubclassID) As Boolean
    Dim bRet As Boolean
Try: On Error GoTo Catch
    bRet = RemoveWindowSubclass(hWnd, AddressOf MSubclassing.SubclassProc, eSubclassID) <> 0
Catch:
    If Err Then
        Debug.Print "Error in UnSubclassWindow: ", Err.Number, Err.Description
        bRet = False
    End If
Finally:
    UnSubclassWindow = bRet
End Function

' returns the object <lPtr> points to
Private Function GetObjectFromPointer(ByVal lPtr As Long) As Object
    Dim oRet As Object
    ' get the object <lPtr> points to
    RtlMoveMemory VarPtr(oRet), VarPtr(lPtr), LenB(lPtr)
    Set GetObjectFromPointer = oRet
    ' free memory
    RtlZeroMemory VarPtr(oRet), 4
End Function
'--- Ende Modul "basSubclassing" alias basSubclassing.bas ---

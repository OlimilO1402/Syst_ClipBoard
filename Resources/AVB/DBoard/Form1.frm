VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "DBoard 2.0"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4800
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Text speichern"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton butLoad 
      Caption         =   "TextLaden"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Inhalt"
      Height          =   2415
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Entferne Eintrag"
      Enabled         =   0   'False
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Programm schließen"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ListBox LB 
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer_Check_Clipboard 
      Enabled         =   0   'False
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clipboard"
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton butNew 
         Caption         =   "Neuer Eintrag"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
      End
      Begin VB.ListBox uLB 
         Height          =   2985
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCB 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNoClip 
         Caption         =   "Keine Clipboard Einträge vorhanden"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Menu schließen"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SearchLB Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const LB_FINDSTRINGEXACT As Long = &H1A2

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Dim t As NOTIFYICONDATA
Dim cCount As Integer, uIndex As Integer
Dim noADDstr As String
Private Sub butLoad_Click()
    Dim File As String
    
    CDialog.DialogTitle = "Textdatei öffnen"
    CDialog.Filter = "Text (*.txt)|*.txt"
    CDialog.ShowOpen
                
    File = CDialog.FileName

    If File <> vbNullString Then Call ReadTXTfile(File)
    
End Sub
Private Sub butNew_Click()
    
    uLB.AddItem "", 0
    LB.AddItem "", 0
    
    uLB.ListIndex = 0
    
    Call WriteMenus
    
    noADDstr = ""
    Clipboard.Clear
    
    Text1.SetFocus
    
End Sub
Private Sub butSave_Click()
    Call SaveTXTfile(Text1.Text)
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Dim X As Integer
    
    X = uLB.ListIndex
    
    If X > -1 Then
        Text1.Text = ""

        uLB.RemoveItem X
        If X < 10 Then LB.RemoveItem X
    
        uLB.ListIndex = -1
    
        Call WriteMenus
    End If
    
End Sub
Private Sub Command3_Click()
    Dim R$
    
    Me.Hide
    
    R = "Dieses Programm ist entstanden, weil in ActiveVB " & _
        "ein Programm namens SuperClipBoard geuppt worden ist." & _
        "Ich wollte zeigen, das man ein wesentlich besseres " & _
        "Programm mit weniger Code zu diesem Thema schreiben kann." & _
        vbCrLf & vbCrLf & _
        "Wolfgang Ehrhardt" & vbCrLf & _
        "woeh@gmx.de"
        
    Call MsgBox(R, vbOKOnly + vbInformation, "Info")
    
    Me.Show
    
End Sub

Private Sub Form_Load()
    
    Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)

    Call init_Timer
    Call init_TryIcon
    Call WriteMenus
    
    LB.Clear
    uLB.Clear
    
    cCount = 0
    
End Sub
Private Sub init_TryIcon()

    Me.Hide
    App.TaskVisible = False

    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
    t.szTip = "DBoard (RMB = Menu / LMB DBoarddialog" & Chr$(0)
    
    Shell_NotifyIcon NIM_ADD, t
    
End Sub
Private Sub init_Timer()
    Timer_Check_Clipboard.Interval = 1000
    Timer_Check_Clipboard.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = pichook.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub
Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then _
        Me.Hide: _
        uLB.ListIndex = -1: _
        Text1.Text = ""
End Sub
Private Sub mnuCB_Click(Index As Integer)
    
    noADDstr = LB.List(Index - 1)
    'noADDstr = mnuCB(Index).Caption
    
    Clipboard.Clear
    Clipboard.SetText noADDstr, vbCFText
    
End Sub
Private Sub Text1_Change()
    
    uLB.List(uIndex) = Text1.Text
    If uIndex < 10 Then LB.List(uIndex) = Text1.Text
    
    If Text1.Text = "" Then
        butSave.Enabled = False
    Else
        butSave.Enabled = True
    End If
    
    noADDstr = Text1.Text
    
    Clipboard.Clear
    Clipboard.SetText Text1.Text
    
    Call WriteMenus
    
End Sub
Private Sub Timer_Check_Clipboard_Timer()
    Dim CB As String
    Dim P As Integer
    Dim L As Long
    Dim R$
    
    Static oCB As String
    
    CB = Clipboard.GetText
    
    If (Not (noADDstr <> "" And (noADDstr = CB))) _
    And oCB <> CB Then
        P = SearchLB(uLB.hwnd, LB_FINDSTRINGEXACT, -1, CB)

        If P > -1 Then _
            LB.RemoveItem P: _
            uLB.RemoveItem P
        
        cCount = cCount + 1
        
        LB.AddItem CB, 0
        uLB.AddItem CB, 0
        
        noADDstr = CB
        
        Do While LB.ListCount - 1 > 9
            LB.RemoveItem 10
        Loop
        
        Call WriteMenus
        
        oCB = CB
    End If
    
End Sub
Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Static Rec As Boolean, Msg As Long
    
    Msg = X / Screen.TwipsPerPixelX
    
    If Not Rec Then
        Rec = True
        
        Select Case Msg
            Case WM_LBUTTONDBLCLK
            Case WM_LBUTTONUP
                Me.WindowState = vbNormal
                Me.Visible = Not Me.Visible
            Case WM_RBUTTONDBLCLK
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
                Me.PopupMenu mnu
        End Select
        
        Rec = False
    End If
End Sub
Private Sub uLB_Click()
    Dim R$
    
    uIndex = uLB.ListIndex
    R$ = uLB.List(uIndex)
    
    uLB.ToolTipText = R$
    Text1.Text = R$
    
    Command2.Enabled = True
    butLoad.Enabled = True
    
    If Me.Visible Then Text1.SetFocus
    
End Sub
Private Sub WriteMenus()
    Dim P As Integer
    Dim L As Long
    Dim R$
    
    For P = 1 To 10
        mnuCB(P).Visible = False
    Next P
        
    For P = 0 To LB.ListCount - 1
        mnuCB(P + 1).Visible = True
        
        R$ = LB.List(P)
            
        L = Len(R$)
        If L > 25 Then L = 25
            
        mnuCB(P + 1).Caption = Left(R$, L)
    Next P
        
    If LB.ListCount = 0 Then
        mnuNoClip.Visible = True
    Else
        mnuNoClip.Visible = False
    End If
    
        
End Sub
Private Function Read_TextFile(Textfile As String) As String
    Dim NextLine As String

    On Local Error Resume Next

    Do Until EOF(Textfile)
        Line Input #FileNum, NextLine
        Read_TextFile = Read_TextFile & NextLine
    Loop

End Function
Private Sub ReadTXTfile(File As String)
    Dim TXT As String, Text As String
    Dim nFile As Integer
    
    On Local Error Resume Next
    
    nFile = FreeFile
    
    Open File For Input As #nFile
        Do While Not EOF(nFile)
            Line Input #nFile, TXT
            Text = Text & TXT & vbCrLf
        Loop
    Close #nFile
    
    Text = Mid(Text, 1, Len(Text) - 2)
    
    Text1.Text = Text
    
End Sub
Private Sub SaveTXTfile(TXT As String)
    Dim nFile As Integer
    Dim File As String

    On Error Resume Next

    nFile = FreeFile

    CDialog.DialogTitle = "Text speichern"
    CDialog.FilterIndex = 1
    CDialog.ShowSave

    File = CDialog.FileName
        
    If File <> vbNullString Then
        Open File For Output As #nFile
            Print #nFile, TXT
        Close #nFile
    End If

End Sub

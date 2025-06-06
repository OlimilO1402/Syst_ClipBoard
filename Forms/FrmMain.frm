VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FMain 
   Caption         =   "ClipBoard"
   ClientHeight    =   10965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15600
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10965
   ScaleWidth      =   15600
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnCopy 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   13575
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   1080
      Width           =   13575
      Begin VB.PictureBox Panel2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   9735
         Left            =   4560
         ScaleHeight     =   9735
         ScaleWidth      =   8895
         TabIndex        =   3
         Top             =   0
         Width           =   8895
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   2175
            Left            =   0
            TabIndex        =   11
            Top             =   4920
            Width           =   8175
            ExtentX         =   14420
            ExtentY         =   3836
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Beides
            TabIndex        =   5
            Top             =   7200
            Width           =   8295
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   2415
            Left            =   0
            ScaleHeight     =   157
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   549
            TabIndex        =   4
            Top             =   0
            Width           =   8295
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   2295
            Left            =   0
            TabIndex        =   6
            Top             =   2520
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4048
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   3
            TextRTF         =   $"FrmMain.frx":1782
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9735
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.CommandButton BtnGetClipBoardConstants 
      Caption         =   "GetClipBoardConstants"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents ClipBoard1 As cClipBoard
Attribute ClipBoard1.VB_VarHelpID = -1

Private WithEvents Splitter1 As Splitter
Attribute Splitter1.VB_VarHelpID = -1
' f�r RichTextBox1:
' * Komponente: Microsoft Rich Textbox Control 6.0 (SP3)
'
' f�r WebBrowser1:
' * Komponente: Microsoft Internet Controls
' * Verweis   : Microsoft HTML Object Library
Private m_CBHistory  As CBHistory
Private m_CBElements As cbElements

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

'Private Sub Command1_Click()
'    Set m_CBElements = m_CBHistory.Add(ClipBoard1.CBElements)
'    MsgBox m_CBElements.Count
'End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set Splitter1 = New Splitter
    Splitter1.New_ False, Me, Panel1, "Splitter1", List1, Panel2
    With Splitter1
        .LeftTopPos = List1.Width
        .BorderStyle = bsXPStyl
    End With
    
    Set ClipBoard1 = New cClipBoard
    ClipBoard1.OwnerHwnd = Me.hWnd
    BtnClear_Click
    Set m_CBHistory = New CBHistory
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    T = Panel1.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then
        Panel1.Move L, T, W, H
    End If
End Sub

Private Sub Panel2_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    W = Panel2.Width
    H = Panel2.Height / 4
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
    T = T + H
    If W > 0 And H > 0 Then RichTextBox1.Move L, T, W, H
    T = T + H
    If W > 0 And H > 0 Then WebBrowser1.Move L, T, W, H
    T = T + H
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub ClipBoard1_Changed(Sender As cClipBoard)
    '
End Sub

Private Sub Splitter1_OnMove(Sender As Splitter)
    '
End Sub

Private Sub BtnGetClipBoardConstants_Click()
    ClipBoard1.ClearCBFormats
    Set m_CBElements = m_CBHistory.Add(ClipBoard1.cbElements)
    'MsgBox m_CBElements.Count
    List1.Clear
    Dim cbe As CBElement
    For Each cbe In m_CBElements.List
        List1.AddItem cbe.ToStr
    Next

'    ClipBoard1.ClearCBFormats
'    List1.Clear
'    Dim cbf() As ClipboardFormat: cbf = ClipBoard1.CBFormats
'    Dim i As Long
'    For i = 1 To UBound(cbf)
'        List1.AddItem ClipBoard1.CBFormat_ToStr(cbf(i))
'    Next
End Sub

Private Sub BtnClear_Click()
    'ClipBoard1.Clear
    List1.Clear
    Text1.Text = ""
    Text2.Text = ""
    Set Picture1.Picture = Nothing
    Picture1.Cls
    RichTextBox1.TextRTF = ""
    WebBrowser1.Navigate2 "about:blank"
End Sub
Private Sub BtnCopy_Click()
    'es soll der in der Liste ausgew�hlte Eintrag allein wieder in die zwischenablage kopiert werden
    'also zuerst aus der ZwAbl lesen dann allein wieder reinkopieren
    'hmm eigentlich br�ucht es eine Klasse ClipBoardEntry oder ClipBoardElement
    'das hat nur eine Eigenschaft ClipBoardFormat und ein ByteArray das s�mtliche Daten beinhaltet
    
    Dim s As String: s = Text2.Text
    Dim cf As Long: cf = CBFormat_Parse(s)
    If cf = 0 Then
        MsgBox "Please select entry from list"
        Exit Sub
    End If
    's = ClipBoard1.StrData(ClipboardFormat.CF_HTML_xls0)
    's = Trim(s)
    'If Len(s) = 0 Then
    s = ClipBoard1.StrData(cf)
    'End If
    s = ExtractHTML(s)
    'Debug.Print s
    'm_CB.StrData(Module1.CF_HTML) = s
    ClipBoard1.Clear
    'ClipBoard1.StrData(cf) = s
    ClipBoard1.StrData(CF_UNICODETEXT) = s
    'WebBrowser1.Document.body.innerHTML = s
End Sub

Private Function CBFormat_Parse(s As String) As ClipboardFormat
    If Len(s) = 0 Then Exit Function
    If InStr(1, s, ",") = 0 Then Exit Function
    Dim sa() As String: sa = Split(s, ",")
    Dim L As Long: L = CLng(sa(0))
    CBFormat_Parse = L
End Function
Private Sub List1_Click()
    If List1.ListIndex < 0 Then Exit Sub
    Text2.Text = List1.List(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
    If List1.ListIndex < 0 Then Exit Sub
    Dim cf As ClipboardFormat:   cf = CLng(Split(List1.List(List1.ListIndex), ", &H")(0))
    Select Case cf
    Case CF_ENHMETAFILE, CF_METAFILEPICT
        Dim hWMF As Long: hWMF = ClipBoard1.ObjHandle(cf)
        If cf = CF_METAFILEPICT Then
            Mewmf.PlayMetaFile Picture1.hDC, hWMF
        ElseIf cf = CF_ENHMETAFILE Then
            Dim H As ENHMETAHEADER
            GetEnhMetaFileHeader hWMF, LenB(H), H
            Dim r As RECT: r.Left = 0: r.Top = 0: r.Right = 200: r.Bottom = 200
            Mewmf.PlayEnhMetaFile Picture1.hDC, hWMF, r
        End If
        'MsgBox Err.LastDllError
        'MsgBox Module1.ErrMessage(Err.LastDllError)
    Case ClipboardFormat.CF_PICTURE, ClipboardFormat.CF_BITMAP
    
        'Dim hPic As Long: hPic = m_CB.ObjHandle(CF_PICTURE)
        'Picture1.Picture.Handle = hPic
        'Set Picture1.Picture = m_CB.ObjData(cf)
        Set Picture1.Picture = ClipBoard1.ObjData(cf)
    Case Else
        Dim s  As String:  s = ClipBoard1.StrData(cf)
        Dim doc As HTMLDocument
        Dim bod As IHTMLElement
        If cf = CF_XRTF Or _
           cf = CF_WRTF Or _
           cf = ClipBoard1.GetCBFormatForName("Rich Text") Or _
           cf = ClipBoard1.GetCBFormatForName("RTF") Then
            RichTextBox1.TextRTF = s
            'RichTextBox1.TextRTF
        Else
            Set doc = WebBrowser1.document
            Set bod = doc.body
            If cf = ClipboardFormat.CF_HTML_xls1 Or _
               cf = ClipboardFormat.CF_HTML_xls2 Or _
               cf = ClipBoard1.GetCBFormatForName("HTML") Then
                's = ExtractHTML(s)
                bod.innerHTML = "html:" & vbCrLf & ExtractHTML(s)
            ElseIf cf = CF_XML Or _
                   cf = ClipBoard1.GetCBFormatForName("XML") Then
                bod.innerHTML = "xml:" & vbCrLf & ExtractHTML(s)
            Else
                'for unknown binary-data we need a Hex-viewer!!
                'OK wir brauchen eine Klasse ClipBoardContent die speichert alle ClipBoardElement Objekte
                'die erh�lt als Property die Clipboard-Konstante
                'und speichert den kompletten Inhalt eines Elements in einem Bytearray
                'anhand der konstante kann die Klasse selbst�ndig ad�quat umwandeln,
                'also im Grunde so �hnlich wie der Variant
                'entweder in Text, rtf, html, bitmap, Riff, wave, bin�r oder sonstwas
                Dim ba() As Byte: ba = s
                s = ""
                Dim i As Long
                For i = 0 To UBound(ba)
                    s = s & Hex2b(ba(i)) & " "
                Next
            End If
        End If
    End Select
    Text1.Text = s
End Sub

Function Hex2(s As String) As String
    Hex2 = IIf(Len(s) = 1, "0", "") & s
End Function

Function Hex2b(ByVal b As Byte) As String
    Hex2b = Hex(b)
    Hex2b = IIf(Len(Hex2b) = 1, "0", "") & Hex2b
End Function

Function Hex4(s As String) As String
    Hex4 = s: If Len(s) < 4 Then Hex4 = String(4 - Len(s), "0") & Hex4
End Function

Private Function ExtractHTML(ByVal scbHTML As String) As String
    Dim s As String
    s = VBA.Mid$(Trim(scbHTML), 2, 1)
    If s = "<" Then
        ExtractHTML = scbHTML
        Exit Function
    End If
    Dim pos As Long ', slen As Long
    Dim StartHTML As Long, EndHTML   As Long
    If Len(scbHTML) = 0 Then Exit Function
    pos = InStr(1, scbHTML, "StartHTML:", vbTextCompare)
    If pos > 0 Then
        s = Mid(scbHTML, pos + 10, 10)
        StartHTML = CLng(s)
        pos = InStr(1, scbHTML, "EndHTML:", vbTextCompare)
        If pos > 0 Then
            s = Mid(scbHTML, pos + 8, 10)
            EndHTML = CLng(s)
            ExtractHTML = Trim(Mid(scbHTML, StartHTML + 2, EndHTML - StartHTML - Len(s) - 4)) ' - 2))
        End If
    End If
End Function

'
'Excel 2010:
'Version:1.0
'StartHTML:0000000196
'EndHTML:0000002905
'StartFragment:0000002496
'EndFragment:0000002845
'SourceURL:file:///C:\Users\Oliver%20Meyer\OneDrive\Documents\ExcelTools\FormatFomula.xlsm
'
'<html xmlns:v="urn:schemas-microsoft-com:vml"
'xmlns: o = "urn:schemas-microsoft-com:office:office"
'xmlns: x = "urn:schemas-microsoft-com:office:excel"
'xmlns="http://www.w3.org/TR/REC-html40">
'
'<head>
'<meta http-equiv=Content-Type content="text/html; charset=utf-8">
'<meta name=ProgId content=Excel.Sheet>
'<meta name=Generator content="Microsoft Excel 14">
'<link id=Main-File rel=Main-File
'href="file:///C:\Users\OLIVER~1\AppData\Local\Temp\msohtmlclip1\01\clip.htm">
'<link rel=File-List
'href="file:///C:\Users\OLIVER~1\AppData\Local\Temp\msohtmlclip1\01\clip_filelist.xml">
'<style>
'<!--table
'    {mso-displayed-decimal-separator:"\,";
'    mso-displayed-thousand-separator:"\.";}
'@page
'    {margin:.79in .7in .79in .7in;
'    mso-header-margin:.3in;
'    mso-footer-margin:.3in;}
'.font0
'    {color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Calibri, sans-serif;
'    mso-font-charset:0;}
'.font5
'    {color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Symbol, serif;
'    mso-font-charset:2;}
'.font6
'    {color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Calibri, sans-serif;
'    mso-font-charset:0;}
'.font7
'    {color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Calibri, sans-serif;
'    mso-font-charset:0;}
'tr
'    {mso-height-source:auto;}
'col
'    {mso-width-source:auto;}
'br
'    {mso-data-placement:same-cell;}
'td
'    {padding-top:1px;
'    padding-right:1px;
'    padding-left:1px;
'    mso-ignore:padding;
'    color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Calibri, sans-serif;
'    mso-font-charset:0;
'    mso-number-format:General;
'    text-align:general;
'    vertical-align:bottom;
'    border:none;
'    mso-background-source:auto;
'    mso-pattern:auto;
'    mso-protection:locked visible;
'    white-space:nowrap;
'    mso-rotate:0;}
'-->
'</style>
'</head>
'
'<body link=blue vlink=purple>
'
'<table border=0 cellpadding=0 cellspacing=0 width=150 style='border-collapse:
' collapse;width:113pt'>
' <col width=150 style='mso-width-source:userset;mso-width-alt:5485;width:113pt'>
' <tr height=25 style='height:18.75pt'>
'<!--StartFragment-->
'  <td height=25 width=150 style='height:18.75pt;width:113pt'>r<font
'  class="font7"><sup>2</sup></font><font class="font0"> * </font><font
'  class="font5">p</font><font class="font0"> - </font><font class="font5">a</font><font
'  class="font0"> * a</font><font class="font6"><sub>i,k</sub></font><font
'  class="font7"><sup>2</sup></font></td>
'<!--EndFragment-->
' </tr>
'</table>
'
'</body>
'
'</html>
'
'
'
'
'
'
'
'
'Excel 2016:
'Version:1.0
'StartHTML:0000000105
'EndHTML:0000001790
'StartFragment:0000001659
'EndFragment:0000001730
'
'
'<html xmlns:o="urn:schemas-microsoft-com:office:office"
'xmlns: X = "urn:schemas-microsoft-com:office:excel"
'xmlns="http://www.w3.org/TR/REC-html40">
'
'
'<head>
'<meta http-equiv=Content-Type content="text/html; charset=utf-8">
'<meta name=ProgId content=Excel.Sheet>
'<meta name=Generator content="Microsoft Excel 15">
'<link id=Main-File rel=Main-File
'href="file:///C:/Users/olimi/AppData/Local/Temp/msohtmlclip1/01/clip.htm">
'<link rel=File-List
'href="file:///C:/Users/olimi/AppData/Local/Temp/msohtmlclip1/01/clip_filelist.xml">
'<style>
'<!--table
'    {mso-displayed-decimal-separator:"\,";
'    mso-displayed-thousand-separator:"\.";}
'@page
'    {margin:.79in .7in .79in .7in;
'    mso-header-margin:.3in;
'    mso-footer-margin:.3in;}
'tr
'    {mso-height-source:auto;}
'col
'    {mso-width-source:auto;}
'br
'    {mso-data-placement:same-cell;}
'td
'    {padding-top:1px;
'    padding-right:1px;
'    padding-left:1px;
'    mso-ignore:padding;
'    color:black;
'    font-size:11.0pt;
'    font-weight:400;
'    font-style:normal;
'    text-decoration:none;
'    font-family:Calibri, sans-serif;
'    mso-font-charset:0;
'    mso-number-format:General;
'    text-align:general;
'    vertical-align:bottom;
'    border:none;
'    mso-background-source:auto;
'    mso-pattern:auto;
'    mso-protection:locked visible;
'    white-space:nowrap;
'    mso-rotate:0;}
'-->
'</style>
'</head>
'
'<body link="#0563C1" vlink="#954F72">
'
'<table border=0 cellpadding=0 cellspacing=0 width=80 style='border-collapse:
' collapse;width:60pt'>
' <col width=80 style='width:60pt'>
' <tr height=20 style='height:15.0pt'>
'<!--StartFragment-->
'  <td height=20 width=80 style='height:15.0pt;width:60pt'>test</td>
'<!--EndFragment-->
' </tr>
'</table>
'
'</body>
'
'</html>


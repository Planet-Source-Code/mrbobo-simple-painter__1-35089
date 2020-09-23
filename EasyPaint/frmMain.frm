VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "   Simple Painter"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicDisplay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   240
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   240
      Width           =   4500
      Begin VB.PictureBox PicTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   2520
         MouseIcon       =   "frmMain.frx":030A
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.PictureBox PicSelect 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   240
         MouseIcon       =   "frmMain.frx":0614
         MousePointer    =   5  'Size
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
         Begin VB.Shape PicSelectShape 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   615
            Left            =   0
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Shape ShapeSelect 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   615
         Left            =   2160
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Visible         =   0   'False
         X1              =   56
         X2              =   152
         Y1              =   128
         Y2              =   128
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuEditBase 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select All"
         Index           =   5
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Select None"
         Enabled         =   0   'False
         Index           =   6
      End
   End
   Begin VB.Menu mnuToolsBase 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools 
         Caption         =   "Select"
         Index           =   0
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Pen"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Eraser"
         Index           =   2
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Flood Fill"
         Index           =   3
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Line"
         Index           =   4
      End
   End
   Begin VB.Menu mnuColorBase 
      Caption         =   "&Color"
      Begin VB.Menu mnuColor 
         Caption         =   "Red"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Green"
         Index           =   1
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Blue"
         Index           =   2
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Black"
         Index           =   3
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Cyan"
         Index           =   4
      End
      Begin VB.Menu mnuColor 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Custom"
         Index           =   6
      End
   End
   Begin VB.Menu mnuDrawWidthBase 
      Caption         =   "&Draw Width"
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "1 Pixel"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "2 Pixel"
         Index           =   2
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "3 Pixel"
         Index           =   3
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "4 Pixel"
         Index           =   4
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "5 Pixel"
         Index           =   5
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "6 Pixel"
         Index           =   6
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "7 Pixel"
         Index           =   7
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "8 Pixel"
         Index           =   8
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "9 Pixel"
         Index           =   9
      End
      Begin VB.Menu mnuDrawWidth 
         Caption         =   "10 Pixel"
         Index           =   10
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive


'This demo shows in the simplest possible code, the basics of
'image editing. It is intended for the novice but others may
'find it useful in optimising or simplifying their code.

'When first starting coding graphics apps, one of the functions
'that can cause problems is making selections and controlling
'input to those selections. I hope this demo makes those tasks
'clearer. Whilst it is in no way a full blown editor, it may provide
'a basis from which to learn.

'If you want an example of a more advanced image editor I
'suggest you look at "Bobo Image Workshop". You can download
'this from Planet Source Code
'address :http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=21482&lngWId=1



Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Dim sx As Long, sy As Long, CurColor As Long, CurTool As Long
Dim DontMouseUp As Boolean, DontDraw As Boolean
Private Sub Form_Load()
    CurColor = vbRed
    CurTool = 1 'pen
    InitCmnDlg Me.hwnd
End Sub
Private Sub Form_Resize()
    On Error Resume Next 'center displayed image
    PicDisplay.Left = (Me.ScaleWidth - PicDisplay.Width) / 2
    PicDisplay.Top = (Me.ScaleHeight - PicDisplay.Height) / 2
End Sub
Private Sub mnuColor_Click(Index As Integer)
    Dim z As Long, col As Long
    For z = 0 To mnuColor.Count - 1 'uncheck all menu items
        If z <> 5 Then mnuColor(z).Checked = False
    Next
    mnuColor(Index).Checked = True 'check clicked menu
    Select Case Index 'change current color accordingly
        Case 0: CurColor = vbRed
        Case 1: CurColor = vbGreen
        Case 2: CurColor = vbBlue
        Case 3: CurColor = vbBlack
        Case 4: CurColor = vbCyan
        Case 6 'select a color using Windows color picker
            col = ShowColor
            If col <> -1 Then CurColor = col
    End Select
End Sub
Private Sub mnuDrawWidth_Click(Index As Integer)
    Dim z As Long
    For z = 1 To mnuDrawWidth.Count 'uncheck all menu items
        mnuDrawWidth(z).Checked = False
    Next
    mnuDrawWidth(Index).Checked = True 'check clicked menu
    PicDisplay.DrawWidth = Index 'set draw width
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
        Case 0 'cut
            Clipboard.Clear
            Clipboard.SetData PicSelect.Picture
            PicSelect.Visible = False
        Case 1 'copy
            Clipboard.Clear
            Clipboard.SetData PicSelect.Picture
        Case 2 'paste
            PicSelect.Picture = LoadPicture()
            PicSelect.AutoSize = True
            PicSelect.Picture = Clipboard.GetData
            PicSelect.AutoSize = False
            PicSelect.Left = 0
            PicSelect.Top = 0
            PicSelect.Visible = True
        Case 3 'delete
            PicSelect.Visible = False
        Case 5 'select all
            PicSelect.Left = 0
            PicSelect.Top = 0
            PicSelect.Width = PicDisplay.Width
            PicSelect.Height = PicDisplay.Height
            PicSelect.Picture = PicDisplay.Image
            PicDisplay.Picture = LoadPicture()
            PicSelect.Visible = True
        Case 6 'select none
            BitBlt PicDisplay.hDC, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicSelect.hDC, 0, 0, vbSrcCopy
            PicDisplay.Picture = PicDisplay.Image
            PicSelect.Visible = False
    End Select
End Sub

Private Sub mnuEditBase_Click()
    mnuEdit(0).Enabled = PicSelect.Visible 'cut enabled
    mnuEdit(1).Enabled = PicSelect.Visible 'copy enabled
    mnuEdit(2).Enabled = Clipboard.GetFormat(vbCFBitmap) 'paste enabled
    mnuEdit(3).Enabled = PicSelect.Visible 'delete enabled
    mnuEdit(6).Enabled = PicSelect.Visible 'select none enabled
End Sub

Private Sub mnuFileNew_Click()
    frmNew.Show vbModal, Me 'see frmNew code
    Form_Resize
End Sub

Private Sub mnuFileOpen_Click()
    DontDraw = True 'stops accidental drawing when the dialog closes
    With cmndlg
        .Filter = "Image files |*.bmp;*.jpg;*.gif;*.ico;*.cur"
        ShowOpen
        If Len(.filename) = 0 Then Exit Sub
        PicDisplay.Picture = LoadPicture(.filename) 'open an image file
    End With
End Sub
Private Sub mnuFileSaveAs_Click()
    With cmndlg
        .Filter = "Bitmap files (*.bmp)|*.bmp"
        ShowSave
        If Len(.filename) = 0 Then Exit Sub
        SavePicture PicDisplay.Image, .filename 'save an image file as bitmap
    End With
End Sub
Private Sub mnuTools_Click(Index As Integer)
    Dim z As Long
    For z = 0 To mnuTools.Count - 1 'uncheck all menu items
        mnuTools(z).Checked = False
    Next
    mnuTools(Index).Checked = True 'check clicked menu
    CurTool = Index 'current tool
    Select Case Index 'provide a cursor to indicate tool
        Case 0, 4 'selection,line
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicTemp.MouseIcon
        Case 3 'flood fill
            PicDisplay.MousePointer = 99
            PicDisplay.MouseIcon = PicSelect.MouseIcon
        Case Else 'just the default arrow
            PicDisplay.MousePointer = 0
    End Select
    
    
End Sub
Private Sub PicDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    sx = x 'remember the start position of editing
    sy = y
    If Button = 1 Then
        DontDraw = False
        If PicSelect.Visible Then 'a selection is in progress so paste in situ
            BitBlt PicDisplay.hDC, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicSelect.hDC, 0, 0, vbSrcCopy
            PicDisplay.Picture = PicDisplay.Image
            PicSelect.Visible = False
            DontMouseUp = True 'flag to tell MouseUp sub not to respond
            Exit Sub
        End If
        Select Case CurTool
            Case 0
                ShapeSelect.Left = x 'display selection area
                ShapeSelect.Top = y
                ShapeSelect.Width = 0
                ShapeSelect.Height = 0
                ShapeSelect.Visible = True
            Case 4
                Line1.X1 = sx 'display line co-ordinates
                Line1.X2 = sx
                Line1.Y1 = sy
                Line1.Y2 = sy
                Line1.Visible = True
        End Select
    End If
End Sub
Private Sub PicDisplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And DontDraw = False Then
        Select Case CurTool
            Case 0 'selection
                ShapeSelect.Left = IIf(x > sx, sx, x)
                ShapeSelect.Top = IIf(y > sy, sy, y)
                ShapeSelect.Width = IIf(x > sx, x - sx, sx - x)
                ShapeSelect.Height = IIf(y > sy, y - sy, sy - y)
            Case 1, 2 'pen, eraser
                PicDisplay.Line (sx, sy)-(x, y), IIf(CurTool = 1, CurColor, vbWhite)
                sx = x
                sy = y
            Case 4 'line
                Line1.X2 = x
                Line1.Y2 = y
        End Select
    End If
End Sub
Private Sub PicDisplay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    sx = -1
    If Button = 1 Then
        Select Case CurTool
            Case 0
                If DontMouseUp Then
                    DontMouseUp = False 'reset flag
                    Exit Sub 'user clicked to remove selection so bail out here
                End If
                ShapeSelect.Visible = False 'hide selection shape
                PicSelect.Left = ShapeSelect.Left 'position PicSelect to match selection
                PicSelect.Top = ShapeSelect.Top
                PicSelect.Width = ShapeSelect.Width
                PicSelect.Height = ShapeSelect.Height
                PicTemp.Width = ShapeSelect.Width 'used to blank out behind PicSelect
                PicTemp.Height = ShapeSelect.Height
                PicDisplay.Picture = PicDisplay.Image 'set the picture property before using Bitblt
                'fill PicSelect with selection
                BitBlt PicSelect.hDC, 0, 0, PicSelect.Width, PicSelect.Height, PicDisplay.hDC, ShapeSelect.Left, ShapeSelect.Top, vbSrcCopy
                'clear image behind PicSelect
                BitBlt PicDisplay.hDC, PicSelect.Left, PicSelect.Top, PicSelect.Width, PicSelect.Height, PicTemp.hDC, 0, 0, vbSrcCopy
                PicDisplay.Picture = PicDisplay.Image 'set the picture property
                PicSelect.Picture = PicSelect.Image 'set the picture property
                PicSelect.Visible = True 'show PicSelect
            Case 3 'Flood Fill
                PicDisplay.FillStyle = 0
                PicDisplay.FillColor = CurColor
                ExtFloodFill PicDisplay.hDC, x, y, GetPixel(PicDisplay.hDC, x, y), 1
                PicDisplay.FillStyle = 1
                PicDisplay.Refresh
            Case 4 'Line
                PicDisplay.Line (Line1.X1, Line1.Y1)-(x, y), CurColor
                Line1.Visible = False
        End Select
    End If
End Sub
Private Sub PicDisplay_Resize()
    On Error Resume Next
    Dim W As Long, H As Long
    'center image
    W = PicDisplay.Width * Screen.TwipsPerPixelX + 480
    H = PicDisplay.Height * Screen.TwipsPerPixelY + 1020
    'adjust form size to fit
    If Me.Width < W Then Me.Width = W
    If Me.Height < H Then Me.Height = H
    Form_Resize
End Sub

Private Sub PicSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'enable dragging of selection
    ReleaseCapture
    SendMessage PicSelect.hwnd, &HA1, 2, 0&
End Sub

Private Sub PicSelect_Resize()
    On Error Resume Next
    'show a dotted edge to selection
    PicSelectShape.Width = PicSelect.Width
    PicSelectShape.Height = PicSelect.Height
End Sub

VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Image"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size in Pixels"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.VScrollBar VSHeight 
         Height          =   255
         Left            =   2280
         Max             =   1000
         TabIndex        =   6
         Top             =   975
         Value           =   800
         Width           =   255
      End
      Begin VB.VScrollBar VSWidth 
         Height          =   255
         Left            =   2280
         Max             =   1000
         TabIndex        =   5
         Top             =   495
         Value           =   700
         Width           =   255
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "200"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Text            =   "300"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Height"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Width"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   540
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Clear and size PicDisplay according to user selection
    frmMain.PicDisplay.Picture = LoadPicture()
    frmMain.PicDisplay.Width = Val(txtWidth.Text)
    frmMain.PicDisplay.Height = Val(txtHeight.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    'Reduce EXE size by using same icon
    Me.Icon = frmMain.Icon
End Sub
Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then 'Numbers only
        If KeyAscii <> 8 Then KeyAscii = 0
    End If

End Sub

Private Sub txtHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(txtHeight.Text) > 1000 Then txtHeight.Text = "1000"
    If Val(txtHeight.Text) < 4 Then txtHeight.Text = "4"
    VSHeight.Value = 1000 - Val(txtHeight.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then 'Numbers only
        If KeyAscii <> 8 Then KeyAscii = 0
    End If

End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(txtWidth.Text) > 1000 Then txtWidth.Text = "1000"
    If Val(txtWidth.Text) < 4 Then txtWidth.Text = "4"
    VSWidth.Value = 1000 - Val(txtWidth.Text)
End Sub

Private Sub VSHeight_Change()
    If VSHeight.Value > 996 Then VSHeight.Value = 996 'equivalent to 4
    txtHeight.Text = 1000 - VSHeight.Value
End Sub

Private Sub VSWidth_Change()
    If VSWidth.Value > 996 Then VSWidth.Value = 996 'equivalent to 4
    txtWidth.Text = 1000 - VSWidth.Value
End Sub

Attribute VB_Name = "Module1"
'******************************************************************
'***************Copyright PSST 2001********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Standard Commondialog - nothing special here
'this code is available EVERYWHERE
Option Explicit
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type CMDialog
    ownerform As Long
    Filter As String
    filetitle As String
    FilterIndex As Long
    filename As String
    initdir As String
    dialogtitle As String
    flags As Long
End Type
Public cmndlg As CMDialog
Dim CustomColors() As Byte
Public Sub ShowOpen()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .initdir
        OFName.lpstrTitle = .dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.flags = .flags
        If GetOpenFileName(OFName) Then
            .filename = StripTerminator(Trim$(OFName.lpstrFile))
            .filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        End If
    End With
End Sub
Public Sub ShowSave()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .initdir
        OFName.lpstrTitle = .dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.flags = .flags
        If GetSaveFileName(OFName) Then
            .filename = StripTerminator(Trim$(OFName.lpstrFile))
            .filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .FilterIndex = OFName.nFilterIndex
        End If
    End With
End Sub

Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Sub InitCmnDlg(mOwner As Long)
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim i As Integer
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 254
    Next
    cmndlg.ownerform = mOwner
End Sub
Public Function ShowColor() As Long
    Dim cc As CHOOSECOLOR, mcc As Long
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = cmndlg.ownerform
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If CHOOSECOLOR(cc) <> 0 Then
        mcc = cc.rgbResult
        If mcc < 0 Then mcc = -mcc
        If mcc > vbWhite Then mcc = vbWhite
        ShowColor = mcc
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function


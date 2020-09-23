VERSION 5.00
Begin VB.UserControl StarTrekProgress 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekProgress.ctx":0000
End
Attribute VB_Name = "StarTrekProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'api
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'
Private BC As OLE_COLOR, TC As OLE_COLOR
'back color of left, middle and right or top, middle and bottom
Private bc2Back As OLE_COLOR
Private BC2 As OLE_COLOR, BC1 As OLE_COLOR, BC2_2 As OLE_COLOR
'
Private spcLFT As Integer, spcRGHT As Integer, lnSPC As Integer

Private mStyle As eSTProgBar
Private minV As Integer, maxV As Integer, currV As Integer

'
'events
Public Event Change()
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------

'

Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    reDraw
    PropertyChanged "BackColor"
End Property
'
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = TC
End Property
Public Property Let TransparentColor(ByVal nV As OLE_COLOR)
    TC = nV
    reDraw
    PropertyChanged "TransparentColor"
End Property
'
Public Property Get SideBackColor() As OLE_COLOR
    SideBackColor = BC1
End Property
Public Property Let SideBackColor(ByVal nV As OLE_COLOR)
    BC1 = nV
    reDraw
    PropertyChanged "SideBackColor"
End Property
'
Public Property Get BarColor1() As OLE_COLOR
    BarColor1 = BC2
End Property
Public Property Let BarColor1(ByVal nV As OLE_COLOR)
    BC2 = nV
    reDraw
    PropertyChanged "BarColor1"
End Property
'
Public Property Get BarColor2() As OLE_COLOR
    BarColor2 = BC2_2
End Property
Public Property Let BarColor2(ByVal nV As OLE_COLOR)
    BC2_2 = nV
    reDraw
    PropertyChanged "BarColor2"
End Property
'
Public Property Get BarBackColor() As OLE_COLOR
    BarBackColor = bc2Back
End Property
Public Property Let BarBackColor(ByVal nV As OLE_COLOR)
    bc2Back = nV
    reDraw
    PropertyChanged "BarBackColor"
End Property

'--------
Public Property Get Style() As eSTProgBar
    Style = mStyle
End Property
Public Property Let Style(ByVal nV As eSTProgBar)
    mStyle = nV
    reDraw
    PropertyChanged "Style"
End Property
'--------
Public Property Get SideSpacingLeft() As Integer
    SideSpacingLeft = spcLFT
End Property
Public Property Let SideSpacingLeft(ByVal nV As Integer)
    spcLFT = nV
    reDraw
    PropertyChanged "SideSpacingLeft"
End Property
'
Public Property Get SideSpacingRight() As Integer
    SideSpacingRight = spcRGHT
End Property
Public Property Let SideSpacingRight(ByVal nV As Integer)
    spcRGHT = nV
    reDraw
    PropertyChanged "SideSpacingRight"
End Property
'
'
Public Property Get LineSpacing() As Integer
    LineSpacing = lnSPC
End Property
Public Property Let LineSpacing(ByVal nV As Integer)
    lnSPC = nV
    If lnSPC < 1 Then lnSPC = 1
    reDraw
    PropertyChanged "LineSpacing"
End Property
'
Public Property Get Min() As Integer
    Min = minV
End Property
Public Property Let Min(ByVal nV As Integer)
    minV = nV
    reDraw
    PropertyChanged "Min"
End Property
'
Public Property Get Max() As Integer
    Max = maxV
End Property
Public Property Let Max(ByVal nV As Integer)
    maxV = nV
    reDraw
    PropertyChanged "Max"
End Property
'
Public Property Get Value() As Integer
    Value = currV
End Property
Public Property Let Value(ByVal nV As Integer)
    If nV >= minV And nV <= maxV Then
        currV = nV
        RaiseEvent Change
        reDraw
    End If
    PropertyChanged "Value"
End Property


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'set defaults
Private Sub UserControl_Initialize()
    TC = vbBlack
    BC = vbBlack
    
    lnSPC = 10
    '
    BC2 = &HEF8E4F
    BC2_2 = &HEF8E4F
    BC1 = &HEF8E4F
    '
    bc2Back = &H80C0FF
    '
    mStyle = stFedRound
    '
    minV = 0
    maxV = 100
    currV = 0
    '
    spcRGHT = 0
    spcLFT = 5
    
    '
    reDraw
End Sub
'redraw control
Private Sub reDraw()
    UserControl.Cls
    
    If mStyle = stFedRound Or mStyle = stFedGrid Then
        drawFederationProgBar
    ElseIf mStyle = stRomulan Then
        drawRomulanProgBar
    End If
    UserControl.Refresh
    'create mask
    UserControl.MaskColor = TC
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
End Sub
'draw federation progress bar
Private Sub drawFederationProgBar()
    'On Error Resume Next
    Dim ucSW As Single, ucSH As Integer
    Dim barSize As Integer
    Dim cGrad As New clsGradient
    
    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    
    UserControl.FillStyle = 0
    UserControl.Line (0, 0)-(ucSW, ucSH), bc2Back, BF
    If mStyle = stFedRound Then
        UserControl.Line (0, 0)-(ucSH / 2 + spcLFT + 1, ucSH), BC, BF
        UserControl.Line (ucSW, 0)-(ucSW - ucSH / 2 - spcRGHT, ucSH), BC, BF
        barSize = (ucSW - (ucSH + spcLFT + spcRGHT + 1)) * (currV - minV) / (maxV - minV)
        'left side
        UserControl.Circle (ucSH / 2, ucSH / 2), ucSH / 2, BC1, 3.14 / 2, 1.5 * 3.14
        UserControl.Line (ucSH / 2, -10)-(ucSH / 2, ucSH + 10), BC1
        UserControl.FillColor = BC1
        ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
        'right side
        UserControl.Circle (ucSW - ucSH / 2, ucSH / 2), ucSH / 2, BC1, 1.5 * 3.14, 3.14 / 2
        UserControl.Line (ucSW - ucSH / 2 - 0.5, 0)-(ucSW - ucSH / 2 - 0.5, ucSH), BC1
        UserControl.FillColor = BC1
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 2, Point(ucSW - 2, ucSH / 2), 1
        'middle
        UserControl.Line (ucSH / 2 + spcLFT + 1, 0)-(ucSH / 2 + spcLFT + barSize + 1, ucSH), BC2, BF
    ElseIf mStyle = stFedGrid Then
        Dim lnNum As Integer, i As Integer, mX As Integer
        '
        'lnSPC = 10
        If lnSPC < 1 Then lnSPC = 10
        UserControl.Line (0, 0)-(ucSW, ucSH), BC, BF
        'If currV > 60 Then MsgBox ucSW * (currV - minV) / (maxV - minV)
        barSize = ucSW * (currV - minV) / (maxV - minV)
        'MsgBox ucSW * (currV - minV) / (maxV - minV)
        lnNum = ucSW / lnSPC
        cGrad.Angle = 0
        cGrad.Color1 = BC2
        cGrad.Color2 = BC2_2
        cGrad.Draw UserControl.hWnd, UserControl.hDC, 0, 0
        
        UserControl.Line (barSize, 0)-(ucSW, ucSH), BC, BF
        mX = 0
        UserControl.FillStyle = 1
        UserControl.Line (0, 0)-(ucSW - 1, ucSH - 1), bc2Back, B
        For i = 1 To lnNum + 1
            UserControl.Line (mX, 0)-(mX, ucSH), bc2Back
            mX = mX + lnSPC
        Next i
        UserControl.FillStyle = 0
    End If
End Sub

Private Sub drawRomulanProgBar()
    Dim ucSW As Single, ucSH As Integer
    Dim barSize As Integer, mLnWid As Integer
    Dim cGrad As New clsGradient
    '
    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    Dim lnNum As Integer, i As Integer, mX As Integer
    If lnSPC < 1 Then lnSPC = 10
    
    mLnWid = 5
    
    
    barSize = ucSW * (currV - minV) / (maxV - minV)
    'If Err.Number <> 0 Then MsgBox ucSW * (currV - minV) / (maxV - minV)
    lnNum = ucSW / (lnSPC + mLnWid)
    UserControl.FillStyle = 1
    UserControl.BackStyle = 1
    cGrad.Angle = 0
    cGrad.Color1 = BC2
    cGrad.Color2 = BC2_2
    cGrad.Draw UserControl.hWnd, UserControl.hDC, 0, 0
    'UserControl.Line (0, 0)-(barSize, ucSH), BC2, BF
    UserControl.Line (barSize, 0)-(ucSW, ucSH), bc2Back, BF
    'UserControl.Line (0, 0)-(ucSW, ucSH), bc2Back, BF
    mX = 0
    For i = 1 To lnNum + 1
        UserControl.Line (mX, 0)-(mX + lnSPC, ucSH), BC, BF
        mX = mX + (lnSPC + mLnWid)
    Next i
    
    UserControl.Line (0, 1)-(ucSW, 2), BC, BF
    UserControl.Line (0, ucSH - 2)-(ucSW, ucSH - 3), BC, BF
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BC = PropBag.ReadProperty("BackColor", vbBlack)
    TC = PropBag.ReadProperty("TransparenColor", vbBlack)
    '
    BC2 = PropBag.ReadProperty("BarColor1", &HEF8E4F)
    BC2_2 = PropBag.ReadProperty("BarColor2", &HEF8E4F)
    bc2Back = PropBag.ReadProperty("BarBackColor", &H80C0FF)
    '
    mStyle = PropBag.ReadProperty("Style", 0)
    '
    spcRGHT = PropBag.ReadProperty("SideSpacingRight", 0)
    spcLFT = PropBag.ReadProperty("SideSpacingLeft", 5)
    lnSPC = PropBag.ReadProperty("LineSpacing", 10)
    BC1 = PropBag.ReadProperty("SideBackColor", &HEF8E4F)
    minV = PropBag.ReadProperty("Min", 0)
    maxV = PropBag.ReadProperty("Max", 100)
    currV = PropBag.ReadProperty("Value", 0)
    reDraw
End Sub

Private Sub UserControl_Resize()
    reDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BC, vbBlack
    PropBag.WriteProperty "TransparenColor", TC, vbBlack
    PropBag.WriteProperty "BarColor1", BC2, &HEF8E4F
    PropBag.WriteProperty "BarColor2", BC2_2, &HEF8E4F
    PropBag.WriteProperty "BarBackColor", bc2Back, &H80C0FF
    PropBag.WriteProperty "SideBackColor", BC1, &HEF8E4F
    PropBag.WriteProperty "Style", mStyle, 0
    '
    PropBag.WriteProperty "SideSpacingLeft", spcLFT, 5
    PropBag.WriteProperty "SideSpacingRight", spcRGHT, 0
    PropBag.WriteProperty "LineSpacing", lnSPC, 10
    PropBag.WriteProperty "Min", minV, 0
    PropBag.WriteProperty "Max", maxV, 100
    PropBag.WriteProperty "Value", currV, 0
End Sub


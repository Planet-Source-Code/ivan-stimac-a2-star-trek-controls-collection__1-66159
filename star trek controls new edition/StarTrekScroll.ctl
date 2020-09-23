VERSION 5.00
Begin VB.UserControl StarTrekScroll 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekScroll.ctx":0000
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2100
      Top             =   780
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   480
   End
End
Attribute VB_Name = "StarTrekScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'api
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'
Private BC As OLE_COLOR, TC As OLE_COLOR
'back color of left, middle and right or top, middle and bottom
Private bc2Back As OLE_COLOR
Private BC1 As OLE_COLOR, BC2 As OLE_COLOR
Private BC1Hover As OLE_COLOR, BC2Hover As OLE_COLOR
Private BC1Down As OLE_COLOR, BC2Down As OLE_COLOR
Private BC1Disabled As OLE_COLOR, BC2Disabled As OLE_COLOR
'
Private buttSpacing As Integer

Private mStyle As eSTScrStyles
Private minV As Integer, maxV As Integer, currV As Integer
Private chStep As Integer

Private enbl As Boolean

Private bState(3) As Byte
Private bDown As Byte, bHover As Byte
Private isDrag As Boolean, dragPos As Integer
'
'events
Public Event Change()
Public Event Scroll()
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property
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
Public Property Get ButtBackColor() As OLE_COLOR
    ButtBackColor = BC1
End Property
Public Property Let ButtBackColor(ByVal nV As OLE_COLOR)
    BC1 = nV
    reDraw
    PropertyChanged "ButtBackColor"
End Property
'
Public Property Get ButtBackColorHover() As OLE_COLOR
    ButtBackColorHover = BC1Hover
End Property
Public Property Let ButtBackColorHover(ByVal nV As OLE_COLOR)
    BC1Hover = nV
    reDraw
    PropertyChanged "ButtBackColorHover"
End Property
'
Public Property Get ButtBackColorDown() As OLE_COLOR
    ButtBackColorDown = BC1Down
End Property
Public Property Let ButtBackColorDown(ByVal nV As OLE_COLOR)
    BC1Down = nV
    reDraw
    PropertyChanged "ButtBackColorDown"
End Property
'
Public Property Get ButtBackColorDisabled() As OLE_COLOR
    ButtBackColorDisabled = BC1Disabled
End Property
Public Property Let ButtBackColorDisabled(ByVal nV As OLE_COLOR)
    BC1Disabled = nV
    reDraw
    PropertyChanged "ButtBackColorDisabled"
End Property
'
'
Public Property Get ScrollBackColor() As OLE_COLOR
    ScrollBackColor = BC2
End Property
Public Property Let ScrollBackColor(ByVal nV As OLE_COLOR)
    BC2 = nV
    reDraw
    PropertyChanged "ScrollBackColor"
End Property
'
Public Property Get ScrollBackColorHover() As OLE_COLOR
    ScrollBackColorHover = BC2Hover
End Property
Public Property Let ScrollBackColorHover(ByVal nV As OLE_COLOR)
    BC2Hover = nV
    reDraw
    PropertyChanged "ScrollBackColorHover"
End Property
'
Public Property Get ScrollBackColorDown() As OLE_COLOR
    ScrollBackColorDown = BC2Down
End Property
Public Property Let ScrollBackColorDown(ByVal nV As OLE_COLOR)
    BC2Down = nV
    reDraw
    PropertyChanged "ScrollBackColorDown"
End Property
'
Public Property Get ScrollBackColorDownDisabled() As OLE_COLOR
    ScrollBackColorDownDisabled = BC2Disabled
End Property
Public Property Let ScrollBackColorDownDisabled(ByVal nV As OLE_COLOR)
    BC2Disabled = nV
    reDraw
    PropertyChanged "ScrollBackColorDownDisabled"
End Property
'
Public Property Get ScrollContainerBackColor() As OLE_COLOR
    ScrollContainerBackColor = bc2Back
End Property
Public Property Let ScrollContainerBackColor(ByVal nV As OLE_COLOR)
    bc2Back = nV
    reDraw
    PropertyChanged "ScrollContainerBackColor"
End Property
'--------
Public Property Get Style() As eSTScrStyles
    Style = mStyle
End Property
Public Property Let Style(ByVal nV As eSTScrStyles)
    mStyle = nV
    reDraw
    PropertyChanged "Style"
End Property
'--------
Public Property Get ButtonSpacing() As Integer
    ButtonSpacing = buttSpacing
End Property
Public Property Let ButtonSpacing(ByVal nV As Integer)
    buttSpacing = nV
    reDraw
    PropertyChanged "ButtonSpacing"
End Property
'
'Public Property Get Min() As Integer
'    Min = minV
'End Property
'Public Property Let Min(ByVal nV As Integer)
'    minV = nV
'    reDraw
'    PropertyChanged "Min"
'End Property
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
    currV = nV
    reDraw
    PropertyChanged "Value"
End Property
'
Public Property Get ChangeStep() As Integer
    ChangeStep = chStep
End Property
Public Property Let ChangeStep(ByVal nV As Integer)
    chStep = nV
    reDraw
    PropertyChanged "ChangeStep"
End Property
'
'
'-----------------------------------------------------------------
'if user click and hold one of the buttons then
'   increase/decrease value or drag scroll then
'   redraw control
Private Sub Timer1_Timer()
    If bDown = 1 Then
        currV = currV - chStep
        If currV < minV Then currV = minV
    ElseIf bDown = 3 Then
        currV = currV + chStep
        If currV > maxV Then currV = maxV
    End If
    reDraw
    'increase change speed by decreasing timer interval
    If Timer1.Interval > 100 Then Timer1.Interval = Timer1.Interval - 50
End Sub
'check is mouse ove control
Private Sub Timer2_Timer()
    Dim lpPos As POINTAPI
    GetCursorPos lpPos
    If WindowFromPoint(lpPos.X, lpPos.Y) <> UserControl.hWnd Then
        setState 0, 0
        reDraw
        Timer2.Enabled = False
    End If
End Sub
'
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'set defaults values
Private Sub UserControl_Initialize()
    TC = vbBlack
    BC = vbBlack
    
    BC1 = &HEF8E4F
    BC1Hover = &HFACAAE
    BC1Down = &HAB6016
    BC1Disabled = &H808080
    '
    BC2 = &HEF8E4F
    BC2Hover = &HFACAAE
    BC2Down = &HAB6016
    BC2Disabled = &H808080

    '
    enbl = True
    '

    '
    bc2Back = &H80C0FF
    
    mStyle = stScrFedHor
    '
    minV = 0
    maxV = 100
    currV = 0
    chStep = 50
    '
    buttSpacing = 5
    '
    bState(1) = 0
    bState(2) = 0
    bState(3) = 0
    '
    isDrag = False
    '
    reDraw
End Sub
'redraw control
Private Sub reDraw()
    UserControl.Enabled = enbl
    
    If mStyle = stScrFedHor Or mStyle = stScrFedVert Then
        drawFederationScroll
    End If
    'create mask
    UserControl.MaskColor = TC
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
End Sub
'draw federation scroll
Private Sub drawFederationScroll()
    Dim ucSH As Integer, ucSW As Integer
    Dim scrSize As Integer, scrPos As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR, mBC3 As OLE_COLOR
    UserControl.Cls
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    'set colors
    If enbl = True Then
        Select Case bState(1)
            Case 0: mBC1 = BC1
            Case 1: mBC1 = BC1Hover
            Case 2: mBC1 = BC1Down
        End Select
        Select Case bState(2)
            Case 0: mBC2 = BC2
            Case 1: mBC2 = BC2Hover
            Case 2: mBC2 = BC2Down
        End Select
        Select Case bState(3)
            Case 0: mBC3 = BC1
            Case 1: mBC3 = BC1Hover
            Case 2: mBC3 = BC1Down
        End Select
    Else
        mBC1 = BC1Disabled: mBC2 = BC2Disabled: mBC3 = BC2Disabled
    End If
    
    UserControl.Line (0, 0)-(ucSW, ucSH), bc2Back, BF
    'draw horizontal scroll
    If mStyle = stScrFedHor Then
        UserControl.Line (0, 0)-(ucSH / 2 + buttSpacing, ucSH), BC, BF
        UserControl.Line (ucSW, 0)-(ucSW - ucSH / 2 - buttSpacing, ucSH), BC, BF
        'find scroll size from maxV, minV and chStep
        scrSize = ((ucSW - ucSH - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        '(ucSW - ucSH - buttSpacing * 2) is size of scroll area
        If scrSize > (ucSW - ucSH - buttSpacing * 2) Then scrSize = (ucSW - ucSH - buttSpacing * 2)
        'find scroll position from value
        scrPos = ucSH / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSW - ucSH - buttSpacing * 2 - scrSize)
    
        UserControl.Line (scrPos, 0)-(scrPos + scrSize, ucSH), mBC2, BF
        '
        UserControl.Circle (ucSH / 2, ucSH / 2), ucSH / 2, mBC1, 3.14 / 2, 1.5 * 3.14
        UserControl.Line (ucSH / 2, 0)-(ucSH / 2, ucSH), mBC1
        '
        UserControl.FillColor = mBC1
        ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
         '
        UserControl.Circle (ucSW - ucSH / 2 - 1, ucSH / 2), ucSH / 2, mBC3, 1.5 * 3.14, 3.14 / 2
        UserControl.Line (ucSW - ucSH / 2 - 1, 0)-(ucSW - ucSH / 2 - 1, ucSH), mBC3
        '
        UserControl.FillColor = mBC3
        ExtFloodFill UserControl.hDC, ucSW - ucSH / 2 + 1, ucSH / 2, UserControl.Point(ucSW - ucSH / 2 + 1, ucSH / 2), 1
    'draw vertical scroll
    ElseIf mStyle = stScrFedVert Then
        UserControl.Line (0, 0)-(ucSW, ucSW / 2 + buttSpacing), BC, BF
        UserControl.Line (0, ucSH - ucSW / 2 - buttSpacing)-(ucSW, ucSH), BC, BF
        scrSize = ((ucSH - ucSW - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        If scrSize > (ucSH - ucSW - buttSpacing * 2) Then scrSize = (ucSH - ucSW - buttSpacing * 2)
        scrPos = ucSW / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSH - ucSW - buttSpacing * 2 - scrSize)
    
        UserControl.Line (0, scrPos)-(ucSW, scrPos + scrSize), mBC2, BF
        '
        UserControl.Circle (ucSW / 2, ucSW / 2), ucSW / 2, mBC1, 0, 3.14
        UserControl.Line (0 / 2, ucSW / 2)-(ucSW, ucSW / 2), mBC1
        '
        UserControl.FillColor = mBC1
        ExtFloodFill UserControl.hDC, ucSW / 2, 2, UserControl.Point(ucSW / 2, 2), 1
         '
        UserControl.Circle (ucSW / 2 - 1, ucSH - ucSW / 2), ucSW / 2, mBC3, 3.14, 0
        UserControl.Line (0, ucSH - ucSW / 2 - 1)-(ucSH, ucSH - ucSW / 2 - 1), mBC3
        '
        UserControl.FillColor = mBC3
        ExtFloodFill UserControl.hDC, ucSW / 2, ucSH - 2, UserControl.Point(ucSW / 2, ucSH - 2), 1

    End If
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim scrSize As Integer, scrPos As Integer
    Dim ucSW As Integer, ucSH As Integer
    
    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    
'    scrSize = ((ucSW - ucSH - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
'    If scrSize > (ucSW - ucSH - buttSpacing * 2) Then scrSize = (ucSW - ucSH - buttSpacing * 2)
'    scrPos = ucSH / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSW - ucSH - buttSpacing * 2 - scrSize)
    '
    If mStyle = stScrFedHor Then
        'find scroll size and position
        scrSize = ((ucSW - ucSH - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        If scrSize > (ucSW - ucSH - buttSpacing * 2) Then scrSize = (ucSW - ucSH - buttSpacing * 2)
        scrPos = ucSH / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSW - ucSH - buttSpacing * 2 - scrSize)
        'check is user click on firs button (decrease)
        If X >= 0 And X <= UserControl.ScaleHeight / 2 Then
            'if true then decrease value
            currV = currV - chStep
            If currV < minV Then currV = minV
            setState 1, 2
            'and redraw control
            If bDown <> 1 Then
                bDown = 1
                reDraw
                RaiseEvent Change
            End If
        'check is user click on increase button
        ElseIf X <= UserControl.ScaleWidth And X >= UserControl.ScaleWidth - UserControl.ScaleHeight / 2 Then
            'if true then increase value
            currV = currV + chStep
            If currV > maxV Then currV = maxV
            setState 3, 2
            'and redraw control
            If bDown <> 3 Then
                bDown = 3
                reDraw
                RaiseEvent Change
            End If
        'check is user click on scroll
        ElseIf X >= scrPos And X <= scrPos + scrSize Then
            'if true then start enable drag:isDrag = True
            dragPos = X - scrPos
            setState 2, 2
            If bDown <> 2 Then
                bDown = 2
                reDraw
            End If
            isDrag = True
        End If
    'all same but there use y position because this is for vertical scroll
    ElseIf mStyle = stScrFedVert Then
        scrSize = ((ucSH - ucSW - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        If scrSize > (ucSH - ucSW - buttSpacing * 2) Then scrSize = (ucSH - ucSW - buttSpacing * 2)
        scrPos = ucSW / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSH - ucSW - buttSpacing * 2 - scrSize)
        If Y >= 0 And Y <= ucSW / 2 Then
            currV = currV - chStep
            If currV < minV Then currV = minV
            setState 1, 2
            If bDown <> 1 Then
                bDown = 1
                reDraw
                RaiseEvent Change
            End If
            
        ElseIf Y <= ucSH And Y >= ucSH - ucSW / 2 Then
            currV = currV + chStep
            If currV > maxV Then currV = maxV
            setState 3, 2
            If bDown <> 3 Then
                bDown = 3
                reDraw
                RaiseEvent Change
            End If
            
        ElseIf Y >= scrPos And Y <= scrPos + scrSize Then
            dragPos = Y - scrPos
            setState 2, 2
            If bDown <> 2 Then
                bDown = 2
                reDraw
            End If
            isDrag = True
        End If
    End If
    'set timer
    If bDown <> 2 Then
        Timer1.Interval = 300
        Timer1.Enabled = True
    Else
        Timer1.Interval = 50
        Timer1.Enabled = True
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
'reset variables and set new
Private Sub setState(ByVal mInd As Integer, ByVal mState As Byte)
    Dim i As Integer
    For i = 1 To 3
        bState(i) = 0
    Next i
    
    bState(mInd) = mState
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim scrSize As Integer, scrPos As Integer
    Dim ucSW As Integer, ucSH As Integer
    
    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    
    '
    'horizontal scroll
    If mStyle = stScrFedHor Then
        'find scroll size and position
        scrSize = ((ucSW - ucSH - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        If scrSize > (ucSW - ucSH - buttSpacing * 2) Then scrSize = (ucSW - ucSH - buttSpacing * 2)
        scrPos = ucSH / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSW - ucSH - buttSpacing * 2 - scrSize)
        'check is user click on decrease button
        If X >= 0 And X <= UserControl.ScaleHeight / 2 Then
            If bState(1) <> 1 And bState(1) <> 2 Then
                bState(1) = 1
                reDraw
                Timer2.Enabled = True
            End If
        'or on increase button
        ElseIf X <= UserControl.ScaleWidth And X >= UserControl.ScaleWidth - UserControl.ScaleHeight / 2 Then
            If bState(3) <> 1 And bState(3) <> 2 Then
                bState(3) = 1
                reDraw
                Timer2.Enabled = True
            End If
        'or is drag
        ElseIf isDrag = True Then
            'if true then find value from scroll position
            currV = (X - ucSH / 2 - buttSpacing * 2 - dragPos) * (maxV - minV) / (ucSW - ucSH - buttSpacing * 2 - scrSize) - minV
            'check is value in scroll area
            If currV < minV Then currV = minV
            If currV > maxV Then currV = maxV
            'setState 2, 1
            'Timer2.Enabled = True
            'reDraw
            RaiseEvent Scroll
            RaiseEvent Change
        'if mouse move on scroll then redraw hover scroll
        ElseIf X >= scrPos And X <= scrPos + scrSize Then
            If bState(2) <> 1 And bState(2) <> 2 Then
                setState 2, 1
                Timer2.Enabled = True
                reDraw
            End If
        'all reset and redraw
        Else
            setState 0, 0
            Timer2.Enabled = False
            reDraw
        End If
    'this is all same but there using y position
    ElseIf mStyle = stScrFedVert Then
        scrSize = ((ucSH - ucSW - buttSpacing * 2) / 2) / ((maxV - minV) / chStep)
        If scrSize > (ucSH - ucSW - buttSpacing * 2) Then scrSize = (ucSH - ucSW - buttSpacing * 2)
        scrPos = ucSW / 2 + buttSpacing + (currV - minV) / (maxV - minV) * (ucSH - ucSW - buttSpacing * 2 - scrSize)
        If Y >= 0 And Y <= ucSW / 2 Then
            If bState(1) <> 1 And bState(1) <> 2 Then
                bState(1) = 1
                Timer2.Enabled = True
                reDraw
            End If
        ElseIf Y <= ucSH And Y >= ucSH - ucSW / 2 Then
            If bState(3) <> 1 And bState(3) <> 2 Then
                bState(3) = 1
                Timer2.Enabled = True
                reDraw
            End If
        ElseIf isDrag = True Then
            currV = (Y - ucSW / 2 - buttSpacing * 2 - dragPos) * (maxV - minV) / (ucSH - ucSW - buttSpacing * 2 - scrSize) - minV
            If currV < minV Then currV = minV
            If currV > maxV Then currV = maxV
           ' MsgBox ""
            'setState 2, 1
            'Timer2.Enabled = True
            'reDraw
            RaiseEvent Scroll
            RaiseEvent Change
        ElseIf Y >= scrPos And Y <= scrPos + scrSize Then
            If bState(2) <> 1 And bState(2) <> 2 Then
                setState 2, 1
                Timer2.Enabled = True
                reDraw
            End If
        Else
            setState 0, 0
            Timer2.Enabled = False
            reDraw
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lpPos As POINTAPI
    If mStyle = stScrFedHor Then
        If X >= 0 And X <= UserControl.ScaleHeight / 2 Then
            setState 1, 1
        ElseIf X <= UserControl.ScaleWidth And X >= UserControl.ScaleWidth - UserControl.ScaleHeight / 2 Then
            setState 3, 1
        ElseIf isDrag = True Then
            setState 2, 1
        End If
    ElseIf mStyle = stScrFedVert Then
        If Y >= 0 And Y <= UserControl.ScaleWidth / 2 Then
            setState 1, 1
        ElseIf Y <= UserControl.ScaleHeight And Y >= UserControl.ScaleHeight - UserControl.ScaleWidth / 2 Then
            setState 3, 1
        ElseIf isDrag = True Then
            setState 2, 1
        End If
    End If
    
    GetCursorPos lpPos
    If WindowFromPoint(lpPos.X, lpPos.Y) <> UserControl.hWnd Then
        Timer2.Enabled = False
        setState 0, 0
        bDown = 0
        bHover = 0
    End If
    
    Timer1.Enabled = False
    bDown = 0
    isDrag = False
    reDraw
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BC = PropBag.ReadProperty("BackColor", vbBlack)
    TC = PropBag.ReadProperty("TransparenColor", vbBlack)
    '
    BC1 = PropBag.ReadProperty("ButtBackColor", &HEF8E4F)
    BC1Hover = PropBag.ReadProperty("ButtBackColorHover", &HFACAAE)
    BC1Down = PropBag.ReadProperty("ButtBackColorDown", &HAB6016)
    BC1Disabled = PropBag.ReadProperty("ButtBackColorDisabled", &H808080)
    '
    BC2 = PropBag.ReadProperty("ScrollBackColor", &HEF8E4F)
    BC2Hover = PropBag.ReadProperty("ScrollBackColorHover", &HFACAAE)
    BC2Down = PropBag.ReadProperty("ScrollBackColorDown", &HAB6016)
    BC2Disabled = PropBag.ReadProperty("ScrollBackColorDisabled", &H808080)
    '
    mStyle = PropBag.ReadProperty("Style", 0)
    '
    buttSpacing = PropBag.ReadProperty("ButtonSpacing", 5)
'    minV = PropBag.ReadProperty("Min", 0)
    maxV = PropBag.ReadProperty("Max", 100)
    currV = PropBag.ReadProperty("Value", 0)
    chStep = PropBag.ReadProperty("ChangeStep", 20)
    '
    enbl = PropBag.ReadProperty("Enabled", True)
    reDraw
End Sub

Private Sub UserControl_Resize()
    reDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BC, vbBlack
    PropBag.WriteProperty "TransparenColor", TC, vbBlack
    

    PropBag.WriteProperty "ButtBackColor", BC1, &HEF8E4F
    PropBag.WriteProperty "ButtBackColorHover", BC1Hover, &HFACAAE
    PropBag.WriteProperty "ButtBackColorDown", BC1Down, &HAB6016
    PropBag.WriteProperty "ButtBackColorDisabled", BC1Disabled, &H808080
    '
    PropBag.WriteProperty "ScrollBackColor", BC2, &HEF8E4F
    PropBag.WriteProperty "ScrollBackColorHover", BC2Hover, &HFACAAE
    PropBag.WriteProperty "ScrollBackColorDown", BC2Down, &HAB6016
    '
    PropBag.WriteProperty "ScrollBackColorDisabled", BC2Disabled, &H808080

    '
    '
    PropBag.WriteProperty "Style", mStyle, 0
    '
    PropBag.WriteProperty "ButtonSpacing", buttSpacing, 5
'    PropBag.WriteProperty "Min", minV, 0
    PropBag.WriteProperty "Max", maxV, 100
    PropBag.WriteProperty "Value", currV, 0
    PropBag.WriteProperty "ChangeStep", chStep, 20
    PropBag.WriteProperty "Enabled", enbl, True
End Sub

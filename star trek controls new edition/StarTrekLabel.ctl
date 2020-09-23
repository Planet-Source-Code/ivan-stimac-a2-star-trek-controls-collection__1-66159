VERSION 5.00
Begin VB.UserControl StarTrekLabel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekLabel.ctx":0000
End
Attribute VB_Name = "StarTrekLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'api
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'
Private BC As OLE_COLOR
Private BC1 As OLE_COLOR, BC2 As OLE_COLOR, BC1Disabled As OLE_COLOR, BC2Disabled As OLE_COLOR
Private BCLFT As OLE_COLOR, BCDisabledLFT As OLE_COLOR
Private BCRGHT As OLE_COLOR, BCDisabledRGHT As OLE_COLOR
Private FC As OLE_COLOR, FCDisabled As OLE_COLOR
Private TC As OLE_COLOR

Private mStyle As eSTLblStyle
Private mAlign As eSTCaptAlign

Private gradAngle As Single
Private enbl As Boolean, aFillSides As Boolean
Private ctlSpacingLFT As Integer, ctlSpacingRGHT As Integer

Private strCaption As String

Private cGrad As New clsGradient
'events
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'---------------------------------------------------
'---------------------------------------------------
'TransparentColor
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = TC
End Property
Public Property Let TransparentColor(ByVal nV As OLE_COLOR)
    TC = nV
    reDraw
    PropertyChanged "TransparentColor"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    reDraw
    PropertyChanged "BackColor"
End Property
'Back color middle 1
Public Property Get BackColorMiddle1() As OLE_COLOR
    BackColorMiddle1 = BC1
End Property
Public Property Let BackColorMiddle1(ByVal nV As OLE_COLOR)
    BC1 = nV
    reDraw
    PropertyChanged "BackColorMiddle1"
End Property
'Back color middle 2
Public Property Get BackColorMiddle2() As OLE_COLOR
    BackColorMiddle2 = BC2
End Property
Public Property Let BackColorMiddle2(ByVal nV As OLE_COLOR)
    BC2 = nV
    reDraw
    PropertyChanged "BackColorMiddle2"
End Property
'Back color middle 1 disabled
Public Property Get BackColorMiddle1Disabled() As OLE_COLOR
    BackColorMiddle1Disabled = BC1Disabled
End Property
Public Property Let BackColorMiddle1Disabled(ByVal nV As OLE_COLOR)
    BC1Disabled = nV
    reDraw
    PropertyChanged "BackColorMiddle1Disabled"
End Property
'Back color middle 2 disabled
Public Property Get BackColorMiddle2Disabled() As OLE_COLOR
    BackColorMiddle2Disabled = BC2Disabled
End Property
Public Property Let BackColorMiddle2Disabled(ByVal nV As OLE_COLOR)
    BC2Disabled = nV
    reDraw
    PropertyChanged "BackColorMiddle2Disabled"
End Property
'Back color left
Public Property Get BackColorLeft() As OLE_COLOR
    BackColorLeft = BCLFT
End Property
Public Property Let BackColorLeft(ByVal nV As OLE_COLOR)
    BCLFT = nV
    reDraw
    PropertyChanged "BackColorLeft"
End Property
'Back color left disabled
Public Property Get BackColorLeftDisabled() As OLE_COLOR
    BackColorLeftDisabled = BCDisabledLFT
End Property
Public Property Let BackColorLeftDisabled(ByVal nV As OLE_COLOR)
    BCDisabledLFT = nV
    reDraw
    PropertyChanged "BackColorLeftDisabled"
End Property
'Back color right
Public Property Get BackColorRight() As OLE_COLOR
    BackColorRight = BCRGHT
End Property
Public Property Let BackColorRight(ByVal nV As OLE_COLOR)
    BCRGHT = nV
    reDraw
    PropertyChanged "BackColorRight"
End Property
'Back color left disabled
Public Property Get BackColorRightDisabled() As OLE_COLOR
    BackColorRightDisabled = BCDisabledRGHT
End Property
Public Property Let BackColorRightDisabled(ByVal nV As OLE_COLOR)
    BCDisabledRGHT = nV
    reDraw
    PropertyChanged "BackColorRightDisabled"
End Property
'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    FC = nV
    reDraw
    PropertyChanged "ForeColor"
End Property
'fore color
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = FCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nV As OLE_COLOR)
    FCDisabled = nV
    reDraw
    PropertyChanged "ForeColorDisabled"
End Property

'-------ENUMS
'style
Public Property Get Style() As eSTLblStyle
    Style = mStyle
End Property
Public Property Let Style(ByVal nV As eSTLblStyle)
    mStyle = nV
    reDraw
    PropertyChanged "Style"
End Property
'caption align
Public Property Get Align() As eSTCaptAlign
    Align = mAlign
End Property
Public Property Let Align(ByVal nV As eSTCaptAlign)
    mAlign = nV
    reDraw
    PropertyChanged "Align"
End Property

'-------num values
'gradient angle
Public Property Get GradientAngle() As Single
    GradientAngle = gradAngle
End Property
Public Property Let GradientAngle(ByVal nV As Single)
    gradAngle = nV
    reDraw
    PropertyChanged "GradientAngle"
End Property
'SpacingLeft
Public Property Get SpacingLeft() As Integer
    SpacingLeft = ctlSpacingLFT
End Property
Public Property Let SpacingLeft(ByVal nV As Integer)
    ctlSpacingLFT = nV
    reDraw
    PropertyChanged "SpacingLeft"
End Property
'SpacingRight
Public Property Get SpacingRight() As Integer
    SpacingRight = ctlSpacingRGHT
End Property
Public Property Let SpacingRight(ByVal nV As Integer)
    ctlSpacingRGHT = nV
    reDraw
    PropertyChanged "SpacingRight"
End Property
'-------string values
'caption
Public Property Get Caption() As String
    Caption = strCaption
End Property
Public Property Let Caption(ByVal nV As String)
    strCaption = nV
    reDraw
    PropertyChanged "Caption"
End Property
'-------other
'font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal nV As Font)
    Set UserControl.Font = nV
'    Set Font = nV
    reDraw
    PropertyChanged "Font"
End Property
'enabled
Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property
'AutoFillSides
Public Property Get AutoFillSides() As Boolean
    AutoFillSides = aFillSides
End Property
Public Property Let AutoFillSides(ByVal nV As Boolean)
    aFillSides = nV
    reDraw
    PropertyChanged "AutoFillSides"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'---------------------------------------------------
'---------------------------------------------------
Private Sub UserControl_Initialize()
    BC = vbBlack
    TC = vbBlack
    
    BC1 = &HFF8080
    BC2 = &HFF0000
    
    BC1Disabled = &HFF8080
    BC2Disabled = &HFF8080
    
    BCLFT = &HFF8080
    BCDisabledLFT = &HFF8080
    
    BCRGHT = &HFF0000
    BCDisabledRGHT = &HFF8080
    
    FC = vbBlack
    FCDisabled = vbBlack
    
    gradAngle = 60
    mStyle = stLblFedRounded
    mAlign = Right
    enbl = True
    
    strCaption = "Label"
    
    aFillSides = False
    
    ctlSpacingLFT = 5
    ctlSpacingRGHT = 4
End Sub


Private Sub reDraw()
    UserControl.Cls
    UserControl.Enabled = enbl
    If mStyle = stLblFedRounded Or mStyle = stLblFedRightRounded Or mStyle = stLblFedLeftRounded Or mStyle = stLblFedRectangle Then
        drawRoundedLabel
    ElseIf mStyle = stLblRomulanDistort Or mStyle = stLblRomulanDistort2 Or mStyle = stLblRomulanMenu Then
        drawRomulanLabel
    ElseIf mStyle = stLblKlingonBottom Or mStyle = stLblKlingonTop Or Style = stLblKlingonTopLeft Or mStyle = stLblKlingonTopRight Or mStyle = stLblKlingonBottomLeft Or mStyle = stLblKlingonBottomRight Then
        drawKlingonLabel
    End If
    
    UserControl.MaskColor = TC
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
End Sub

'---------------------------------------------------------------------
'------------------- styles ------------------------------------------
'rounded and rectangle
Private Sub drawRoundedLabel()
    Dim ucSW As Integer, ucSH As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    
    If enbl = True Then
        mBC1 = BC1
        mBC2 = BC2
        mFC = FC
        mBC1LFT = BCLFT
        mBC1RGHT = BCRGHT
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + ctlSpacingLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - ctlSpacingRGHT, ucSH / 2)
    End If
    'draw left side
    If mStyle = stLblFedRounded Or mStyle = stLblFedLeftRounded Then
        'fill with back color
        UserControl.Line (0, 0)-(ctlSpacingLFT + 1 + ucSH / 2, ucSH), BC, BF
        'draw left side
        UserControl.Circle (ucSH / 2, ucSH / 2), ucSH / 2, mBC1LFT, 3.14 / 2, 1.5 * 3.14
        UserControl.Line (ucSH / 2, 0)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
    End If
    
    If mStyle = stLblFedRightRounded Or mStyle = stLblFedRounded Then
        'fill with back color
        UserControl.Line (ucSW - ctlSpacingRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
        'draw right side
        UserControl.Circle (ucSW - ucSH / 2, ucSH / 2), ucSH / 2, mBC1RGHT, 1.5 * 3.14, 3.14 / 2
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 2, UserControl.Point(ucSW - 2, ucSH / 2), 1
    End If
    
    If mAlign = Left And (mStyle = stLblFedLeftRounded Or mStyle = stLblFedRounded) Then
        UserControl.CurrentX = ucSH / 2 + ctlSpacingLFT + 5
    ElseIf mAlign = Right And (mStyle = stLblFedRightRounded Or mStyle = stLblFedRounded) Then
        UserControl.CurrentX = ucSW - ucSH / 2 - ctlSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Left Then
        UserControl.CurrentX = 5
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    
    UserControl.ForeColor = mFC
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub
' drawRomulanLabel
Private Sub drawRomulanLabel()
    Dim ucSW As Integer, ucSH As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    
    If enbl = True Then
        mBC1 = BC1
        mBC2 = BC2
        mFC = FC
        mBC1LFT = BCLFT
        mBC1RGHT = BCRGHT
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'Draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    '
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + ctlSpacingLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - ctlSpacingRGHT, ucSH / 2)
    End If
    '
    'fill with back color
    UserControl.Line (0, 0)-(ctlSpacingLFT + 1 + ucSH / 2, ucSH), BC, BF
    UserControl.Line (ucSW - ctlSpacingRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
    If mStyle = stLblRomulanDistort Then
        'left side
        UserControl.Line (0, ucSH / 1.5)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(0, ucSH / 1.5), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 1.5, UserControl.Point(2, ucSH / 2), 1
        'right side
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 4, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 4, 0)-(ucSW, ucSH - ucSH / 1.5), mBC1RGHT
        UserControl.Line (ucSW, ucSH - ucSH / 1.5)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH - ucSH / 1.5, UserControl.Point(2, ucSH / 2), 1
    ElseIf mStyle = stLblRomulanDistort2 Then
        'left side
        UserControl.Line (0, ucSH / 2 + ucSH / 4)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.Line (0, ucSH / 2 + ucSH / 4)-(ucSH / 2 - ucSH / 4, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2 - ucSH / 4, ucSH)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        'UserControl.Line (ucSH / 2 - ucSH / 6, 0)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 1.5, UserControl.Point(2, ucSH / 2), 1
        'right
        UserControl.Line (ucSW, ucSH / 2 + ucSH / 4)-(ucSW, ucSH / 2 - ucSH / 4), mBC1RGHT
        UserControl.Line (ucSW, ucSH / 2 + ucSH / 4)-(ucSW - ucSH / 2 + ucSH / 4, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2 + ucSH / 4, ucSH)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, ucSH)-(ucSW - ucSH / 2, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW, ucSH / 2 - ucSH / 4), mBC1RGHT
        'UserControl.Line (ucSH / 2 - ucSH / 6, 0)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 1.5, UserControl.Point(ucSW - 2, ucSH / 2), 1
    ElseIf mStyle = stLblRomulanMenu Then
        'left side
        UserControl.Line (0, ucSH / 2 + ucSH / 4)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.Line (0, ucSH / 2 + ucSH / 4)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
        'right
        UserControl.Line (ucSW, ucSH / 2 + ucSH / 4)-(ucSW, ucSH / 2 - ucSH / 4), mBC1RGHT
        UserControl.Line (ucSW, ucSH / 2 + ucSH / 4)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, ucSH)-(ucSW - ucSH / 2, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW, ucSH / 2 - ucSH / 4), mBC1RGHT
        'UserControl.Line (ucSH / 2 - ucSH / 6, 0)-(0, ucSH / 2 - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 2, UserControl.Point(ucSW - 2, ucSH / 2), 1
    End If
    
    
    If mAlign = Left Then
        UserControl.CurrentX = ucSH / 2 + ctlSpacingLFT + 5
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - ucSH / 2 - ctlSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    
    UserControl.ForeColor = mFC
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub
' drawKlingonLabel
Private Sub drawKlingonLabel()
    Dim ucSW As Integer, ucSH As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    
    If enbl = True Then
        mBC1 = BC1
        mBC2 = BC2
        mFC = FC
        mBC1LFT = BCLFT
        mBC1RGHT = BCRGHT
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'Draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    '
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + ctlSpacingLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - ctlSpacingRGHT, ucSH / 2)
    End If
    '
    'fill with back color
    If mStyle = stLblKlingonTop Or mStyle = stLblKlingonTopLeft Or mStyle = stLblKlingonTopRight Then
        UserControl.Line (0, 0)-(ucSW, ucSH / 8), BC, BF
    ElseIf mStyle = stLblKlingonBottom Or mStyle = stLblKlingonBottomLeft Or mStyle = stLblKlingonBottomRight Then
        UserControl.Line (0, ucSH)-(ucSW, ucSH - ucSH / 8), BC, BF
    End If
    
    
    If mStyle = stLblKlingonTop Or mStyle = stLblKlingonTopLeft Then
        UserControl.Line (0, 0)-(ctlSpacingLFT + 1 + ucSH / 2, ucSH), BC, BF
        'draw
        UserControl.Line (0, ucSH - ucSH / 4)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 4, ucSH), mBC1LFT
        UserControl.Line (ucSH / 4, ucSH)-(0, ucSH - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH - ucSH / 4, UserControl.Point(2, ucSH - ucSH / 4), 1
    End If
    If mStyle = stLblKlingonTop Or mStyle = stLblKlingonTopRight Then
        UserControl.Line (ucSW - ctlSpacingRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
        'draw
        UserControl.Line (ucSW, ucSH - ucSH / 4)-(ucSW - ucSH / 2, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, ucSH)-(ucSW - ucSH / 4, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 4, ucSH)-(ucSW, ucSH - ucSH / 4), mBC1RGHT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH - ucSH / 4, UserControl.Point(ucSW - 2, ucSH - ucSH / 4), 1
    End If
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 + ucSH / 16
    'draw bottom
    If mStyle = stLblKlingonBottom Or mStyle = stLblKlingonBottomLeft Then
        UserControl.Line (0, ucSH)-(ctlSpacingLFT + 1 + ucSH / 2, 0), BC, BF
        'draw
        UserControl.Line (0, ucSH / 4)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(ucSH / 4, 0), mBC1LFT
        UserControl.Line (ucSH / 4, 0)-(0, ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 4, UserControl.Point(2, ucSH / 4), 1
        
        UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 - ucSH / 16
    End If
    If mStyle = stLblKlingonBottom Or mStyle = stLblKlingonBottomRight Then
        UserControl.Line (ucSW - ctlSpacingRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
        'draw
        UserControl.Line (ucSW, ucSH / 4)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, ucSH)-(ucSW - ucSH / 2, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 4, 0), mBC1RGHT
        UserControl.Line (ucSW - ucSH / 4, 0)-(ucSW, ucSH / 4), mBC1RGHT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 4, UserControl.Point(ucSW - 2, ucSH / 4), 1
        
        UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 - ucSH / 16
    End If
    
    If mAlign = Left Then
        UserControl.CurrentX = ucSH / 2 + ctlSpacingLFT + 5
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - ucSH / 2 - ctlSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    'UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 - ucSH / 16
    UserControl.ForeColor = mFC
    
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'---------------------------------------------------------------------
'------------------- user control ------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    BC = PropBag.ReadProperty("BackColor", vbBlack)
    TC = PropBag.ReadProperty("TransparentColor", vbBlack)
    
    BC1 = PropBag.ReadProperty("BackColorMiddle1", &HFF8080)
    BC2 = PropBag.ReadProperty("BackColorMiddle2", &HFF0000)
    BC1Disabled = PropBag.ReadProperty("BackColorMiddle1Disabled", &HFF0000)
    BC2Disabled = PropBag.ReadProperty("BackColorMiddle2Disabled", &HFF0000)
    
    BCLFT = PropBag.ReadProperty("BackColorLeft", &HFF8080)
    BCDisabledLFT = PropBag.ReadProperty("BackColorLeftDisabled", &HFF0000)
    BCRGHT = PropBag.ReadProperty("BackColorRight", &HFF0000)
    BCDisabledRGHT = PropBag.ReadProperty("BackColorRightDisabled", &HFF0000)
    
    
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    
    
    mStyle = PropBag.ReadProperty("Style", 0)
    mAlign = PropBag.ReadProperty("Align", 0)
    
    gradAngle = PropBag.ReadProperty("GradientAngle", 60)
    ctlSpacingLFT = PropBag.ReadProperty("SpacingLeft", 5)
    ctlSpacingRGHT = PropBag.ReadProperty("SpacingRight", 5)
    
    strCaption = PropBag.ReadProperty("Caption", "Label")
    
    enbl = PropBag.ReadProperty("Enabled", True)
    aFillSides = PropBag.ReadProperty("AutoFillSides", False)
    reDraw
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    
    PropBag.WriteProperty "BackColor", BC, vbBlack
    PropBag.WriteProperty "TransparentColor", TC, vbBlack
    
    PropBag.WriteProperty "BackColorMiddle1", BC1, &HFF8080
    PropBag.WriteProperty "BackColorMiddle2", BC2, &HFF0000
    PropBag.WriteProperty "BackColorMiddle1Disabled", BC1Disabled, &HFF0000
    PropBag.WriteProperty "BackColorMiddle2Disabled", BC2Disabled, &HFF0000
    
    PropBag.WriteProperty "BackColorLeft", BCLFT, &HFF8080
    PropBag.WriteProperty "BackColorLeftDisabled", BCDisabledLFT, &HFF0000
    
    PropBag.WriteProperty "BackColorRight", BCRGHT, &HFF0000
    PropBag.WriteProperty "BackColorRightDisabled", BCDisabledRGHT, &HFF0000

    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    
    PropBag.WriteProperty "Style", mStyle, 0
    PropBag.WriteProperty "Align", mAlign, 0
    

    
    PropBag.WriteProperty "GradientAngle", gradAngle, 60
    PropBag.WriteProperty "SpacingLeft", ctlSpacingLFT, 5
    PropBag.WriteProperty "SpacingRight", ctlSpacingRGHT, 5
    
    
    PropBag.WriteProperty "Caption", strCaption, "label"
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "AutoFillSides", aFillSides, False
End Sub
Private Sub UserControl_Resize()
    reDraw
End Sub



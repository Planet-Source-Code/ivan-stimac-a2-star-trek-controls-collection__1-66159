VERSION 5.00
Begin VB.UserControl StarTrekButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1260
      Top             =   960
   End
End
Attribute VB_Name = "StarTrekButton"
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
Private BC As OLE_COLOR
Private BC1 As OLE_COLOR, BC2 As OLE_COLOR, BC1Disabled As OLE_COLOR, BC2Disabled As OLE_COLOR, BC1Hover As OLE_COLOR, BC2Hover As OLE_COLOR, BC1Down As OLE_COLOR, BC2Down As OLE_COLOR
Private BCLFT As OLE_COLOR, BCDisabledLFT As OLE_COLOR, BCLFTHover As OLE_COLOR, BCLFTDown As OLE_COLOR

Private BCRGHT As OLE_COLOR, BCDisabledRGHT As OLE_COLOR, BCRGHTHover As OLE_COLOR, BCRGHTDown As OLE_COLOR
Private FC As OLE_COLOR, FCDisabled As OLE_COLOR, FCHover As OLE_COLOR, FCDown As OLE_COLOR
Private TC As OLE_COLOR

Private mStyle As eSTButtStyle
Private mAlign As eSTCaptAlign

Private gradAngle As Single
Private enbl As Boolean, aFillSides As Boolean
Private ctlSpacingLFT As Integer, ctlSpacingRGHT As Integer
Private ctlSpacingLFTHover As Integer, ctlSpacingRGHTHover As Integer
Private ctlSpacingLFTDown As Integer, ctlSpacingRGHTDown As Integer
Private ctlSpacingLFTDisabled As Integer, ctlSpacingRGHTDisabled As Integer

Private strCaption As String

Private cGrad As New clsGradient

'0=normal: 1= hover: 2=down
Private bState As Byte


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
'Back color middle 1 hover
Public Property Get BackColorMiddle1Hover() As OLE_COLOR
    BackColorMiddle1Hover = BC1Hover
End Property
Public Property Let BackColorMiddle1Hover(ByVal nV As OLE_COLOR)
    BC1Hover = nV
    reDraw
    PropertyChanged "BackColorMiddle1Hover"
End Property
'Back color middle 1 down
Public Property Get BackColorMiddle1Down() As OLE_COLOR
    BackColorMiddle1Down = BC1Down
End Property
Public Property Let BackColorMiddle1Down(ByVal nV As OLE_COLOR)
    BC1Down = nV
    reDraw
    PropertyChanged "BackColorMiddle1Down"
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
'Back color middle 2 hover
Public Property Get BackColorMiddle2Hover() As OLE_COLOR
    BackColorMiddle2Hover = BC2Hover
End Property
Public Property Let BackColorMiddle2Hover(ByVal nV As OLE_COLOR)
    BC2Hover = nV
    reDraw
    PropertyChanged "BackColorMiddle2Hover"
End Property
'Back color middle 2 down
Public Property Get BackColorMiddle2Down() As OLE_COLOR
    BackColorMiddle2Down = BC2Down
End Property
Public Property Let BackColorMiddle2Down(ByVal nV As OLE_COLOR)
    BC2Down = nV
    reDraw
    PropertyChanged "BackColorMiddle2Down"
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
'Back color left hover
Public Property Get BackColorLeftHover() As OLE_COLOR
    BackColorLeftHover = BCLFTHover
End Property
Public Property Let BackColorLeftHover(ByVal nV As OLE_COLOR)
    BCLFTHover = nV
    reDraw
    PropertyChanged "BackColorLeftHover"
End Property
'Back color left down
Public Property Get BackColorLeftDown() As OLE_COLOR
    BackColorLeftDown = BCLFTDown
End Property
Public Property Let BackColorLeftDown(ByVal nV As OLE_COLOR)
    BCLFTDown = nV
    reDraw
    PropertyChanged "BackColorLeftDown"
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
'Back color right hover
Public Property Get BackColorRightHover() As OLE_COLOR
    BackColorRightHover = BCRGHTHover
End Property
Public Property Let BackColorRightHover(ByVal nV As OLE_COLOR)
    BCRGHTHover = nV
    reDraw
    PropertyChanged "BackColorRightHover"
End Property
'Back color right down
Public Property Get BackColorRightDown() As OLE_COLOR
    BackColorRightDown = BCRGHTDown
End Property
Public Property Let BackColorRightDown(ByVal nV As OLE_COLOR)
    BCRGHTDown = nV
    reDraw
    PropertyChanged "BackColorRightDown"
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
'fore color hover
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = FCHover
End Property
Public Property Let ForeColorHover(ByVal nV As OLE_COLOR)
    FCHover = nV
    reDraw
    PropertyChanged "ForeColorHover"
End Property
'fore color down
Public Property Get ForeColorDown() As OLE_COLOR
    ForeColorDown = FCDown
End Property
Public Property Let ForeColorDown(ByVal nV As OLE_COLOR)
    FCDown = nV
    reDraw
    PropertyChanged "ForeColorDown"
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
Public Property Get Style() As eSTButtStyle
    Style = mStyle
End Property
Public Property Let Style(ByVal nV As eSTButtStyle)
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
'SpacingLeft hover
Public Property Get SpacingLeftHover() As Integer
    SpacingLeftHover = ctlSpacingLFTHover
End Property
Public Property Let SpacingLeftHover(ByVal nV As Integer)
    ctlSpacingLFTHover = nV
    reDraw
    PropertyChanged "SpacingLeftHover"
End Property
'SpacingLeft down
Public Property Get SpacingLeftDown() As Integer
    SpacingLeftDown = ctlSpacingLFTDown
End Property
Public Property Let SpacingLeftDown(ByVal nV As Integer)
    ctlSpacingLFTDown = nV
    reDraw
    PropertyChanged "SpacingLeftDown"
End Property
'SpacingLeft disabled
Public Property Get SpacingLeftDisabled() As Integer
    SpacingLeftDisabled = ctlSpacingLFTDisabled
End Property
Public Property Let SpacingLeftDisabled(ByVal nV As Integer)
    ctlSpacingLFTDisabled = nV
    reDraw
    PropertyChanged "SpacingLeftDisabled"
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
'SpacingRight hover
Public Property Get SpacingRightHover() As Integer
    SpacingRightHover = ctlSpacingRGHTHover
End Property
Public Property Let SpacingRightHover(ByVal nV As Integer)
    ctlSpacingRGHTHover = nV
    reDraw
    PropertyChanged "SpacingRightHover"
End Property
'SpacingRight down
Public Property Get SpacingRightDown() As Integer
    SpacingRightDown = ctlSpacingRGHTDown
End Property
Public Property Let SpacingRightDown(ByVal nV As Integer)
    ctlSpacingRGHTDown = nV
    reDraw
    PropertyChanged "SpacingRightDown"
End Property
'SpacingRight disabled
Public Property Get SpacingRightDisabled() As Integer
    SpacingRightDisabled = ctlSpacingRGHTDisabled
End Property
Public Property Let SpacingRightDisabled(ByVal nV As Integer)
    ctlSpacingRGHTDisabled = nV
    reDraw
    PropertyChanged "SpacingRightDisabled"
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



Private Sub Timer1_Timer()
    If bState <> 0 Then
        Dim lpPos As POINTAPI
        GetCursorPos lpPos
        If WindowFromPoint(lpPos.X, lpPos.Y) <> UserControl.hWnd Then
            bState = 0
            reDraw
            Timer1.Enabled = False
        End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'---------------------------------------------------
'---------------------------------------------------
Private Sub UserControl_Initialize()
    BC1 = &HFF8080
    BC2 = &HFF0000
    BC1Hover = &HFF8080
    BC2Hover = &HFF0000
    BC1Down = &HFF0000
    BC2Down = &HFF0000
    BC1Disabled = &H808080
    BC2Disabled = &H808080
    
    BC = vbBlack
    TC = vbBlack
    
    BCLFT = &HFF8080
    BCDisabledLFT = &H808080
    BCLFTHover = &HFF8080
    BCLFTDown = &HFF8080
    
    BCRGHT = &HFF0000
    BCDisabledRGHT = &H808080
    BCRGHTHover = &HFF0000
    BCRGHTDown = &HFF0000
    
    FC = vbBlack
    FCDisabled = &H404040
    FCHover = vbBlack
    FCDown = vbBlack
    
    gradAngle = 60
    mStyle = stLblFedRounded
    mAlign = Right
    enbl = True
    aFillSides = False
    
    strCaption = "Button"
    
    ctlSpacingLFT = 5
    ctlSpacingLFTHover = 5
    ctlSpacingLFTDown = 5
    
    ctlSpacingRGHT = 5
    ctlSpacingRGHTHover = 5
    ctlSpacingRGHTDown = 5
    
    bState = 0
    
    
    reDraw
End Sub


Private Sub reDraw()
    UserControl.Cls
    UserControl.Enabled = enbl
    If mStyle = stButtFedRounded Or mStyle = stButtFedRightRounded Or mStyle = stButtFedLeftRounded Or mStyle = stButtFedRectangle Then
        drawRoundedButton
    ElseIf mStyle = stButtRomulanDistort Or mStyle = stButtRomulanDistort2 Or mStyle = stButtRomulanMenu Then
        drawRomulanButton
    ElseIf mStyle = stButtKlingonBottom Or mStyle = stButtKlingonTop Or Style = stButtKlingonTopLeft Or mStyle = stButtKlingonTopRight Or mStyle = stButtKlingonBottomLeft Or mStyle = stButtKlingonBottomRight Then
        drawKlingonButton
    End If
    
    UserControl.MaskColor = TC
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
    'createTransparent UserControl.hWnd, BC
End Sub

'---------------------------------------------------------------------
'------------------- styles ------------------------------------------
'rounded and rectangle
Private Sub drawRoundedButton()
    Dim ucSW As Integer, ucSH As Integer
    Dim mSPCLFT As Integer, mSPCRGHT As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    Dim txSpacingLFT As Integer, txSpacingRGHT As Integer
    
    txSpacingLFT = ctlSpacingLFT
    
    If txSpacingLFT < ctlSpacingLFTDown Then
        txSpacingLFT = ctlSpacingLFTDown
    End If
    If txSpacingLFT < ctlSpacingLFTHover Then
        txSpacingLFT = ctlSpacingLFTHover
    End If
    If txSpacingLFT < ctlSpacingLFTDisabled Then
        txSpacingLFT = ctlSpacingLFTDisabled
    End If
    
    txSpacingRGHT = ctlSpacingRGHT
    If txSpacingRGHT < ctlSpacingRGHTDown Then
        txSpacingRGHT = ctlSpacingRGHTDown
    End If
    If txSpacingRGHT < ctlSpacingRGHTHover Then
        txSpacingRGHT = ctlSpacingRGHTHover
    End If
    If txSpacingRGHT < ctlSpacingRGHTDisabled Then
        txSpacingRGHT = ctlSpacingRGHTDisabled
    End If
    
    If enbl = True Then
        If bState = 0 Then
            mBC1 = BC1
            mBC2 = BC2
            mFC = FC
            mBC1LFT = BCLFT
            mBC1RGHT = BCRGHT
            
            mSPCLFT = ctlSpacingLFT
            mSPCRGHT = ctlSpacingRGHT
        ElseIf bState = 1 Then
            mBC1 = BC1Hover
            mBC2 = BC2Hover
            mFC = FCHover
            mBC1LFT = BCLFTHover
            mBC1RGHT = BCRGHTHover
            
            mSPCLFT = ctlSpacingLFTHover
            mSPCRGHT = ctlSpacingRGHTHover
        ElseIf bState = 2 Then
            mBC1 = BC1Down
            mBC2 = BC2Down
            mFC = FCDown
            mBC1LFT = BCLFTDown
            mBC1RGHT = BCRGHTDown
            
            mSPCLFT = ctlSpacingLFTDown
            mSPCRGHT = ctlSpacingRGHTDown
        End If
        
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
        
        mSPCLFT = ctlSpacingLFTDisabled
        mSPCRGHT = ctlSpacingRGHTDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + mSPCLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - mSPCRGHT, ucSH / 2)
    End If
        
    'draw left side
    If mStyle = stButtFedRounded Or mStyle = stButtFedLeftRounded Then
        'fill with back color
        UserControl.Line (0, 0)-(mSPCLFT + 1 + ucSH / 2, ucSH), BC, BF
        'draw left side
        UserControl.Circle (ucSH / 2, ucSH / 2), ucSH / 2, mBC1LFT, 3.14 / 2, 1.5 * 3.14
        UserControl.Line (ucSH / 2, 0)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
    End If
    
    If mStyle = stLblFedRightRounded Or mStyle = stLblFedRounded Then
        'fill with back color
        UserControl.Line (ucSW - mSPCRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
        'draw left side
        UserControl.Circle (ucSW - ucSH / 2, ucSH / 2), ucSH / 2, mBC1RGHT, 1.5 * 3.14, 3.14 / 2
        UserControl.Line (ucSW - ucSH / 2, 0)-(ucSW - ucSH / 2, ucSH), mBC1RGHT
        UserControl.FillColor = mBC1RGHT
        ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 2, UserControl.Point(ucSW - 2, ucSH / 2), 1
    End If
    
'    If mAlign = Left And (mStyle = stLblFedLeftRounded Or mStyle = stLblFedRounded) Then
'        UserControl.CurrentX = ucSH / 2 + mSPCLFT + 5
'    ElseIf mAlign = Right And (mStyle = stLblFedRightRounded Or mStyle = stLblFedRounded) Then
'        UserControl.CurrentX = ucSW - ucSH / 2 - mSPCRGHT - 5 - UserControl.TextWidth(strCaption)
    If mAlign = Left Then
        UserControl.CurrentX = ucSH / 2 + 5 + txSpacingLFT
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - ucSH / 2 - txSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    
    UserControl.ForeColor = mFC
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub
' imagination
Private Sub drawRomulanButton()
    Dim ucSW As Integer, ucSH As Integer
    Dim mSPCLFT As Integer, mSPCRGHT As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    Dim txSpacingLFT As Integer, txSpacingRGHT As Integer
    
    txSpacingLFT = ctlSpacingLFT
    
    If txSpacingLFT < ctlSpacingLFTDown Then
        txSpacingLFT = ctlSpacingLFTDown
    End If
    If txSpacingLFT < ctlSpacingLFTHover Then
        txSpacingLFT = ctlSpacingLFTHover
    End If
    If txSpacingLFT < ctlSpacingLFTDisabled Then
        txSpacingLFT = ctlSpacingLFTDisabled
    End If
    
    txSpacingRGHT = ctlSpacingRGHT
    If txSpacingRGHT < ctlSpacingRGHTDown Then
        txSpacingRGHT = ctlSpacingRGHTDown
    End If
    If txSpacingRGHT < ctlSpacingRGHTHover Then
        txSpacingRGHT = ctlSpacingRGHTHover
    End If
    If txSpacingRGHT < ctlSpacingRGHTDisabled Then
        txSpacingRGHT = ctlSpacingRGHTDisabled
    End If
    
    
    If enbl = True Then
        If bState = 0 Then
            mBC1 = BC1
            mBC2 = BC2
            mFC = FC
            mBC1LFT = BCLFT
            mBC1RGHT = BCRGHT
            
            mSPCLFT = ctlSpacingLFT
            mSPCRGHT = ctlSpacingRGHT
        ElseIf bState = 1 Then
            mBC1 = BC1Hover
            mBC2 = BC2Hover
            mFC = FCHover
            mBC1LFT = BCLFTHover
            mBC1RGHT = BCRGHTHover
            
            mSPCLFT = ctlSpacingLFTHover
            mSPCRGHT = ctlSpacingRGHTHover
        ElseIf bState = 2 Then
            mBC1 = BC1Down
            mBC2 = BC2Down
            mFC = FCDown
            mBC1LFT = BCLFTDown
            mBC1RGHT = BCRGHTDown
            
            mSPCLFT = ctlSpacingLFTDown
            mSPCRGHT = ctlSpacingRGHTDown
        End If
        
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
        
        mSPCLFT = ctlSpacingLFTDisabled
        mSPCRGHT = ctlSpacingRGHTDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'Draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    '
    '
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + mSPCLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - mSPCRGHT, ucSH / 2)
    End If
    'fill with back color
    UserControl.Line (0, 0)-(mSPCLFT + 1 + ucSH / 2, ucSH), BC, BF
    UserControl.Line (ucSW - mSPCRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
    If mStyle = stButtRomulanDistort Then
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
    ElseIf mStyle = stButtRomulanDistort2 Then
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
    ElseIf mStyle = stButtRomulanMenu Then
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
        UserControl.CurrentX = ucSH / 2 + txSpacingLFT + 5
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - ucSH / 2 - txSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    
    UserControl.ForeColor = mFC
    UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub

' drawKlingonButton
Private Sub drawKlingonButton()
    Dim ucSW As Integer, ucSH As Integer
    Dim mSPCLFT As Integer, mSPCRGHT As Integer
    Dim mBC1 As OLE_COLOR, mBC2 As OLE_COLOR
    Dim mBC1LFT As OLE_COLOR
    Dim mBC1RGHT As OLE_COLOR
    Dim mFC As OLE_COLOR
    Dim txSpacingLFT As Integer, txSpacingRGHT As Integer
    
    txSpacingLFT = ctlSpacingLFT
    
    If txSpacingLFT < ctlSpacingLFTDown Then
        txSpacingLFT = ctlSpacingLFTDown
    End If
    If txSpacingLFT < ctlSpacingLFTHover Then
        txSpacingLFT = ctlSpacingLFTHover
    End If
    If txSpacingLFT < ctlSpacingLFTDisabled Then
        txSpacingLFT = ctlSpacingLFTDisabled
    End If
    
    txSpacingRGHT = ctlSpacingRGHT
    If txSpacingRGHT < ctlSpacingRGHTDown Then
        txSpacingRGHT = ctlSpacingRGHTDown
    End If
    If txSpacingRGHT < ctlSpacingRGHTHover Then
        txSpacingRGHT = ctlSpacingRGHTHover
    End If
    If txSpacingRGHT < ctlSpacingRGHTDisabled Then
        txSpacingRGHT = ctlSpacingRGHTDisabled
    End If
    
    
    If enbl = True Then
        If bState = 0 Then
            mBC1 = BC1
            mBC2 = BC2
            mFC = FC
            mBC1LFT = BCLFT
            mBC1RGHT = BCRGHT
            
            mSPCLFT = ctlSpacingLFT
            mSPCRGHT = ctlSpacingRGHT
        ElseIf bState = 1 Then
            mBC1 = BC1Hover
            mBC2 = BC2Hover
            mFC = FCHover
            mBC1LFT = BCLFTHover
            mBC1RGHT = BCRGHTHover
            
            mSPCLFT = ctlSpacingLFTHover
            mSPCRGHT = ctlSpacingRGHTHover
        ElseIf bState = 2 Then
            mBC1 = BC1Down
            mBC2 = BC2Down
            mFC = FCDown
            mBC1LFT = BCLFTDown
            mBC1RGHT = BCRGHTDown
            
            mSPCLFT = ctlSpacingLFTDown
            mSPCRGHT = ctlSpacingRGHTDown
        End If
        
    Else
        mBC1 = BC1Disabled
        mBC2 = BC2Disabled
        mBC1LFT = BCDisabledLFT
        
        mBC1RGHT = BCDisabledRGHT
        
        mFC = FCDisabled
        
        mSPCLFT = ctlSpacingLFTDisabled
        mSPCRGHT = ctlSpacingRGHTDisabled
    End If
        

    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    'Draw middle
    cGrad.Color1 = mBC1: cGrad.Color2 = mBC2
    cGrad.Angle = gradAngle
    cGrad.Draw UserControl.hWnd, UserControl.hDC, ucSH / 2, 0
    '
    If aFillSides = True Then
        mBC1LFT = UserControl.Point(ucSH / 2 + mSPCLFT + 1, ucSH / 2)
        mBC1RGHT = UserControl.Point(ucSW - ucSH / 2 - mSPCRGHT, ucSH / 2)
    End If
    '
    'fill with back color
    If mStyle = stButtKlingonTop Or mStyle = stButtKlingonTopLeft Or mStyle = stButtKlingonTopRight Then
        UserControl.Line (0, 0)-(ucSW, ucSH / 8), BC, BF
    ElseIf mStyle = stButtKlingonBottom Or mStyle = stButtKlingonBottomLeft Or mStyle = stButtKlingonBottomRight Then
        UserControl.Line (0, ucSH)-(ucSW, ucSH - ucSH / 8), BC, BF
    End If
    
    
    If mStyle = stButtKlingonTop Or mStyle = stButtKlingonTopLeft Then
        UserControl.Line (0, 0)-(mSPCLFT + 1 + ucSH / 2, ucSH), BC, BF
        'draw
        UserControl.Line (0, ucSH - ucSH / 4)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 4, ucSH), mBC1LFT
        UserControl.Line (ucSH / 4, ucSH)-(0, ucSH - ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH - ucSH / 4, UserControl.Point(2, ucSH - ucSH / 4), 1
    End If
    If mStyle = stButtKlingonTop Or mStyle = stButtKlingonTopRight Then
        UserControl.Line (ucSW - mSPCRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
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
    If mStyle = stButtKlingonBottom Or mStyle = stButtKlingonBottomLeft Then
        UserControl.Line (0, ucSH)-(mSPCLFT + 1 + ucSH / 2, 0), BC, BF
        'draw
        UserControl.Line (0, ucSH / 4)-(ucSH / 2, ucSH), mBC1LFT
        UserControl.Line (ucSH / 2, ucSH)-(ucSH / 2, 0), mBC1LFT
        UserControl.Line (ucSH / 2, 0)-(ucSH / 4, 0), mBC1LFT
        UserControl.Line (ucSH / 4, 0)-(0, ucSH / 4), mBC1LFT
        UserControl.FillColor = mBC1LFT
        ExtFloodFill UserControl.hDC, 2, ucSH / 4, UserControl.Point(2, ucSH / 4), 1
        
        UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 - ucSH / 16
    End If
    If mStyle = stButtKlingonBottom Or mStyle = stButtKlingonBottomRight Then
        UserControl.Line (ucSW - mSPCRGHT - ucSH / 2, 0)-(ucSW, ucSH), BC, BF
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
        UserControl.CurrentX = ucSH / 2 + txSpacingLFT + 5
    ElseIf mAlign = Right Then
        UserControl.CurrentX = ucSW - ucSH / 2 - txSpacingRGHT - 5 - UserControl.TextWidth(strCaption)
    ElseIf mAlign = Center Then
        UserControl.CurrentX = ucSW / 2 - UserControl.TextWidth(strCaption) / 2
    End If
    'UserControl.CurrentY = ucSH / 2 - UserControl.TextHeight(strCaption) / 2 - ucSH / 16
    UserControl.ForeColor = mFC
    
    UserControl.Print strCaption
    
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bState <> 2 Then
        bState = 2
        reDraw
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bState <> 1 And bState <> 2 Then
        bState = 1
        reDraw
        Timer1.Enabled = True
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lpPos As POINTAPI
    GetCursorPos lpPos
    If WindowFromPoint(lpPos.X, lpPos.Y) <> UserControl.hWnd Then
        bState = 0
        reDraw
    Else
        bState = 1
        reDraw
    End If
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
    BC1Hover = PropBag.ReadProperty("BackColorMiddle1Hover", &HFF8080)
    BC2Hover = PropBag.ReadProperty("BackColorMiddle2Hover", &HFF0000)
    BC1Down = PropBag.ReadProperty("BackColorMiddle1Down", &HFF8080)
    BC2Down = PropBag.ReadProperty("BackColorMiddle2Down", &HFF0000)
    BC1Disabled = PropBag.ReadProperty("BackColorMiddle1Disabled", &HFF0000)
    BC2Disabled = PropBag.ReadProperty("BackColorMiddle2Disabled", &HFF0000)
    
    BCLFT = PropBag.ReadProperty("BackColorLeft", &HFF8080)
    BCLFTHover = PropBag.ReadProperty("BackColorLeftHover", &HFF8080)
    BCLFTDown = PropBag.ReadProperty("BackColorLeftDown", &HFF8080)
    BCDisabledLFT = PropBag.ReadProperty("BackColorLeftDisabled", &HFF0000)
    
    BCRGHTHover = PropBag.ReadProperty("BackColorRightHover", &HFF0000)
    BCRGHTDown = PropBag.ReadProperty("BackColorRightDown", &HFF0000)
    BCRGHT = PropBag.ReadProperty("BackColorRight", &HFF0000)
    BCDisabledRGHT = PropBag.ReadProperty("BackColorRightDisabled", &HFF0000)
    
    
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCHover = PropBag.ReadProperty("ForeColorHover", vbBlack)
    FCDown = PropBag.ReadProperty("ForeColorDown", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    
    
    mStyle = PropBag.ReadProperty("Style", 0)
    mAlign = PropBag.ReadProperty("Align", 0)
    
    gradAngle = PropBag.ReadProperty("GradientAngle", 60)
    ctlSpacingLFT = PropBag.ReadProperty("SpacingLeft", 5)
    ctlSpacingLFTHover = PropBag.ReadProperty("SpacingLeftHover", 5)
    ctlSpacingLFTDown = PropBag.ReadProperty("SpacingLeftDown", 5)
    ctlSpacingLFTDisabled = PropBag.ReadProperty("SpacingLeftDisabled", 5)
    
    ctlSpacingRGHT = PropBag.ReadProperty("SpacingRight", 5)
    ctlSpacingRGHTHover = PropBag.ReadProperty("SpacingRightHover", 5)
    ctlSpacingRGHTDown = PropBag.ReadProperty("SpacingRightDown", 5)
    ctlSpacingRGHTDisabled = PropBag.ReadProperty("SpacingRightDisabled", 5)
    
    strCaption = PropBag.ReadProperty("Caption", "Button")
    
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
    PropBag.WriteProperty "BackColorMiddle1Hover", BC1Hover, &HFF8080
    PropBag.WriteProperty "BackColorMiddle2Hover", BC2Hover, &HFF0000
    PropBag.WriteProperty "BackColorMiddle1Down", BC1Down, &HFF8080
    PropBag.WriteProperty "BackColorMiddle2down", BC2Down, &HFF0000
    PropBag.WriteProperty "BackColorMiddle1Disabled", BC1Disabled, &HFF0000
    PropBag.WriteProperty "BackColorMiddle2Disabled", BC2Disabled, &HFF0000
    
    PropBag.WriteProperty "BackColorLeft", BCLFT, &HFF8080
    PropBag.WriteProperty "BackColorLeftHover", BCLFTHover, &HFF8080
    PropBag.WriteProperty "BackColorLeftDown", BCLFTDown, &HFF8080
    PropBag.WriteProperty "BackColorLeftDisabled", BCDisabledLFT, &HFF0000
    
    PropBag.WriteProperty "BackColorRight", BCRGHT, &HFF0000
    PropBag.WriteProperty "BackColorRightHover", BCRGHTHover, &HFF0000
    PropBag.WriteProperty "BackColorRightDown", BCRGHTDown, &HFF0000
    PropBag.WriteProperty "BackColorRightDisabled", BCDisabledRGHT, &HFF0000

    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorHover", FCHover, vbBlack
    PropBag.WriteProperty "ForeColorDown", FCDown, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    
    PropBag.WriteProperty "Style", mStyle, 0
    PropBag.WriteProperty "Align", mAlign, 0
    

    
    PropBag.WriteProperty "GradientAngle", gradAngle, 60
    PropBag.WriteProperty "SpacingLeft", ctlSpacingLFT, 5
    PropBag.WriteProperty "SpacingLeftHover", ctlSpacingLFTHover, 5
    PropBag.WriteProperty "SpacingLeftDown", ctlSpacingLFTDown, 5
    PropBag.WriteProperty "SpacingLeftDisabled", ctlSpacingLFTDisabled, 5
    
    PropBag.WriteProperty "SpacingRight", ctlSpacingRGHT, 5
    PropBag.WriteProperty "SpacingRightHover", ctlSpacingRGHTHover, 5
    PropBag.WriteProperty "SpacingRightDown", ctlSpacingRGHTDown, 5
    PropBag.WriteProperty "SpacingRightDisabled", ctlSpacingRGHTDisabled, 5
    
    PropBag.WriteProperty "Caption", strCaption, "Button"
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "AutoFillSides", aFillSides, False
End Sub
Private Sub UserControl_Resize()
    reDraw
End Sub




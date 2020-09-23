VERSION 5.00
Begin VB.UserControl StarTrekShape 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekShape.ctx":0000
End
Attribute VB_Name = "StarTrekShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long


Private BC As OLE_COLOR, TC As OLE_COLOR
Private BC1 As OLE_COLOR

Private mWidFixed As Integer, mHeigFixed As Integer
Private lnWid As Integer
Private lnAngle As Integer

Private mShape As eSTShape
Private mVOrj As eSTVertOrj
Private mHOrj As eSTHorOrj

'LineWidth
Public Property Get LineAngle() As Integer
    LineAngle = lnAngle
End Property
Public Property Let LineAngle(ByVal nV As Integer)
    lnAngle = nV
    reDraw
    PropertyChanged "LineAngle"
End Property
'LineWidth
Public Property Get LineWidth() As Integer
    LineWidth = lnWid
End Property
Public Property Let LineWidth(ByVal nV As Integer)
    lnWid = nV
    reDraw
    PropertyChanged "LineWidth"
End Property
'TransparentColor
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = TC
End Property
Public Property Let TransparentColor(ByVal nV As OLE_COLOR)
    TC = nV
    reDraw
    PropertyChanged "TransparentColor"
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
Public Property Get ShapeColor() As OLE_COLOR
    ShapeColor = BC1
End Property
Public Property Let ShapeColor(ByVal nV As OLE_COLOR)
    BC1 = nV
    reDraw
    PropertyChanged "ShapeColor"
End Property
'
Public Property Get FixedWidth() As Integer
    FixedWidth = mWidFixed
End Property
Public Property Let FixedWidth(ByVal nV As Integer)
    mWidFixed = nV
    reDraw
    PropertyChanged "FixedWidth"
End Property
'
Public Property Get FixedHeight() As Integer
    FixedHeight = mHeigFixed
End Property
Public Property Let FixedHeight(ByVal nV As Integer)
    mHeigFixed = nV
    reDraw
    PropertyChanged "FixedHeight"
End Property
'
Public Property Get Shape() As eSTShape
    Shape = mShape
End Property
Public Property Let Shape(ByVal nV As eSTShape)
    mShape = nV
    reDraw
    PropertyChanged "Shape"
End Property
'
Public Property Get HorizontalOrientation() As eSTHorOrj
    HorizontalOrientation = mHOrj
End Property
Public Property Let HorizontalOrientation(ByVal nV As eSTHorOrj)
    mHOrj = nV
    reDraw
    PropertyChanged "HorizontalOrientation"
End Property
'
Public Property Get VerticalOrientation() As eSTVertOrj
    VerticalOrientation = mVOrj
End Property
Public Property Let VerticalOrientation(ByVal nV As eSTVertOrj)
    mVOrj = nV
    reDraw
    PropertyChanged "VerticalOrientation"
End Property



Private Sub UserControl_Initialize()
    BC = vbBlack
    BC1 = &HFF8080
    
    mWidFixed = 80
    mHeigFixed = 20
    
    mVOrj = stVORBottom
    mHOrj = stHORight
    
    mShape = stFedCorner

    lnWid = 5
    lnAngle = 0
    
    reDraw
End Sub



Private Sub reDraw()
    UserControl.BackColor = BC
    UserControl.Cls
    UserControl.DrawWidth = 1
    If mShape = stFedCorner Or mShape = stFedRoundHor Or mShape = stFedRoundVert Then
        drawFederationShape
    ElseIf mShape = stRomulanCorner Or mShape = stRomulanLine Then
        drawRomulanShape
    ElseIf mShape = stCustomLine Then
        drawCustomLine
    End If
    
    UserControl.MaskColor = TC
    UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
End Sub


Private Sub drawFederationShape()
    Dim ucSH As Integer, ucSW As Integer
    Dim c1R As Integer, c2R As Integer
    
    UserControl.Cls
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    
    
    If mShape = stFedCorner Then
        c1R = mWidFixed / 3
        c2R = mHeigFixed
        'If c1R < c2R Then c1R = mWidFixed
        If c2R > c1R Then c2R = c1R
        
        'c2R = 20
        If mHOrj = stHOLeft And mVOrj = stVOTop Then
            UserControl.Circle (c1R, c1R), c1R, BC1, 3.14 / 2, 3.14
            UserControl.Line (c1R, 0)-(ucSW, mHeigFixed), BC1, BF
            UserControl.Line (0, c1R)-(0, ucSH), BC1
            UserControl.Line (0, ucSH - 1)-(mWidFixed, ucSH - 1), BC1
            UserControl.Line (mWidFixed, ucSH - 1)-(mWidFixed, mHeigFixed + c2R - 6), BC1
            UserControl.Circle (mWidFixed + c2R, mHeigFixed + c2R), c2R, BC1, 3.14 / 2, 3.14
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, 2, ucSH - 2, UserControl.Point(2, ucSH - 2), 1
        ElseIf mHOrj = stHOLeft And mVOrj = stVORBottom Then
            UserControl.Circle (c1R, ucSH - c1R), c1R, BC1, 3.14, 1.5 * 3.14
            UserControl.Line (c1R, ucSH)-(ucSW, ucSH - mHeigFixed), BC1, BF
            UserControl.Line (0, ucSH - c1R)-(0, 0), BC1
            UserControl.Line (0, 0)-(mWidFixed, 0), BC1
            UserControl.Line (mWidFixed, 0)-(mWidFixed, ucSH - mHeigFixed - c2R + 6), BC1
            UserControl.Circle (mWidFixed + c2R, ucSH - mHeigFixed - c2R), c2R, BC1, 3.14, 1.5 * 3.14
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, 2, 2, UserControl.Point(2, 2), 1
        ElseIf mHOrj = stHORight And mVOrj = stVOTop Then
            UserControl.Circle (ucSW - c1R, c1R), c1R, BC1, 0, 3.14 / 2
            UserControl.Line (ucSW - c1R, 0)-(0, mHeigFixed), BC1, BF
            UserControl.Line (ucSW - 1, c1R)-(ucSW - 1, ucSH), BC1
            UserControl.Line (ucSW, ucSH - 1)-(ucSW - mWidFixed, ucSH - 1), BC1
            UserControl.Line (ucSW - mWidFixed, ucSH - 1)-(ucSW - mWidFixed, mHeigFixed + c2R - 6), BC1
            UserControl.Circle (ucSW - mWidFixed - c2R, mHeigFixed + c2R), c2R, BC1, 0, 3.14 / 2
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW - 2, ucSH - 2, UserControl.Point(ucSW - 2, ucSH - 2), 1
        ElseIf mHOrj = stHORight And mVOrj = stVORBottom Then
            UserControl.Circle (ucSW - c1R, ucSH - c1R), c1R, BC1, 1.5 * 3.14, 0
            UserControl.Line (ucSW - c1R, ucSH)-(0, ucSH - mHeigFixed), BC1, BF
            UserControl.Line (ucSW - 1, ucSH - c1R)-(ucSW - 1, 0), BC1
            UserControl.Line (ucSW, 0)-(ucSW - mWidFixed, 0), BC1
            UserControl.Line (ucSW - mWidFixed, 0)-(ucSW - mWidFixed, ucSH - mHeigFixed - c2R + 6), BC1
            UserControl.Circle (ucSW - mWidFixed - c2R, ucSH - mHeigFixed - c2R), c2R, BC1, 1.5 * 3.14, 0
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW - 2, 2, UserControl.Point(ucSW - 2, 2), 1
        End If
    ElseIf mShape = stFedRoundHor Then
        If mHOrj = stHOLeft Then
            UserControl.Circle (ucSH / 2, ucSH / 2), ucSH / 2, BC1, 3.14 / 2, 1.5 * 3.14
            UserControl.Line (ucSH / 2, 0)-(ucSW, ucSH), BC1, BF
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, 2, ucSH / 2, UserControl.Point(2, ucSH / 2), 1
        Else
            UserControl.Circle (ucSW - ucSH / 2, ucSH / 2), ucSH / 2, BC1, 1.5 * 3.14, 3.14 / 2
            UserControl.Line (0, 0)-(ucSW - ucSH / 2, ucSH), BC1, BF
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW - 2, ucSH / 2, UserControl.Point(ucSW - 2, ucSH / 2), 1
        End If
    ElseIf mShape = stFedRoundVert Then
        If mVOrj = stVOTop Then
            UserControl.Circle (ucSW / 2, ucSW / 2), ucSW / 2, BC1, 0, 3.14
            UserControl.Line (0, ucSW / 2)-(ucSW, ucSH), BC1, BF
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW / 2, 2, UserControl.Point(ucSW / 2, 2), 1
        Else
            UserControl.Circle (ucSW / 2, ucSH - ucSW / 2), ucSW / 2, BC1, 3.14, 0
            UserControl.Line (0, 0)-(ucSW, ucSH - ucSW / 2), BC1, BF
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW / 2, ucSH - 2, UserControl.Point(ucSW / 2, ucSH - 2), 1
        End If
    End If
End Sub

Private Sub drawRomulanShape()
    Dim ucSH As Integer, ucSW As Integer
    
    UserControl.Cls
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    
    If mShape = stRomulanCorner Then
        If mHOrj = stHOLeft And mVOrj = stVOTop Then
            UserControl.Line (0, 0)-(lnWid, ucSH), BC1, BF
            UserControl.Line (lnWid, 0)-(mWidFixed, mHeigFixed / 2), BC1
            UserControl.Line (lnWid, lnWid)-(mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.Line (mWidFixed, mHeigFixed / 2)-(mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, mWidFixed - 1, mHeigFixed / 2 + 1, UserControl.Point(mWidFixed - 1, mHeigFixed / 2 + 1), 1
            '
            UserControl.Line (mWidFixed - lnWid, mHeigFixed / 2)-(mWidFixed - lnWid, mHeigFixed), BC1
            UserControl.Line (mWidFixed, mHeigFixed / 2)-(mWidFixed, mHeigFixed), BC1
            UserControl.Line (mWidFixed - lnWid, mHeigFixed)-(mWidFixed, mHeigFixed), BC1
            ExtFloodFill UserControl.hDC, mWidFixed - lnWid / 2, mHeigFixed - lnWid / 2, UserControl.Point(mWidFixed - lnWid / 2, mHeigFixed - lnWid / 2), 1
            '
            UserControl.Line (mWidFixed, mHeigFixed)-(ucSW, mHeigFixed + (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (mWidFixed - lnWid, mHeigFixed + lnWid / 2 + 1)-(mWidFixed - lnWid, mHeigFixed - 2), BC1
            UserControl.Line (mWidFixed - lnWid, mHeigFixed + lnWid / 2)-(ucSW - lnWid * 2, mHeigFixed + (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (ucSW, mHeigFixed + (ucSH - mHeigFixed) / 2)-(lnWid, ucSH), BC1
            UserControl.Line (ucSW - lnWid * 2, mHeigFixed + (ucSH - mHeigFixed) / 2)-(lnWid, ucSH - lnWid), BC1
            '
            ExtFloodFill UserControl.hDC, mWidFixed - lnWid + 1, mHeigFixed + lnWid / 2, UserControl.Point(mWidFixed - lnWid + 1, mHeigFixed + lnWid / 2), 1
        ElseIf mHOrj = stHOLeft And mVOrj = stVORBottom Then
            UserControl.Line (0, ucSH)-(lnWid, 0), BC1, BF
            UserControl.Line (lnWid, ucSH)-(mWidFixed, ucSH - mHeigFixed / 2), BC1
            UserControl.Line (lnWid, ucSH - lnWid)-(mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.Line (mWidFixed, ucSH - mHeigFixed / 2)-(mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, mWidFixed - 1, ucSH - mHeigFixed / 2 - 1, UserControl.Point(mWidFixed - 1, ucSH - mHeigFixed / 2 - 1), 1
            '
            UserControl.Line (mWidFixed - lnWid, ucSH - mHeigFixed / 2)-(mWidFixed - lnWid, ucSH - mHeigFixed), BC1
            UserControl.Line (mWidFixed, ucSH - mHeigFixed / 2)-(mWidFixed, ucSH - mHeigFixed), BC1
            UserControl.Line (mWidFixed - lnWid, ucSH - mHeigFixed)-(mWidFixed, ucSH - mHeigFixed), BC1
            ExtFloodFill UserControl.hDC, mWidFixed - lnWid / 2, ucSH - mHeigFixed + lnWid / 2, UserControl.Point(mWidFixed - lnWid / 2, ucSH - mHeigFixed + lnWid / 2), 1
            '
            UserControl.Line (mWidFixed, ucSH - mHeigFixed)-(ucSW, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (mWidFixed - lnWid, ucSH - mHeigFixed - lnWid / 2 - 1)-(mWidFixed - lnWid, ucSH - mHeigFixed + 2), BC1
            UserControl.Line (mWidFixed - lnWid, ucSH - mHeigFixed - lnWid / 2)-(ucSW - lnWid * 2, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (ucSW, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2)-(lnWid, 0), BC1
            UserControl.Line (ucSW - lnWid * 2, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2)-(lnWid, lnWid), BC1
            '
            ExtFloodFill UserControl.hDC, lnWid + 1, lnWid / 2, UserControl.Point(lnWid + 1, lnWid / 2), 1
        ElseIf mHOrj = stHORight And mVOrj = stVOTop Then
            UserControl.Line (ucSW, 0)-(ucSW - lnWid, ucSH), BC1, BF
            UserControl.Line (ucSW - lnWid, 0)-(ucSW - mWidFixed, mHeigFixed / 2), BC1
            UserControl.Line (ucSW - lnWid, lnWid)-(ucSW - mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.Line (ucSW - mWidFixed, mHeigFixed / 2)-(ucSW - mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW - mWidFixed + 1, mHeigFixed / 2 + 1, UserControl.Point(ucSW - mWidFixed + 1, mHeigFixed / 2 + 1), 1
            '
            UserControl.Line (ucSW - mWidFixed + lnWid, mHeigFixed / 2)-(ucSW - mWidFixed + lnWid, mHeigFixed), BC1
            UserControl.Line (ucSW - mWidFixed, mHeigFixed / 2)-(ucSW - mWidFixed, mHeigFixed), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, mHeigFixed)-(ucSW - mWidFixed, mHeigFixed), BC1
            ExtFloodFill UserControl.hDC, ucSW - mWidFixed + lnWid / 2, mHeigFixed - lnWid / 2, UserControl.Point(ucSW - mWidFixed + lnWid / 2, mHeigFixed - lnWid / 2), 1
            '
            UserControl.Line (ucSW - mWidFixed, mHeigFixed)-(0, mHeigFixed + (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, mHeigFixed + lnWid / 2 + 1)-(ucSW - mWidFixed + lnWid, mHeigFixed - 2), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, mHeigFixed + lnWid / 2)-(ucSW - ucSW + lnWid * 2, mHeigFixed + (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (0, mHeigFixed + (ucSH - mHeigFixed) / 2)-(ucSW - lnWid, ucSH), BC1
            UserControl.Line (lnWid * 2, mHeigFixed + (ucSH - mHeigFixed) / 2)-(ucSW - lnWid, ucSH - lnWid), BC1
            '
            ExtFloodFill UserControl.hDC, ucSW - mWidFixed + lnWid - 1, mHeigFixed + lnWid / 2, UserControl.Point(ucSW - mWidFixed + lnWid - 1, mHeigFixed + lnWid / 2), 1
        ElseIf mHOrj = stHORight And mVOrj = stVORBottom Then
            UserControl.Line (ucSW, ucSH)-(ucSW - lnWid, 0), BC1, BF
            UserControl.Line (ucSW - lnWid, ucSH)-(ucSW - mWidFixed, ucSH - mHeigFixed / 2), BC1
            UserControl.Line (ucSW - lnWid, ucSH - lnWid)-(ucSW - mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.Line (ucSW - mWidFixed, ucSH - mHeigFixed / 2)-(ucSW - mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, ucSW - mWidFixed + 1, ucSH - mHeigFixed / 2 - 1, UserControl.Point(ucSW - mWidFixed + 1, ucSH - mHeigFixed / 2 - 1), 1
            '
            UserControl.Line (ucSW - mWidFixed + lnWid, ucSH - mHeigFixed / 2)-(ucSW - mWidFixed + lnWid, ucSH - mHeigFixed), BC1
            UserControl.Line (ucSW - mWidFixed, ucSH - mHeigFixed / 2)-(ucSW - mWidFixed, ucSH - mHeigFixed), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, ucSH - mHeigFixed)-(ucSW - mWidFixed, ucSH - mHeigFixed), BC1
            ExtFloodFill UserControl.hDC, ucSW - mWidFixed + lnWid / 2, ucSH - mHeigFixed + lnWid / 2, UserControl.Point(ucSW - mWidFixed + lnWid / 2, ucSH - mHeigFixed + lnWid / 2), 1
            '
            UserControl.Line (ucSW - mWidFixed, ucSH - mHeigFixed)-(0, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, ucSH - mHeigFixed - lnWid / 2 - 1)-(ucSW - mWidFixed + lnWid, ucSH - mHeigFixed + 2), BC1
            UserControl.Line (ucSW - mWidFixed + lnWid, ucSH - mHeigFixed - lnWid / 2)-(ucSW - ucSW + lnWid * 2, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2), BC1
            UserControl.Line (0, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2)-(ucSW - lnWid, 0), BC1
            UserControl.Line (lnWid * 2, ucSH - mHeigFixed - (ucSH - mHeigFixed) / 2)-(ucSW - lnWid, lnWid), BC1
            '
            ExtFloodFill UserControl.hDC, ucSW - lnWid - 1, lnWid / 2, UserControl.Point(ucSW - lnWid - 1, lnWid / 2), 1
        End If
    ElseIf mShape = stRomulanLine Then
        If mHOrj = stHOLeft Then
            mWidFixed = ucSW
            UserControl.Line (0, 0)-(lnWid, ucSH), BC1, BF
            UserControl.Line (lnWid, 0)-(mWidFixed, mHeigFixed / 2), BC1
            UserControl.Line (lnWid, lnWid)-(mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.Line (mWidFixed, mHeigFixed / 2)-(mWidFixed, mHeigFixed / 2 + lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, mWidFixed - 1, mHeigFixed / 2 + 1, UserControl.Point(mWidFixed - 1, mHeigFixed / 2 + 1), 1
            '
            UserControl.Line (0, ucSH)-(lnWid, 0), BC1, BF
            UserControl.Line (lnWid, ucSH)-(mWidFixed, ucSH - mHeigFixed / 2), BC1
            UserControl.Line (lnWid, ucSH - lnWid)-(mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.Line (mWidFixed, ucSH - mHeigFixed / 2)-(mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            ExtFloodFill UserControl.hDC, mWidFixed - 1, ucSH - mHeigFixed / 2 - 1, UserControl.Point(mWidFixed - 1, ucSH - mHeigFixed / 2 - 1), 1
            '
            UserControl.Line (mWidFixed - lnWid, ucSH - mHeigFixed / 2 - lnWid)-(mWidFixed - 1, mHeigFixed / 2 + lnWid), BC1, BF
        ElseIf mHOrj = stHORight Then
            mWidFixed = ucSW
            UserControl.Line (mWidFixed, 0)-(mWidFixed - lnWid, ucSH), BC1, BF
            UserControl.Line (mWidFixed - lnWid, 0)-(0, mHeigFixed / 2), BC1
            UserControl.Line (mWidFixed - lnWid, lnWid)-(0, mHeigFixed / 2 + lnWid), BC1
            UserControl.Line (0, mHeigFixed / 2)-(0, mHeigFixed / 2 + lnWid), BC1
            UserControl.FillColor = BC1
            ExtFloodFill UserControl.hDC, 1, mHeigFixed / 2 + 1, UserControl.Point(1, mHeigFixed / 2 + 1), 1
            '
            UserControl.Line (mWidFixed, ucSH)-(mWidFixed - lnWid, 0), BC1, BF
            UserControl.Line (mWidFixed - lnWid, ucSH)-(0, ucSH - mHeigFixed / 2), BC1
            UserControl.Line (mWidFixed - lnWid, ucSH - lnWid)-(mWidFixed - mWidFixed, ucSH - mHeigFixed / 2 - lnWid), BC1
            UserControl.Line (0, ucSH - mHeigFixed / 2)-(0, ucSH - mHeigFixed / 2 - lnWid), BC1
            ExtFloodFill UserControl.hDC, 1, ucSH - mHeigFixed / 2 - 1, UserControl.Point(1, ucSH - mHeigFixed / 2 - 1), 1
            '
            UserControl.Line (lnWid, ucSH - mHeigFixed / 2 - lnWid)-(0, mHeigFixed / 2 + lnWid), BC1, BF
        End If
    End If
End Sub

Private Sub drawCustomLine()
    Dim ucSH As Integer, ucSW As Integer
    Dim mRad As Single
    UserControl.Cls
    ucSH = UserControl.ScaleHeight
    ucSW = UserControl.ScaleWidth
    
    'find angle in radians
    mRad = lnAngle * 3.14 / 180
    '
    UserControl.DrawWidth = lnWid
    UserControl.Line (ucSW / 2, ucSH / 2)-(ucSW / 2 + ucSW * Cos(mRad), ucSH / 2 - ucSH * Sin(mRad)), BC1
    UserControl.Line (ucSW / 2, ucSH / 2)-(ucSW / 2 - ucSW * Cos(mRad), ucSH / 2 + ucSH * Sin(mRad)), BC1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BC = PropBag.ReadProperty("BackColor", vbBlack)
    TC = PropBag.ReadProperty("TransparentColor", vbBlack)
    BC1 = PropBag.ReadProperty("ShapeColor", &HFF8080)
    mWidFixed = PropBag.ReadProperty("FixedWidth", 80)
    mHeigFixed = PropBag.ReadProperty("FixedHeight", 20)
    mShape = PropBag.ReadProperty("Shape", 0)
    mVOrj = PropBag.ReadProperty("VerticalOrientation", 0)
    mHOrj = PropBag.ReadProperty("HorizontalOrientation", 0)
    lnWid = PropBag.ReadProperty("LineWidth", 5)
    lnAngle = PropBag.ReadProperty("LineAngle", 0)
    
    reDraw
End Sub

Private Sub UserControl_Resize()
    reDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BC, vbBlack
    PropBag.WriteProperty "TransparentColor", TC, vbBlack
    PropBag.WriteProperty "ShapeColor", BC1, &HFF8080
    PropBag.WriteProperty "FixedWidth", mWidFixed, 80
    PropBag.WriteProperty "FixedHeight", mHeigFixed, 20
    PropBag.WriteProperty "Shape", mShape, 0
    PropBag.WriteProperty "VerticalOrientation", mVOrj, 0
    PropBag.WriteProperty "HorizontalOrientation", mHOrj, 0
    PropBag.WriteProperty "LineWidth", lnWid, 5
    PropBag.WriteProperty "LineAngle", lnAngle, 0
End Sub

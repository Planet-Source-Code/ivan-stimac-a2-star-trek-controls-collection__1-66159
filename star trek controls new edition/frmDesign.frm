VERSION 5.00
Begin VB.Form frmDesign 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   Picture         =   "frmDesign.frx":0000
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   8
      Top             =   1140
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   15699535
      BackColorMiddle2=   15699535
      BackColorMiddle1Hover=   16631721
      BackColorMiddle2Hover=   16631721
      BackColorMiddle1Down=   16631721
      BackColorMiddle2down=   16631721
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   15699535
      BackColorLeftHover=   15699535
      BackColorLeftDown=   15699535
      BackColorLeftDisabled=   8421504
      BackColorRight  =   15699535
      BackColorRightHover=   16631721
      BackColorRightDown=   16631721
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "Button styles"
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1296
      ShapeColor      =   16413972
      FixedWidth      =   40
      FixedHeight     =   10
      Shape           =   2
      VerticalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape2 
      Height          =   345
      Left            =   10140
      TabIndex        =   3
      Top             =   2625
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   609
      ShapeColor      =   16435886
      FixedWidth      =   40
      FixedHeight     =   10
      Shape           =   1
      HorizontalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   1620
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   15699535
      BackColorMiddle2=   15699535
      BackColorMiddle1Hover=   16631721
      BackColorMiddle2Hover=   16631721
      BackColorMiddle1Down=   16631721
      BackColorMiddle2down=   16631721
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   15699535
      BackColorLeftHover=   15699535
      BackColorLeftDown=   15699535
      BackColorLeftDisabled=   8421504
      BackColorRight  =   15699535
      BackColorRightHover=   16631721
      BackColorRightDown=   16631721
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "Scroll styles"
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   10
      Top             =   2100
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   15699535
      BackColorMiddle2=   15699535
      BackColorMiddle1Hover=   16631721
      BackColorMiddle2Hover=   16631721
      BackColorMiddle1Down=   16631721
      BackColorMiddle2down=   16631721
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   15699535
      BackColorLeftHover=   15699535
      BackColorLeftDown=   15699535
      BackColorLeftDisabled=   8421504
      BackColorRight  =   15699535
      BackColorRightHover=   16631721
      BackColorRightDown=   16631721
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "ProgressBar"
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   3
      Left            =   4380
      TabIndex        =   11
      Top             =   1140
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   15699535
      BackColorMiddle2=   15699535
      BackColorMiddle1Hover=   16631721
      BackColorMiddle2Hover=   16631721
      BackColorMiddle1Down=   16631721
      BackColorMiddle2down=   16631721
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   15699535
      BackColorLeftHover=   15699535
      BackColorLeftDown=   15699535
      BackColorLeftDisabled=   8421504
      BackColorRight  =   15699535
      BackColorRightHover=   16631721
      BackColorRightDown=   16631721
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "Shape styles"
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   4
      Left            =   4380
      TabIndex        =   12
      Top             =   1620
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   15699535
      BackColorMiddle2=   15699535
      BackColorMiddle1Hover=   16631721
      BackColorMiddle2Hover=   16631721
      BackColorMiddle1Down=   16631721
      BackColorMiddle2down=   16631721
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   15699535
      BackColorLeftHover=   15699535
      BackColorLeftDown=   15699535
      BackColorLeftDisabled=   8421504
      BackColorRight  =   15699535
      BackColorRightHover=   16631721
      BackColorRightDown=   16631721
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "About"
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   315
      Index           =   5
      Left            =   4380
      TabIndex        =   13
      Top             =   2100
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FederationBold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   4194368
      TransparentColor=   4194368
      BackColorMiddle1=   16576
      BackColorMiddle2=   16576
      BackColorMiddle1Hover=   33023
      BackColorMiddle2Hover=   33023
      BackColorMiddle1Down=   16576
      BackColorMiddle2down=   16576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16576
      BackColorLeftHover=   16576
      BackColorLeftDown=   16576
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16576
      BackColorRightHover=   33023
      BackColorRightDown=   16576
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      Caption         =   "Exit"
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape3 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   2820
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1296
      ShapeColor      =   14376982
      FixedWidth      =   40
      FixedHeight     =   10
      Shape           =   2
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape4 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   7140
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      ShapeColor      =   14376982
      FixedWidth      =   40
      Shape           =   2
      VerticalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape5 
      Height          =   300
      Left            =   10140
      TabIndex        =   26
      Top             =   7560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      ShapeColor      =   33023
      FixedWidth      =   40
      FixedHeight     =   10
      Shape           =   1
      HorizontalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape6 
      Height          =   1395
      Left            =   960
      TabIndex        =   27
      Top             =   3360
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2461
      ShapeColor      =   16435886
      FixedWidth      =   10
      FixedHeight     =   5
      Shape           =   2
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape7 
      Height          =   1395
      Left            =   960
      TabIndex        =   28
      Top             =   5760
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2461
      ShapeColor      =   16435886
      FixedWidth      =   10
      FixedHeight     =   5
      Shape           =   2
      VerticalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape8 
      Height          =   1395
      Left            =   3000
      TabIndex        =   29
      Top             =   3360
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2461
      ShapeColor      =   16435886
      FixedWidth      =   10
      FixedHeight     =   5
      Shape           =   2
      HorizontalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape9 
      Height          =   1395
      Left            =   3000
      TabIndex        =   30
      Top             =   5760
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   2461
      ShapeColor      =   16435886
      FixedWidth      =   10
      FixedHeight     =   5
      Shape           =   2
      VerticalOrientation=   1
      HorizontalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape10 
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Top             =   3360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      ShapeColor      =   16435886
      FixedWidth      =   10
      Shape           =   2
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape11 
      Height          =   495
      Left            =   3720
      TabIndex        =   42
      Top             =   6660
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      ShapeColor      =   16435886
      FixedWidth      =   10
      Shape           =   2
      VerticalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape12 
      Height          =   495
      Left            =   7350
      TabIndex        =   44
      Top             =   3360
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
      ShapeColor      =   16435886
      FixedWidth      =   10
      Shape           =   2
      HorizontalOrientation=   1
   End
   Begin StarTrekControlsNE.StarTrekShape StarTrekShape13 
      Height          =   495
      Left            =   9690
      TabIndex        =   47
      Top             =   6660
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      ShapeColor      =   16435886
      FixedWidth      =   10
      Shape           =   2
      VerticalOrientation=   1
      HorizontalOrientation=   1
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "contact: flashboy01@gmail.com"
      BeginProperty Font 
         Name            =   "Trek TNG Monitors"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Left            =   4140
      TabIndex        =   55
      Top             =   5280
      Width           =   5475
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   1140
      Picture         =   "frmDesign.frx":6EBA
      Top             =   4080
      Width           =   1950
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank:    Captain"
      BeginProperty Font 
         Name            =   "Trek TNG Monitors"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Left            =   4140
      TabIndex        =   54
      Top             =   4740
      Width           =   5475
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:   Ivan Stimac"
      BeginProperty Font 
         Name            =   "Trek TNG Monitors"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Left            =   4140
      TabIndex        =   53
      Top             =   4200
      Width           =   5475
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FACAAE&
      Height          =   300
      Left            =   8100
      TabIndex        =   52
      Top             =   6855
      Width           =   1545
   End
   Begin VB.Label Label31 
      BackColor       =   &H00EF8E4F&
      Height          =   300
      Index           =   1
      Left            =   5700
      TabIndex        =   51
      Top             =   6855
      Width           =   2355
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FACAAE&
      Height          =   300
      Left            =   4140
      TabIndex        =   50
      Top             =   6855
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4140
      Picture         =   "frmDesign.frx":A867
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT CREW MEMBER"
      BeginProperty Font 
         Name            =   "Trek TNG Monitors"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   4680
      TabIndex        =   49
      Top             =   3300
      Width           =   3255
   End
   Begin VB.Label Label30 
      BackColor       =   &H00AB6016&
      Height          =   1110
      Left            =   9900
      TabIndex        =   48
      Top             =   3900
      Width           =   165
   End
   Begin VB.Label Label29 
      BackColor       =   &H00AB6016&
      Height          =   1110
      Left            =   9900
      TabIndex        =   46
      Top             =   5490
      Width           =   165
   End
   Begin VB.Label Label28 
      BackColor       =   &H000080FF&
      Height          =   390
      Left            =   9900
      TabIndex        =   45
      Top             =   5055
      Width           =   105
   End
   Begin VB.Label Label27 
      BackColor       =   &H00AB6016&
      Height          =   1110
      Left            =   3720
      TabIndex        =   43
      Top             =   3900
      Width           =   165
   End
   Begin VB.Label Label26 
      BackColor       =   &H00AB6016&
      Height          =   1110
      Left            =   3720
      TabIndex        =   41
      Top             =   5490
      Width           =   165
   End
   Begin VB.Label Label25 
      BackColor       =   &H000080FF&
      Height          =   390
      Left            =   3780
      TabIndex        =   40
      Top             =   5055
      Width           =   105
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FACAAE&
      Height          =   225
      Left            =   3270
      TabIndex        =   38
      Top             =   5160
      Width           =   30
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FACAAE&
      Height          =   225
      Left            =   3180
      TabIndex        =   37
      Top             =   5160
      Width           =   30
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FACAAE&
      Height          =   225
      Left            =   1080
      TabIndex        =   36
      Top             =   5160
      Width           =   30
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FACAAE&
      Height          =   225
      Left            =   990
      TabIndex        =   35
      Top             =   5160
      Width           =   30
   End
   Begin VB.Label Label20 
      BackColor       =   &H00AB6016&
      Height          =   225
      Left            =   3150
      TabIndex        =   34
      Top             =   5490
      Width           =   165
   End
   Begin VB.Label Label19 
      BackColor       =   &H00AB6016&
      Height          =   225
      Left            =   3150
      TabIndex        =   33
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label Label18 
      BackColor       =   &H00AB6016&
      Height          =   225
      Left            =   960
      TabIndex        =   32
      Top             =   5490
      Width           =   165
   End
   Begin VB.Label Label17 
      BackColor       =   &H00AB6016&
      Height          =   225
      Left            =   960
      TabIndex        =   31
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FACAAE&
      Height          =   300
      Left            =   6120
      TabIndex        =   25
      Top             =   7560
      Width           =   3975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00AB6016&
      Height          =   300
      Left            =   2220
      TabIndex        =   24
      Top             =   7560
      Width           =   3855
   End
   Begin VB.Label Label14 
      BackColor       =   &H000080FF&
      Height          =   450
      Left            =   120
      TabIndex        =   22
      Top             =   6630
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080C0FF&
      Height          =   1110
      Left            =   120
      TabIndex        =   21
      Top             =   5475
      Width           =   615
   End
   Begin VB.Label Label12 
      BackColor       =   &H00AB6016&
      Height          =   1830
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00DB6016&
      Height          =   150
      Left            =   1200
      TabIndex        =   19
      Top             =   2820
      Width           =   2355
   End
   Begin VB.Label Label10 
      BackColor       =   &H000040C0&
      Height          =   150
      Left            =   3600
      TabIndex        =   18
      Top             =   2820
      Width           =   315
   End
   Begin VB.Label Label9 
      BackColor       =   &H00DB6016&
      Height          =   105
      Left            =   3960
      TabIndex        =   17
      Top             =   2835
      Width           =   2355
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080C0FF&
      Height          =   150
      Left            =   6360
      TabIndex        =   16
      Top             =   2820
      Width           =   2355
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Star Trek Control Set: New Edition"
      BeginProperty Font 
         Name            =   "Trek TNG Monitors"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDC7A9&
      Height          =   615
      Left            =   1260
      TabIndex        =   14
      Top             =   300
      Width           =   9075
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FACAAE&
      Height          =   1830
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00DB6016&
      Height          =   150
      Left            =   6360
      TabIndex        =   6
      Top             =   2625
      Width           =   2355
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DB6016&
      Height          =   105
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FACAAE&
      Height          =   150
      Left            =   3600
      TabIndex        =   4
      Top             =   2625
      Width           =   315
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FA7514&
      Height          =   345
      Left            =   8760
      TabIndex        =   2
      Top             =   2625
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DB6016&
      Height          =   150
      Left            =   1200
      TabIndex        =   1
      Top             =   2625
      Width           =   2355
   End
End
Attribute VB_Name = "frmDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Star Trek Controls Set: New Edition
'   by Ivan Stimac, ivan.stimac@po.htnet.hr
'
'   I think this is largest free star trek
'   controls collection. But you can use them
'   in other project that is not trekkie.
'
'   This is excelent showcase how to create activeX controls
'   without any other control (except timer control)
'
'   There is much styles
'   and I have spent much time to develop them
'   SO IF YOU LIKE IT PLEASE VOTE
'
'
'   NOTE:
'   ------
'   I have use clsGradient from Kath-Rock Gradient Demo source
'   Thanks to Kath-Rock
'
Private Sub StarTrekButton1_Click(Index As Integer)
    Select Case Index
        Case 0
            frmTest.Show vbModal
        Case 1
            frmScroll.Show
        Case 2
            frmProgBars.Show vbModal
        Case 3
            FrmShapes.Show vbModal
        Case 4
            MsgBox "Star Trek Controls Set: New Edition" & vbCrLf & _
                    "Author: Ivan Stimac, ivan.stimac@po.htnet.hr" & vbCrLf & _
                    "PLEASE VOTE"
                    
        Case 5
            End
    End Select
End Sub


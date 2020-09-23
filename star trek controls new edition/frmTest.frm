VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "6"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   12570
   StartUpPosition =   2  'CenterScreen
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton5 
      Height          =   435
      Left            =   8100
      TabIndex        =   4
      Top             =   4320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   6
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton3 
      Height          =   435
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle1Hover=   16761024
      BackColorMiddle2Hover=   16761024
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorHover  =   8421504
      ForeColorDown   =   12632256
      ForeColorDisabled=   4210752
      SpacingLeftHover=   1
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton1 
      Height          =   435
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16777215
      BackColorMiddle2=   8388608
      BackColorMiddle1Hover=   14737632
      BackColorMiddle2Hover=   8388608
      BackColorMiddle1Down=   12632256
      BackColorMiddle2down=   8388608
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16711680
      BackColorLeftHover=   16711680
      BackColorLeftDown=   16711680
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      SpacingLeftDisabled=   1
      SpacingRightDisabled=   1
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel1 
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   8454016
      BackColorMiddle2=   32768
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   32768
      BackColorLeftDisabled=   8421504
      BackColorRight  =   32768
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Align           =   1
      Caption         =   "Fedaration button styles"
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel2 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   4275
      _ExtentX        =   6482
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   12632319
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   6
      Align           =   1
      GradientAngle   =   0
      Caption         =   "Romulan button styles"
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton2 
      Height          =   435
      Left            =   6360
      TabIndex        =   5
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   8388736
      BackColorMiddle1Hover=   16761024
      BackColorMiddle1Down=   16761087
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftHover=   16761024
      BackColorLeftDisabled=   8421504
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton4 
      Height          =   435
      Left            =   9360
      TabIndex        =   6
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16777215
      BackColorMiddle2=   16744576
      BackColorMiddle1Hover=   14737632
      BackColorMiddle2Hover=   16744576
      BackColorMiddle1Down=   12632256
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftHover=   16761024
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Align           =   2
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton6 
      Height          =   435
      Left            =   360
      TabIndex        =   7
      Top             =   1500
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle1Hover=   16761024
      BackColorMiddle2Hover=   16761024
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftHover=   16761024
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorHover  =   8421504
      ForeColorDown   =   12632256
      ForeColorDisabled=   4210752
      Style           =   1
      Align           =   2
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton7 
      Height          =   435
      Left            =   3360
      TabIndex        =   8
      Top             =   1500
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle2Hover=   16744576
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   1
      Align           =   2
      SpacingLeftHover=   1
      SpacingLeftDown =   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton8 
      Height          =   435
      Left            =   360
      TabIndex        =   9
      Top             =   2100
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle1Hover=   16761024
      BackColorMiddle2Hover=   16761024
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorHover  =   8421504
      ForeColorDown   =   12632256
      ForeColorDisabled=   4210752
      Style           =   2
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton9 
      Height          =   435
      Left            =   3360
      TabIndex        =   10
      Top             =   2100
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle2Hover=   16744576
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftDisabled=   8421504
      BackColorRight  =   8421504
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   2
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRightHover=   1
      SpacingRightDown=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton10 
      Height          =   435
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle2=   16744576
      BackColorMiddle1Hover=   16761024
      BackColorMiddle2Hover=   16761024
      BackColorMiddle2down=   16744576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744576
      BackColorRightHover=   16761024
      BackColorRightDown=   16744576
      BackColorRightDisabled=   8421504
      ForeColorHover  =   8421504
      ForeColorDown   =   12632256
      ForeColorDisabled=   4210752
      Style           =   3
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton11 
      Height          =   435
      Left            =   420
      TabIndex        =   12
      Top             =   4320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   6
      GradientAngle   =   0
      SpacingLeftHover=   1
      SpacingLeftDown =   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton12 
      Height          =   435
      Left            =   4320
      TabIndex        =   13
      Top             =   4320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   6
      Align           =   2
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRightHover=   1
      SpacingRightDown=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton13 
      Height          =   435
      Left            =   8100
      TabIndex        =   14
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   5
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton14 
      Height          =   435
      Left            =   420
      TabIndex        =   15
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   5
      GradientAngle   =   0
      SpacingLeftHover=   1
      SpacingLeftDown =   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton15 
      Height          =   435
      Left            =   4320
      TabIndex        =   16
      Top             =   4920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   16733618
      BackColorMiddle2down=   16733618
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   5
      Align           =   2
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRightHover=   1
      SpacingRightDown=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton17 
      Height          =   435
      Left            =   420
      TabIndex        =   17
      Top             =   5460
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   53259
      BackColorMiddle2down=   12648384
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   53259
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   12648384
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   4
      GradientAngle   =   0
      SpacingLeftHover=   1
      SpacingLeftDown =   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton16 
      Height          =   435
      Left            =   8100
      TabIndex        =   18
      Top             =   5460
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   53259
      BackColorMiddle2down=   12648384
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   53259
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   12648384
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   4
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton18 
      Height          =   435
      Left            =   4320
      TabIndex        =   19
      Top             =   5460
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16733618
      BackColorMiddle2=   16744140
      BackColorMiddle1Hover=   4454843
      BackColorMiddle2Hover=   12648384
      BackColorMiddle1Down=   53259
      BackColorMiddle2down=   12648384
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   53259
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   12648384
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   4
      Align           =   2
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRightHover=   1
      SpacingRightDown=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel3 
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   6300
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   192
      BackColorMiddle2=   8438015
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   7
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   ""
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel4 
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   6720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16576
      BackColorMiddle2=   192
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   10
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   "Klingon button styles"
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton21 
      Height          =   435
      Left            =   240
      TabIndex        =   22
      Top             =   7200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   255
      BackColorMiddle2=   64
      BackColorMiddle1Hover=   8421631
      BackColorMiddle2Hover=   64
      BackColorMiddle1Down=   192
      BackColorMiddle2down=   64
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   7
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton22 
      Height          =   435
      Left            =   240
      TabIndex        =   23
      Top             =   7740
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   33023
      BackColorMiddle2=   64
      BackColorMiddle1Hover=   8438015
      BackColorMiddle2Hover=   64
      BackColorMiddle1Down=   12640511
      BackColorMiddle2down=   64
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   10
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel5 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   8160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   8438015
      BackColorMiddle2=   192
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   7
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   ""
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel6 
      Height          =   495
      Left            =   8220
      TabIndex        =   25
      Top             =   6720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   8438015
      BackColorMiddle2=   192
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   10
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   ""
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel7 
      Height          =   495
      Left            =   3960
      TabIndex        =   26
      Top             =   8580
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   192
      BackColorMiddle2=   16576
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   10
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   ""
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekLabel StarTrekLabel8 
      Height          =   495
      Left            =   7800
      TabIndex        =   27
      Top             =   8160
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   16576
      BackColorMiddle2=   192
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   4210752
      BackColorLeft   =   49152
      BackColorLeftDisabled=   8421504
      BackColorRightDisabled=   4210752
      ForeColorDisabled=   4210752
      Style           =   7
      Align           =   1
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingRight    =   0
      Caption         =   ""
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton23 
      Height          =   435
      Left            =   4260
      TabIndex        =   28
      Top             =   7020
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   255
      BackColorMiddle2=   64
      BackColorMiddle1Hover=   8421631
      BackColorMiddle2Hover=   64
      BackColorMiddle1Down=   192
      BackColorMiddle2down=   64
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   11
      GradientAngle   =   0
      SpacingLeft     =   4
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton24 
      Height          =   435
      Left            =   4500
      TabIndex        =   29
      Top             =   7500
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   128
      BackColorMiddle2=   255
      BackColorMiddle1Hover=   128
      BackColorMiddle2Hover=   8421631
      BackColorMiddle1Down=   64
      BackColorMiddle2down=   192
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   12
      Align           =   2
      GradientAngle   =   0
      SpacingLeft     =   0
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   4
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
   Begin StarTrekControlsNE.StarTrekButton StarTrekButton25 
      Height          =   435
      Left            =   4200
      TabIndex        =   30
      Top             =   7980
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorMiddle1=   255
      BackColorMiddle2=   64
      BackColorMiddle1Hover=   8421631
      BackColorMiddle2Hover=   64
      BackColorMiddle1Down=   192
      BackColorMiddle2down=   64
      BackColorMiddle1Disabled=   8421504
      BackColorMiddle2Disabled=   8421504
      BackColorLeft   =   16733618
      BackColorLeftHover=   4454843
      BackColorLeftDown=   16733618
      BackColorLeftDisabled=   8421504
      BackColorRight  =   16744140
      BackColorRightHover=   12648384
      BackColorRightDown=   16733618
      BackColorRightDisabled=   8421504
      ForeColorDisabled=   4210752
      Style           =   11
      GradientAngle   =   0
      SpacingLeft     =   4
      SpacingLeftHover=   0
      SpacingLeftDown =   0
      SpacingLeftDisabled=   0
      SpacingRight    =   0
      SpacingRightHover=   0
      SpacingRightDown=   0
      SpacingRightDisabled=   0
      AutoFillSides   =   -1  'True
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============== properties: ========================================================
'
'---- label and button ------------------
'
'   > align: caption align (left, center, right)
'   > auto fill sides: if true left side color will be same as first pixel on
'                      middle and right side color will be same as last
'                      pixel color. This is usefull if you use gradient fill
'   > BackColor: control back color
'   > BackColorLeft(*): back color of button left side
'   > BackColorMiddle1(*): gradinet start color
'   > BackColorMiddle2(*): gradinet end color
'   > BackColorRight(*): back color of button right side
'   > caption: you don't know?
'   > enabled:     = || =
'   > font, foreColor(*): = || =
'
'   > SpacingLeft(*): space between left side and middle od control
'   > SpacingRight(*): space between right side and middle od control
'   > style: it's style
'



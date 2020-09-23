VERSION 5.00
Begin VB.Form frmScroll 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin StarTrekControlsNE.StarTrekScroll StarTrekScroll1 
      Height          =   375
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      ChangeStep      =   50
   End
   Begin StarTrekControlsNE.StarTrekScroll StarTrekScroll4 
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5741
      ButtBackColor   =   16711680
      Style           =   1
      ChangeStep      =   50
   End
   Begin VB.Label lblVal 
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1140
      Width           =   3015
   End
End
Attribute VB_Name = "frmScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StarTrekScroll1_Change()
    lblVal.Caption = Me.StarTrekScroll1.Value
End Sub

Private Sub StarTrekScroll4_Change()
    lblVal.Caption = Me.StarTrekScroll4.Value
End Sub

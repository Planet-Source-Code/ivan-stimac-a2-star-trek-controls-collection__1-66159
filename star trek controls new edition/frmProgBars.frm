VERSION 5.00
Begin VB.Form frmProgBars 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProgBars.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4980
      Top             =   1500
   End
   Begin StarTrekControlsNE.StarTrekProgress StarTrekProgress1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      BarColor1       =   8438015
      BarColor2       =   192
      BarBackColor    =   12632256
      Style           =   2
      LineSpacing     =   5
      Value           =   5
   End
   Begin StarTrekControlsNE.StarTrekProgress StarTrekProgress2 
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BarColor2       =   192
      LineSpacing     =   5
      Value           =   5
   End
   Begin StarTrekControlsNE.StarTrekProgress StarTrekProgress3 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   556
      BarColor2       =   12582912
      BarBackColor    =   12632256
      Style           =   1
      LineSpacing     =   5
      Value           =   50
   End
   Begin StarTrekControlsNE.StarTrekProgress StarTrekProgress4 
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      BarColor2       =   192
      SideSpacingRight=   5
      LineSpacing     =   5
      Value           =   5
   End
   Begin VB.Label lblVal 
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
   End
End
Attribute VB_Name = "frmProgBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rndV As Integer

Private Sub StarTrekProgress1_Change()
    lblVal.Caption = Me.StarTrekProgress1.Value
End Sub

Private Sub Timer1_Timer()
    Randomize Timer
    rndV = Int(Rnd * 100)
    If rndV > 100 Then rndV = 100
'    rndV = rndV + 1
    StarTrekProgress2.Value = rndV
    StarTrekProgress3.Value = rndV
    StarTrekProgress4.Value = rndV
    StarTrekProgress1.Value = rndV
    
End Sub

VERSION 5.00
Begin VB.Form frmTetris4 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   30000
   ClientTop       =   0
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrZ 
      Interval        =   100
      Left            =   0
      Top             =   2040
   End
   Begin VB.Image imgZ 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgSW1 
      Height          =   1410
      Left            =   3120
      Picture         =   "frmTetris4.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image imgSW2 
      Height          =   1410
      Left            =   3000
      Picture         =   "frmTetris4.frx":6B82
      Top             =   1200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW3 
      Height          =   1410
      Left            =   3360
      Picture         =   "frmTetris4.frx":D87C
      Top             =   1560
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW4 
      Height          =   1410
      Left            =   3360
      Picture         =   "frmTetris4.frx":14576
      Top             =   1680
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW5 
      Height          =   1410
      Left            =   3480
      Picture         =   "frmTetris4.frx":1B270
      Top             =   1560
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW6 
      Height          =   1410
      Left            =   3480
      Picture         =   "frmTetris4.frx":21F6A
      Top             =   1200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW7 
      Height          =   1410
      Left            =   3600
      Picture         =   "frmTetris4.frx":28C64
      Top             =   1080
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgSW8 
      Height          =   1410
      Left            =   3240
      Picture         =   "frmTetris4.frx":2F95E
      Top             =   840
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgGun1 
      Height          =   1410
      Left            =   3000
      Picture         =   "frmTetris4.frx":36658
      Top             =   1200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image img1Lin1 
      Height          =   1410
      Left            =   2280
      Picture         =   "frmTetris4.frx":3D352
      Top             =   600
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image img2Lin1 
      Height          =   1410
      Left            =   3000
      Picture         =   "frmTetris4.frx":43BE4
      Top             =   1800
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image img2Lin2 
      Height          =   1440
      Left            =   2400
      Picture         =   "frmTetris4.frx":4710E
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgGun2 
      Height          =   1410
      Left            =   2880
      Picture         =   "frmTetris4.frx":4C3D8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Image imgBG 
      Height          =   3000
      Left            =   0
      Picture         =   "frmTetris4.frx":530D2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmTetris4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrZ_Timer()
WHICHone = WHICHone + 1
If HOWmany = 1 Then
    If WHICHone = 1 Then imgZ.Picture = img1Lin1.Picture
    If WHICHone = 10 Then Unload frmTetris4
End If
If HOWmany = 2 Then
    If WHICHone = 1 Then imgZ.Picture = img2Lin1.Picture
    If WHICHone = 4 Then imgZ.Picture = img2Lin2.Picture
    If WHICHone = 12 Then Unload frmTetris4
End If
If HOWmany = 3 Then
    If WHICHone = 1 Then imgZ.Picture = imgGun1.Picture
    If WHICHone = 4 Then imgZ.Picture = imgGun2.Picture
    If WHICHone = 7 Then imgZ.Picture = imgGun1.Picture
    If WHICHone = 10 Then imgZ.Picture = imgGun2.Picture
    If WHICHone = 13 Then imgZ.Picture = imgGun1.Picture
    If WHICHone = 16 Then imgZ.Picture = imgGun2.Picture
    If WHICHone = 19 Then imgZ.Picture = imgGun1.Picture
    If WHICHone = 22 Then imgZ.Picture = imgGun2.Picture
    If WHICHone = 25 Then Unload frmTetris4
End If
If HOWmany = 4 Then
    If WHICHone = 1 Then imgZ.Picture = imgSW1.Picture
    If WHICHone = 4 Then imgZ.Picture = imgSW2.Picture
    If WHICHone = 7 Then imgZ.Picture = imgSW3.Picture
    If WHICHone = 10 Then imgZ.Picture = imgSW4.Picture
    If WHICHone = 13 Then imgZ.Picture = imgSW5.Picture
    If WHICHone = 16 Then imgZ.Picture = imgSW6.Picture
    If WHICHone = 19 Then imgZ.Picture = imgSW7.Picture
    If WHICHone = 22 Then imgZ.Picture = imgSW8.Picture
    If WHICHone = 25 Then Unload frmTetris4
End If
frmTetris4.Width = imgZ.Width
frmTetris4.Height = imgZ.Height
imgBG.Width = frmTetris4.Width
imgBG.Height = frmTetris4.Height
frmTetris.SetFocus
End Sub

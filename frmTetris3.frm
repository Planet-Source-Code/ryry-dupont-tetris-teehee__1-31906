VERSION 5.00
Begin VB.Form frmTetris3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You've Found a Secret"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   2520
   End
   Begin VB.Label lblProd 
      AutoSize        =   -1  'True
      Caption         =   "A StainMaster Production"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label lblA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmTetris3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CAPPY(20) As String, tmpCAP As Integer
Private Sub Command1_Click()
xMIKI = 0
Unload frmTetris3
End Sub
Private Sub Form_Load()
CAPPY(1) = "Don't you just love secrets?"
CAPPY(2) = "So do I"
CAPPY(3) = "Let's get them credits rollin'"
CAPPY(4) = "Programmed by:" & vbCrLf & "Ryan DuPont"
CAPPY(5) = "Music by:" & vbCrLf & "Ryan DuPont"
CAPPY(6) = "Pretty much everything:" & vbCrLf & "Ryan DuPont"
CAPPY(7) = "I wanna thank Michelle for always being there..."
CAPPY(8) = "...And Mary for thinking my programs are evil..."
CAPPY(9) = "...And both of them for distracting me while I was making this"
CAPPY(10) = "Hmmm...I think i'm gonna dedicate this program"
CAPPY(11) = "to Michelle...I will always love you, no matter what"
CAPPY(12) = "Enjoy the game everyone!!!"
CAPPY(13) = ""
End Sub
Private Sub Timer1_Timer()
lblA.Caption = CAPPY(tmpCAP)
tmpCAP = tmpCAP + 1
If tmpCAP = 14 Then tmpCAP = 0
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Randomize
lblProd.ForeColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
lblProd.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub

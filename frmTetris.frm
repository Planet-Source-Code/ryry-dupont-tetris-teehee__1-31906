VERSION 5.00
Begin VB.Form frmTetris 
   BackColor       =   &H00404040&
   Caption         =   "Tetris"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   Icon            =   "frmTetris.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMENU 
      Interval        =   100
      Left            =   2640
      Top             =   3960
   End
   Begin VB.Timer tmrA 
      Left            =   2400
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Piece"
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
      Height          =   240
      Left            =   3000
      TabIndex        =   36
      Top             =   0
      Width           =   975
   End
   Begin VB.Image imgNext 
      Height          =   255
      Index           =   3
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgNext 
      Height          =   255
      Index           =   2
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgNext 
      Height          =   255
      Index           =   1
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgNext 
      Height          =   255
      Index           =   0
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblINS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
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
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   35
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label lblINS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ESC or P  -  Pause"
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
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   34
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lblINS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arrows  -  left/right"
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
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblINS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S  -  Rotate Left"
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
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblINS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A  -  Rotate Right"
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
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
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
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   30
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblSC 
      BackStyle       =   0  'Transparent
      Caption         =   "NA"
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
      Height          =   495
      Index           =   10
      Left            =   0
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      BackStyle       =   0  'Transparent
      Caption         =   "NA"
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
      Height          =   495
      Index           =   9
      Left            =   0
      TabIndex        =   28
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      BackStyle       =   0  'Transparent
      Caption         =   "NA"
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
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      BackStyle       =   0  'Transparent
      Caption         =   "NA"
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
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      BackStyle       =   0  'Transparent
      Caption         =   "NA"
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
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
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
      Height          =   495
      Index           =   5
      Left            =   360
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
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
      Height          =   495
      Index           =   4
      Left            =   840
      TabIndex        =   23
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
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
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
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
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
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
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblSC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SC"
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
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblMUS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblMUS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Music"
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblMUS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume 100%"
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Resume Game"
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
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores"
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
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
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
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Music"
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
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblMNU 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image S3 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image S3 
      Height          =   255
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image S3 
      Height          =   255
      Index           =   1
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image S4 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image S4 
      Height          =   255
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   255
   End
   Begin VB.Image S4 
      Height          =   255
      Index           =   1
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   255
   End
   Begin VB.Image S5 
      Height          =   255
      Index           =   3
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   255
   End
   Begin VB.Image S5 
      Height          =   255
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image S5 
      Height          =   255
      Index           =   1
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image S6 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image S6 
      Height          =   255
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   255
   End
   Begin VB.Image S6 
      Height          =   255
      Index           =   1
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   255
   End
   Begin VB.Image S7 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   255
   End
   Begin VB.Image S7 
      Height          =   255
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   255
   End
   Begin VB.Image S7 
      Height          =   255
      Index           =   1
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   255
   End
   Begin VB.Image S7 
      Height          =   255
      Index           =   0
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   255
   End
   Begin VB.Image S6 
      Height          =   255
      Index           =   0
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image S5 
      Height          =   255
      Index           =   0
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   255
   End
   Begin VB.Image S4 
      Height          =   255
      Index           =   0
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   255
   End
   Begin VB.Image S3 
      Height          =   255
      Index           =   0
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   255
   End
   Begin VB.Image S2 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image S2 
      Height          =   255
      Index           =   2
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image S2 
      Height          =   255
      Index           =   1
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image S2 
      Height          =   255
      Index           =   0
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image S1 
      Height          =   255
      Index           =   3
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image S1 
      Height          =   255
      Index           =   2
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   255
   End
   Begin VB.Image S1 
      Height          =   255
      Index           =   1
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image S1 
      Height          =   255
      Index           =   0
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image imgA 
      Height          =   255
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image pic0 
      Height          =   495
      Left            =   5160
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image pic7 
      Height          =   270
      Left            =   5160
      Picture         =   "frmTetris.frx":030A
      Top             =   5520
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic6 
      Height          =   270
      Left            =   5520
      Picture         =   "frmTetris.frx":071D
      Top             =   5520
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic5 
      Height          =   270
      Left            =   5160
      Picture         =   "frmTetris.frx":0B5A
      Top             =   5760
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic4 
      Height          =   270
      Left            =   5880
      Picture         =   "frmTetris.frx":0FCA
      Top             =   5160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic3 
      Height          =   270
      Left            =   5520
      Picture         =   "frmTetris.frx":1433
      Top             =   5160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic2 
      Height          =   270
      Left            =   5160
      Picture         =   "frmTetris.frx":1865
      Top             =   5880
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image pic1 
      Height          =   270
      Left            =   5160
      Picture         =   "frmTetris.frx":1CDD
      Top             =   5160
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   6
      Left            =   3960
      TabIndex        =   10
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   5
      Left            =   3960
      TabIndex        =   9
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   7
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblSHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level: 0"
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
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblLines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lines: 0"
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
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0"
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
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblA 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblB 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   2280
      TabIndex        =   37
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmTetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZAT
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'egg
If xMIKI = 0 And KeyCode = 88 Then xMIKI = 1
If xMIKI = 1 And KeyCode = 77 Then xMIKI = 2
If xMIKI = 2 And KeyCode = 73 Then xMIKI = 3
If xMIKI = 3 And KeyCode = 67 Then xMIKI = 4
If xMIKI = 4 And KeyCode = 72 Then xMIKI = 5
If xMIKI = 5 And KeyCode = 69 Then xMIKI = 6
If xMIKI = 6 And KeyCode = 76 Then xMIKI = 7
If xMIKI = 7 And KeyCode = 76 Then xMIKI = 8
If xMIKI = 8 And KeyCode = 69 Then xMIKI = 9
If xMIKI = 9 And KeyCode = 89 Then xMIKI = 10
If xMIKI = 10 Then Call MichelleX
'end egg
'display my menu stuff(way too complex...i should have commented while i was writing this)
If MENUview = True Then
    If MENUnum < 2 Then
        If KeyCode = 40 Then MENUsel = MENUsel + 1
        If KeyCode = 38 Then MENUsel = MENUsel - 1
    End If
    If MENUnum = 0 Then
        If KeyCode = 13 Then
            If MENUsel = 0 Then
                MENUview = False
                Call NewGame
            End If
            If MENUsel = 1 Then
                MENUnum = 1
                MENUsel = 0
                Exit Sub
            End If
            If MENUsel = 2 Then
                MENUview = False
                tmrA.Interval = INTY
                Exit Sub
            End If
            If MENUsel = 3 Then
                MENUsel = 5
                MENUnum = 2
                Call ScoreLoad
                Exit Sub
            End If
            If MENUsel = 4 Then
                MENUsel = 4
                MENUnum = 3
                Exit Sub
            End If
            If MENUsel = 5 Then End
        End If
        If MENUsel < 0 Then MENUsel = 5
        If MENUsel > 5 Then MENUsel = 0
    End If
    If MENUnum = 1 Then
        If KeyCode = 37 Then
            If MENUsel = 0 Then
                If VOL > 0 Then VOL = VOL - 1
                lblMUS(0).Caption = "Volume " & VOL & "%"
                frmTetris2.MP.Volume = (VOL * 50) - 5000
            End If
            If MENUsel = 1 Then
                On Error Resume Next
                SONGnum = SONGnum - 1
                If SONGnum < 0 Then SONGnum = SONGtotal
                frmTetris2.MP.FileName = SONGpath(SONGnum)
                lblMUS(1).Caption = SONGname(SONGnum)
            End If
        End If
        If KeyCode = 39 Then
            If MENUsel = 0 Then
                If VOL < 100 Then VOL = VOL + 1
                lblMUS(0).Caption = "Volume " & VOL & "%"
                frmTetris2.MP.Volume = (VOL * 50) - 5000
            End If
            If MENUsel = 1 Then
                On Error Resume Next
                SONGnum = SONGnum + 1
                If SONGnum > SONGtotal Then SONGnum = 0
                frmTetris2.MP.FileName = SONGpath(SONGnum)
                lblMUS(1).Caption = SONGname(SONGnum)
            End If
        End If
        If KeyCode = 13 Then
            If MENUsel = 2 Then
                MENUnum = 0
                MENUsel = 0
                Exit Sub
            End If
        End If
        If MENUsel < 0 Then MENUsel = 2
        If MENUsel > 2 Then MENUsel = 0
    End If
    If MENUnum = 2 And MENUsel > 5 Then
        If KeyCode = 8 Then
            On Error Resume Next
            lblSC(MENUsel).Caption = Left(lblSC(MENUsel).Caption, Len(lblSC(MENUsel).Caption) - 2) & "_"
            Exit Sub
        End If
        If KeyCode = 13 Then
            NAM(MENUsel - 6) = Left(lblSC(MENUsel).Caption, Len(lblSC(MENUsel).Caption) - 1)
            lblSC(MENUsel).Caption = Left(lblSC(MENUsel).Caption, Len(lblSC(MENUsel).Caption) - 1)
            Call WriteFile
            MENUsel = 5
            Exit Sub
        End If
        lblSC(MENUsel).Caption = Left(lblSC(MENUsel).Caption, Len(lblSC(MENUsel).Caption) - 1) & Chr(KeyCode) & "_"
    End If
    If MENUnum = 2 And MENUsel = 5 And KeyCode = 13 Then
        MENUsel = 0
        MENUnum = 0
    End If
    If MENUnum = 3 Then
        MENUsel = 4
        If KeyCode = 13 Then
            MENUsel = 0
            MENUnum = 0
            Exit Sub
        End If
    End If
End If
If MENUview = True Then Exit Sub
If KeyCode = 27 Or KeyCode = 80 Then 'pause the game
    tmrA.Interval = 0
    MENUview = True
    MENUsel = 2
    MENUnum = 0
    tmrMENU.Interval = 100
End If
If KeyCode = 65 Then Call GoSpin 'spin
If KeyCode = 37 Then Call GoLeft 'move piece left
If KeyCode = 39 Then Call GoRight 'move piece right
If KeyCode = 40 Then Call GoDown 'move piece down
If KeyCode = 83 Then Call GoSpin2 'spin opposite direction
End Sub
Private Sub Form_Load()
For z = 1 To 299 'load image array(no pics yet) and put them in right places
    Load imgA(z)
    imgA(z).Left = (z Mod 10) * 255
    imgA(z).Top = (z \ 10) * 255
    imgA(z).Stretch = True
    imgA(z).Visible = True
Next z
Load lblA(1)
lblA(1).Height = imgA(299).Top + imgA(299).Height
lblA(1).Width = imgA(299).Left + imgA(299).Width
lblA(1).Visible = True
lblB.Left = lblA(1).Width
For z = 300 To 339 'make sure bottom rows are occupied(needed to prevent errors when making sure a move/spin is legal)
    SP(z) = 30
Next z
For z = 0 To 3
    S1(z).Picture = pic1.Picture
    S2(z).Picture = pic2.Picture
    S3(z).Picture = pic3.Picture
    S4(z).Picture = pic4.Picture
    S5(z).Picture = pic5.Picture
    S6(z).Picture = pic6.Picture
    S7(z).Picture = pic7.Picture
Next z
MENUview = True
MENUsel = 0
MENUnum = 0
VOL = 100
INTY = 1000
Call SONGset 'play some music
If Dir$(App.Path & "\tet.dat") = "" Then Call WriteDefault 'if first time using, make default high scores file
Call OpenFile
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub tmrA_Timer()
Call GoDown 'move piece down
End Sub
Private Sub tmrMENU_Timer() 'more menu stuff
ZAT = ZAT + 1
If ZAT = 4 Then ZAT = 0
If MENUnum = 0 Then
    For z = 0 To 4
        lblINS(z).Visible = False
    Next z
    For z = 0 To 2
        lblMUS(z).Visible = False
    Next z
    For z = 0 To 10
        lblSC(z).Visible = False
    Next z
    For z = 0 To 5
        lblMNU(z).Visible = True
        If MENUsel = z Then
            If ZAT = 0 Then lblMNU(z).FontSize = 12
            If ZAT = 1 Then lblMNU(z).FontSize = 14
            If ZAT = 2 Then lblMNU(z).FontSize = 16
            If ZAT = 3 Then lblMNU(z).FontSize = 14
        Else
            lblMNU(z).FontSize = 12
        End If
    Next z
    If MENUview = False Then
        For z = 0 To 2
            lblMUS(z).Visible = False
        Next z
        For z = 0 To 5
            lblMNU(z).Visible = False
        Next z
        tmrMENU.Interval = 0
    End If
End If
If MENUnum = 1 Then
    For z = 0 To 4
        lblINS(z).Visible = False
    Next z
    For z = 0 To 5
        lblMNU(z).Visible = False
    Next z
    For z = 0 To 10
        lblSC(z).Visible = False
    Next z
    For z = 0 To 2
        lblMUS(z).Visible = True
        If MENUsel = z Then
            If ZAT = 0 Then lblMUS(z).FontSize = 12
            If ZAT = 1 Then lblMUS(z).FontSize = 14
            If ZAT = 2 Then lblMUS(z).FontSize = 16
            If ZAT = 3 Then lblMUS(z).FontSize = 14
        Else
            lblMUS(z).FontSize = 12
        End If
    Next z
End If
If MENUnum = 2 Then
    For z = 0 To 4
        lblINS(z).Visible = False
    Next z
    For z = 0 To 5
        lblMNU(z).Visible = False
    Next z
    For z = 0 To 2
        lblMUS(z).Visible = False
    Next z
    For z = 0 To 10
        lblSC(z).Visible = True
        If lblSC(z).Caption = "_" Then MENUsel = z
        If MENUsel = z Then
            If ZAT = 0 Then lblSC(z).FontSize = 12
            If ZAT = 1 Then lblSC(z).FontSize = 14
            If ZAT = 2 Then lblSC(z).FontSize = 16
            If ZAT = 3 Then lblSC(z).FontSize = 14
        Else
            lblSC(z).FontSize = 12
        End If
    Next z
End If
If MENUnum = 3 Then
    For z = 0 To 5
        lblMNU(z).Visible = False
    Next z
    For z = 0 To 2
        lblMUS(z).Visible = False
    Next z
    For z = 0 To 10
        lblSC(z).Visible = False
    Next z
    For z = 0 To 4
        lblINS(z).Visible = True
        If MENUsel = z Then
            If ZAT = 0 Then lblINS(z).FontSize = 12
            If ZAT = 1 Then lblINS(z).FontSize = 14
            If ZAT = 2 Then lblINS(z).FontSize = 16
            If ZAT = 3 Then lblINS(z).FontSize = 14
        Else
            lblINS(z).FontSize = 12
        End If
    Next z
End If
End Sub

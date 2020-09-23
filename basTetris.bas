Attribute VB_Name = "basTetris"
Public SP(339) As Integer 'place marker(0 if empty)
Public SHAPEnum As Integer, SPIN As Integer 'self explanatory(spin is between 1 and 4)
Public SPOT As Integer, NEXTshape As Integer 'spot=current marker position
Public isLEGAL As Boolean, isLEGALleft As Boolean 'temp vars for checking legal moves
Public isLEGALright As Boolean, isLEGALspin As Boolean ' " " " " legal spins
Public SCORE As Long, addSCORE As Integer 'current score/amt to add to score
Public LINN As Integer, LEV As Integer '# of lines/level
Public INTY As Integer, tmpLINN As Integer
Public PIElist(7) As Integer
Public MENUview As Boolean, MENUsel As Integer 'menu is on/what menu is selected
Public MENUnum As Integer, SONGnum As Integer 'which menu/song to play
Public SONGname(10), SONGpath(10), SONGtotal As Integer 'name(to display, not filename) and path of songs
Public VOL As Integer 'volume
Public SCO(4) As Long, tmpSCO, NAM(4), CHK(4)
Public xMIKI As Integer
Public HOWmany As Integer, WHICHone As Integer

Public Sub NewGame() 'set up for a new game
NEXTshape = Int(Rnd * 7) + 1
For z = 0 To 299
    SP(z) = 0
Next z
LINN = 0
LEV = 0
SCORE = 0
INTY = 1000
SPOT = 100
frmTetris.lblScore.Caption = "Score: 0"
frmTetris.lblLevel.Caption = "Level: 0"
frmTetris.lblLines.Caption = "Lines: 0"
For z = 0 To 6
    PIElist(z + 1) = 0
    frmTetris.lblSHP(z).Caption = "0"
Next z
Call StartShape
frmTetris.tmrA.Interval = 1000
End Sub
Public Sub GiveColor() 'set the pieces with their proper pictures
For z = 0 To 299
    If SP(z) Mod 7 = 0 Then frmTetris.imgA(z).Picture = frmTetris.pic7.Picture
    If SP(z) Mod 7 = 1 Then frmTetris.imgA(z).Picture = frmTetris.pic1.Picture
    If SP(z) Mod 7 = 2 Then frmTetris.imgA(z).Picture = frmTetris.pic2.Picture
    If SP(z) Mod 7 = 3 Then frmTetris.imgA(z).Picture = frmTetris.pic3.Picture
    If SP(z) Mod 7 = 4 Then frmTetris.imgA(z).Picture = frmTetris.pic4.Picture
    If SP(z) Mod 7 = 5 Then frmTetris.imgA(z).Picture = frmTetris.pic5.Picture
    If SP(z) Mod 7 = 6 Then frmTetris.imgA(z).Picture = frmTetris.pic6.Picture
    If SP(z) = 0 Then frmTetris.imgA(z).Picture = frmTetris.pic0.Picture
Next z
End Sub
Public Sub StartShape() 'choose a new shape
For z = 0 To 299
    If SP(z) > 0 And SP(z) < 8 Then SP(z) = SP(z) + 7
Next z
If SPOT < 10 Then Call GameOver
Call CheckLines
SPOT = 5
SPIN = 1
Randomize
SHAPEnum = NEXTshape
NEXTshape = Int(Rnd * 7) + 1
PIElist(SHAPEnum) = PIElist(SHAPEnum) + 1
For z = 1 To 7
    frmTetris.lblSHP(z - 1).Caption = PIElist(z)
Next z
Call DispNextShape
End Sub
Public Sub DispNextShape() 'start showing the new shape
If NEXTshape = 1 Then
    frmTetris.imgNext(0).Left = 3240
    frmTetris.imgNext(1).Left = 3480
    frmTetris.imgNext(2).Left = 3240
    frmTetris.imgNext(3).Left = 3480
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 360
    frmTetris.imgNext(2).Top = 600
    frmTetris.imgNext(3).Top = 600
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic1.Picture
    Next z
End If
If NEXTshape = 2 Then
    frmTetris.imgNext(0).Left = 3240
    frmTetris.imgNext(1).Left = 3480
    frmTetris.imgNext(2).Left = 3000
    frmTetris.imgNext(3).Left = 3240
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 360
    frmTetris.imgNext(2).Top = 600
    frmTetris.imgNext(3).Top = 600
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic2.Picture
    Next z
End If
If NEXTshape = 3 Then
    frmTetris.imgNext(0).Left = 3000
    frmTetris.imgNext(1).Left = 3240
    frmTetris.imgNext(2).Left = 3240
    frmTetris.imgNext(3).Left = 3480
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 360
    frmTetris.imgNext(2).Top = 600
    frmTetris.imgNext(3).Top = 600
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic3.Picture
    Next z
End If
If NEXTshape = 4 Then
    frmTetris.imgNext(0).Left = 3000
    frmTetris.imgNext(1).Left = 3240
    frmTetris.imgNext(2).Left = 3480
    frmTetris.imgNext(3).Left = 3720
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 360
    frmTetris.imgNext(2).Top = 360
    frmTetris.imgNext(3).Top = 360
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic4.Picture
    Next z
End If
If NEXTshape = 5 Then
    frmTetris.imgNext(0).Left = 3240
    frmTetris.imgNext(1).Left = 3240
    frmTetris.imgNext(2).Left = 3240
    frmTetris.imgNext(3).Left = 3480
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 600
    frmTetris.imgNext(2).Top = 840
    frmTetris.imgNext(3).Top = 840
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic5.Picture
    Next z
End If
If NEXTshape = 6 Then
    frmTetris.imgNext(0).Left = 3240
    frmTetris.imgNext(1).Left = 3240
    frmTetris.imgNext(2).Left = 3000
    frmTetris.imgNext(3).Left = 3240
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 600
    frmTetris.imgNext(2).Top = 840
    frmTetris.imgNext(3).Top = 840
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic6.Picture
    Next z
End If
If NEXTshape = 7 Then
    frmTetris.imgNext(0).Left = 3240
    frmTetris.imgNext(1).Left = 3000
    frmTetris.imgNext(2).Left = 3240
    frmTetris.imgNext(3).Left = 3480
    frmTetris.imgNext(0).Top = 360
    frmTetris.imgNext(1).Top = 600
    frmTetris.imgNext(2).Top = 600
    frmTetris.imgNext(3).Top = 600
    For z = 0 To 3
        frmTetris.imgNext(z).Picture = frmTetris.pic7.Picture
    Next z
End If
End Sub
Public Sub SH1()
On Error Resume Next
SP(SPOT) = 1
SP(SPOT + 1) = 1
SP(SPOT - 10) = 1
SP(SPOT - 9) = 1
End Sub
Public Sub SH2()
On Error Resume Next
SP(SPOT) = 2
If SPIN = 1 Or SPIN = 3 Then
    SP(SPOT - 1) = 2
    SP(SPOT - 9) = 2
    SP(SPOT - 10) = 2
End If
If SPIN = 2 Or SPIN = 4 Then
    SP(SPOT + 10) = 2
    SP(SPOT - 1) = 2
    SP(SPOT - 11) = 2
End If
End Sub
Public Sub SH3()
On Error Resume Next
SP(SPOT) = 3
If SPIN = 1 Or SPIN = 3 Then
    SP(SPOT + 1) = 3
    SP(SPOT - 10) = 3
    SP(SPOT - 11) = 3
End If
If SPIN = 2 Or SPIN = 4 Then
    SP(SPOT + 10) = 3
    SP(SPOT + 1) = 3
    SP(SPOT - 9) = 3
End If
End Sub
Public Sub SH4()
On Error Resume Next
SP(SPOT) = 4
If SPIN = 1 Or SPIN = 3 Then
    SP(SPOT + 10) = 4
    SP(SPOT - 10) = 4
    SP(SPOT - 20) = 4
End If
If SPIN = 2 Or SPIN = 4 Then
    SP(SPOT + 1) = 4
    SP(SPOT - 1) = 4
    SP(SPOT - 2) = 4
End If
End Sub
Public Sub SH5()
On Error Resume Next
SP(SPOT) = 5
If SPIN = 1 Then
    SP(SPOT + 1) = 5
    SP(SPOT - 10) = 5
    SP(SPOT - 20) = 5
End If
If SPIN = 2 Then
    SP(SPOT + 10) = 5
    SP(SPOT + 2) = 5
    SP(SPOT + 1) = 5
End If
If SPIN = 3 Then
    SP(SPOT + 20) = 5
    SP(SPOT + 10) = 5
    SP(SPOT - 1) = 5
End If
If SPIN = 4 Then
    SP(SPOT - 1) = 5
    SP(SPOT - 2) = 5
    SP(SPOT - 10) = 5
End If
End Sub
Public Sub SH6()
On Error Resume Next
SP(SPOT) = 6
If SPIN = 1 Then
    SP(SPOT - 1) = 6
    SP(SPOT - 10) = 6
    SP(SPOT - 20) = 6
End If
If SPIN = 2 Then
    SP(SPOT + 2) = 6
    SP(SPOT + 1) = 6
    SP(SPOT - 10) = 6
End If
If SPIN = 3 Then
    SP(SPOT + 20) = 6
    SP(SPOT + 10) = 6
    SP(SPOT + 1) = 6
End If
If SPIN = 4 Then
    SP(SPOT + 10) = 6
    SP(SPOT - 1) = 6
    SP(SPOT - 2) = 6
End If
End Sub
Public Sub SH7()
On Error Resume Next
SP(SPOT) = 7
If SPIN = 1 Then
    SP(SPOT + 1) = 7
    SP(SPOT - 1) = 7
    SP(SPOT - 10) = 7
End If
If SPIN = 2 Then
    SP(SPOT + 10) = 7
    SP(SPOT + 1) = 7
    SP(SPOT - 10) = 7
End If
If SPIN = 3 Then
    SP(SPOT + 10) = 7
    SP(SPOT + 1) = 7
    SP(SPOT - 1) = 7
End If
If SPIN = 4 Then
    SP(SPOT + 10) = 7
    SP(SPOT - 1) = 7
    SP(SPOT - 10) = 7
End If
End Sub
'LEGALxD    x=shape number(moving down[D]) make sure it is legal to move the shape down
Public Sub LEGAL1D()
If SP(SPOT + 10) = 0 And SP(SPOT + 11) = 0 Then isLEGAL = True
End Sub
Public Sub LEGAL2D()
If SPIN = 1 Or SPIN = 3 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 1) = 0 Then isLEGAL = True
End If
If SPIN = 2 Or SPIN = 4 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
End Sub
Public Sub LEGAL3D()
If SPIN = 1 Or SPIN = 3 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 11) = 0 And SP(SPOT - 1) = 0 Then isLEGAL = True
End If
If SPIN = 2 Or SPIN = 4 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 11) = 0 Then isLEGAL = True
End If
End Sub
Public Sub LEGAL4D()
If SPIN = 1 Or SPIN = 3 Then
    If SP(SPOT + 20) = 0 Then isLEGAL = True
End If
If SPIN = 2 Or SPIN = 4 Then
    If SP(SPOT + 11) = 0 And SP(SPOT + 10) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 8) = 0 Then isLEGAL = True
End If
End Sub
Public Sub LEGAL5D()
If SPIN = 1 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 11) = 0 Then isLEGAL = True
End If
If SPIN = 2 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 11) = 0 And SP(SPOT + 12) = 0 Then isLEGAL = True
End If
If SPIN = 3 Then
    If SP(SPOT + 30) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
If SPIN = 4 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 8) = 0 Then isLEGAL = True
End If
End Sub
Public Sub LEGAL6D()
If SPIN = 1 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
If SPIN = 2 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 11) = 0 And SP(SPOT + 12) = 0 Then isLEGAL = True
End If
If SPIN = 3 Then
    If SP(SPOT + 30) = 0 And SP(SPOT + 11) = 0 Then isLEGAL = True
End If
If SPIN = 4 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 8) = 0 Then isLEGAL = True
End If
End Sub
Public Sub LEGAL7D()
If SPIN = 1 Then
    If SP(SPOT + 11) = 0 And SP(SPOT + 10) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
If SPIN = 2 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 11) = 0 Then isLEGAL = True
End If
If SPIN = 3 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 11) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
If SPIN = 4 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 9) = 0 Then isLEGAL = True
End If
End Sub
'same as above, but moving right instead of down
Public Sub LEGAL1R()
On Error Resume Next
If SPOT Mod 10 < 8 Then
    If SP(SPOT + 2) = 0 And SP(SPOT - 8) = 0 Then isLEGALright = True
End If
End Sub
Public Sub LEGAL2R()
On Error Resume Next
If SPOT Mod 10 < 8 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT + 1) = 0 And SP(SPOT - 8) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 9 Then
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT + 1) = 0 And SP(SPOT + 11) = 0 And SP(SPOT - 10) = 0 Then isLEGALright = True
    End If
End If
End Sub
Public Sub LEGAL3R()
On Error Resume Next
If SPOT Mod 10 < 8 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT + 2) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT + 2) = 0 And SP(SPOT + 11) = 0 And SP(SPOT - 8) = 0 Then isLEGALright = True
    End If
End If
End Sub
Public Sub LEGAL4R()
On Error Resume Next
If SPOT Mod 10 < 9 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT + 11) = 0 And SP(SPOT + 1) = 0 And SP(SPOT - 9) = 0 And SP(SPOT - 19) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 8 Then
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT + 2) = 0 Then isLEGALright = True
    End If
End If
End Sub
Public Sub LEGAL5R()
On Error Resume Next
If SPOT Mod 10 < 8 Then
    If SPIN = 1 Then
        If SP(SPOT + 2) = 0 And SP(SPOT - 9) = 0 And SP(SPOT - 19) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 7 Then
    If SPIN = 2 Then
        If SP(SPOT + 3) = 0 And SP(SPOT + 11) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 9 Then
    If SPIN = 3 Then
        If SP(SPOT + 1) = 0 And SP(SPOT + 11) = 0 And SP(SPOT + 21) = 0 Then isLEGALright = True
    End If
    If SPIN = 4 Then
        If SP(SPOT + 1) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
End If
End Sub
Public Sub LEGAL6R()
On Error Resume Next
If SPOT Mod 10 < 9 Then
    If SPIN = 1 Then
        If SP(SPOT + 1) = 0 And SP(SPOT - 9) = 0 And SP(SPOT - 19) = 0 Then isLEGALright = True
    End If
    If SPIN = 4 Then
        If SP(SPOT + 1) = 0 And SP(SPOT + 11) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 7 Then
    If SPIN = 2 Then
        If SP(SPOT + 3) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 8 Then
    If SPIN = 3 Then
        If SP(SPOT + 2) = 0 And SP(SPOT + 11) = 0 And SP(SPOT + 21) = 0 Then isLEGALright = True
    End If
End If
End Sub
Public Sub LEGAL7R()
On Error Resume Next
If SPOT Mod 10 < 8 Then
    If SPIN = 1 Then
        If SP(SPOT + 2) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
    If SPIN = 2 Then
        If SP(SPOT + 11) = 0 And SP(SPOT + 2) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
    If SPIN = 3 Then
        If SP(SPOT + 11) = 0 And SP(SPOT + 2) = 0 Then isLEGALright = True
    End If
End If
If SPOT Mod 10 < 9 Then
    If SPIN = 4 Then
        If SP(SPOT + 11) = 0 And SP(SPOT + 1) = 0 And SP(SPOT - 9) = 0 Then isLEGALright = True
    End If
End If
End Sub
'same, but moving left
Public Sub LEGAL1L()
On Error Resume Next
If SPOT Mod 10 > 0 Then
    If SP(SPOT - 1) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
End If
End Sub
Public Sub LEGAL2L()
On Error Resume Next
If SPOT Mod 10 > 1 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT - 2) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 2) = 0 And SP(SPOT - 12) = 0 Then isLEGALleft = True
    End If
End If
End Sub
Public Sub LEGAL3L()
On Error Resume Next
If SPOT Mod 10 > 1 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 12) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 0 Then
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 1) = 0 And SP(SPOT - 10) = 0 Then isLEGALleft = True
    End If
End If
End Sub
Public Sub LEGAL4L()
On Error Resume Next
If SPOT Mod 10 > 0 Then
    If SPIN = 1 Or SPIN = 3 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 1) = 0 And SP(SPOT - 11) = 0 And SP(SPOT - 21) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 2 Then
    If SPIN = 2 Or SPIN = 4 Then
        If SP(SPOT - 3) = 0 Then isLEGALleft = True
    End If
End If
End Sub
Public Sub LEGAL5L()
On Error Resume Next
If SPOT Mod 10 > 0 Then
    If SPIN = 1 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 11) = 0 And SP(SPOT - 21) = 0 Then isLEGALleft = True
    End If
    If SPIN = 2 Then
        If SP(SPOT - 1) = 0 And SP(SPOT + 9) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 1 Then
    If SPIN = 3 Then
        If SP(SPOT - 2) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 19) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 2 Then
    If SPIN = 4 Then
        If SP(SPOT - 3) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
End If
End Sub
Public Sub LEGAL6L()
On Error Resume Next
If SPOT Mod 10 > 1 Then
    If SPIN = 1 Then
        If SP(SPOT - 2) = 0 And SP(SPOT - 11) = 0 And SP(SPOT - 21) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 0 Then
    If SPIN = 2 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
    If SPIN = 3 Then
        If SP(SPOT - 1) = 0 And SP(SPOT + 9) = 0 And SP(SPOT + 19) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 2 Then
    If SPIN = 4 Then
        If SP(SPOT - 3) = 0 And SP(SPOT + 9) = 0 Then isLEGALleft = True
    End If
End If
End Sub
Public Sub LEGAL7L()
On Error Resume Next
If SPOT Mod 10 > 1 Then
    If SPIN = 1 Then
        If SP(SPOT - 2) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
    If SPIN = 3 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 2) = 0 Then isLEGALleft = True
    End If
    If SPIN = 4 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 2) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
End If
If SPOT Mod 10 > 0 Then
    If SPIN = 2 Then
        If SP(SPOT + 9) = 0 And SP(SPOT - 1) = 0 And SP(SPOT - 11) = 0 Then isLEGALleft = True
    End If
End If
End Sub
'check to see if spinning the shape is legal
Public Sub LEGAL1S()
isLEGALspin = True
End Sub
Public Sub LEGAL2S()
On Error Resume Next
If SPIN = 1 Or SPIN = 3 Then
    If SP(SPOT - 11) = 0 And SP(SPOT + 10) = 0 Then isLEGALspin = True
End If
If SPIN = 2 Or SPIN = 4 Then
    If SPOT Mod 10 < 9 Then
        If SP(SPOT - 10) = 0 And SP(SPOT - 9) = 0 Then isLEGALspin = True
    End If
End If
End Sub
Public Sub LEGAL3S()
On Error Resume Next
If SPIN = 1 Or SPIN = 3 Then
    If SP(SPOT + 10) = 0 And SP(SPOT - 9) = 0 Then isLEGALspin = True
End If
If SPIN = 2 Or SPIN = 4 Then
    If SPOT Mod 10 > 0 Then
        If SP(SPOT - 11) = 0 And SP(SPOT - 10) = 0 Then isLEGALspin = True
    End If
End If
End Sub
Public Sub LEGAL4S()
On Error Resume Next
If SPIN = 1 Or SPIN = 3 Then
    If SPOT Mod 10 > 1 And SPOT Mod 10 < 9 Then
        If SP(SPOT + 1) = 0 And SP(SPOT - 1) = 0 And SP(SPOT - 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 2 Or SPIN = 4 Then
    If SP(SPOT + 10) = 0 And SP(SPOT - 10) = 0 And SP(SPOT - 20) = 0 Then isLEGALspin = True
End If
End Sub
Public Sub LEGAL5S()
On Error Resume Next
If SPIN = 1 Then
    If SPOT Mod 10 < 8 Then
        If SP(SPOT + 10) = 0 And SP(SPOT + 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 2 Then
    If SPOT Mod 10 > 0 Then
        If SP(SPOT - 1) = 0 And SP(SPOT + 20) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 3 Then
    If SPOT Mod 10 > 1 Then
        If SP(SPOT - 2) = 0 And SP(SPOT - 10) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 4 Then
    If SPOT Mod 10 < 9 Then
        If SP(SPOT - 20) = 0 And SP(SPOT + 1) = 0 Then isLEGALspin = True
    End If
End If
End Sub
Public Sub LEGAL6S()
On Error Resume Next
If SPIN = 1 Then
    If SPOT Mod 10 < 8 Then
        If SP(SPOT + 1) = 0 And SP(SPOT + 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 2 Then
    If SP(SPOT + 10) = 0 And SP(SPOT + 20) = 0 Then isLEGALspin = True
End If
If SPIN = 3 Then
    If SPOT Mod 10 > 1 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 4 Then
    If SP(SPOT - 10) = 0 And SP(SPOT - 20) = 0 Then isLEGALspin = True
End If
End Sub
Public Sub LEGAL7S()
On Error Resume Next
If SPIN = 1 Then
    If SP(SPOT + 10) = 0 Then isLEGALspin = True
End If
If SPIN = 2 Then
    If SPOT Mod 10 > 0 Then
        If SP(SPOT - 1) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 3 Then
    If SP(SPOT - 10) = 0 Then isLEGALspin = True
End If
If SPIN = 4 Then
    If SPOT Mod 10 < 9 Then
        If SP(SPOT + 1) = 0 Then isLEGALspin = True
    End If
End If
End Sub
'check to see if spin the other way is legal(only required for shape 5, 6, and 7
Public Sub LEGAL5S2()
On Error Resume Next
If SPIN = 1 Then
    If SPOT Mod 10 > 1 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 2 Then
    If SP(SPOT - 10) = 0 And SP(SPOT - 20) = 0 Then isLEGALspin = True
End If
If SPIN = 3 Then
    If SPOT Mod 10 < 8 Then
        If SP(SPOT + 1) = 0 And SP(SPOT + 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 4 Then
    If SP(SPOT + 20) = 0 And SP(SPOT + 10) = 0 Then isLEGALspin = True
End If
End Sub
Public Sub LEGAL6S2()
On Error Resume Next
If SPIN = 1 Then
    If SPOT Mod 10 > 1 Then
        If SP(SPOT + 10) = 0 And SP(SPOT - 2) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 2 Then
    If SPOT Mod 10 > 0 Then
        If SP(SPOT - 1) = 0 And SP(SPOT - 20) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 3 Then
    If SP(SPOT + 2) = 0 And SP(SPOT - 10) = 0 Then isLEGALspin = True
End If
If SPIN = 4 Then
    If SPOT Mod 10 < 9 Then
        If SP(SPOT + 20) = 0 And SP(SPOT + 1) = 0 Then isLEGALspin = True
    End If
End If
End Sub
Public Sub LEGAL7S2()
If SPIN = 1 Then
    If SP(SPOT + 10) = 0 Then isLEGALspin = True
End If
If SPIN = 2 Then
    If SPOT Mod 10 > 0 Then
        If SP(SPOT - 1) = 0 Then isLEGALspin = True
    End If
End If
If SPIN = 3 Then
    If SP(SPOT - 10) = 0 Then isLEGALspin = True
End If
If SPIN = 4 Then
    If SPOT Mod 10 < 9 Then
        If SP(SPOT + 1) = 0 Then isLEGALspin = True
    End If
End If
End Sub
Public Sub GoLeft()
isLEGALleft = False
If SHAPEnum = 1 Then Call LEGAL1L
If SHAPEnum = 2 Then Call LEGAL2L
If SHAPEnum = 3 Then Call LEGAL3L
If SHAPEnum = 4 Then Call LEGAL4L
If SHAPEnum = 5 Then Call LEGAL5L
If SHAPEnum = 6 Then Call LEGAL6L
If SHAPEnum = 7 Then Call LEGAL7L
If isLEGALleft = False Then Exit Sub
SPOT = SPOT - 1
For z = 0 To 299
    If SP(z) < 8 Then SP(z) = 0
Next z
If SHAPEnum = 1 Then Call SH1
If SHAPEnum = 2 Then Call SH2
If SHAPEnum = 3 Then Call SH3
If SHAPEnum = 4 Then Call SH4
If SHAPEnum = 5 Then Call SH5
If SHAPEnum = 6 Then Call SH6
If SHAPEnum = 7 Then Call SH7
Call GiveColor
End Sub
Public Sub GoRight()
isLEGALright = False
If SHAPEnum = 1 Then Call LEGAL1R
If SHAPEnum = 2 Then Call LEGAL2R
If SHAPEnum = 3 Then Call LEGAL3R
If SHAPEnum = 4 Then Call LEGAL4R
If SHAPEnum = 5 Then Call LEGAL5R
If SHAPEnum = 6 Then Call LEGAL6R
If SHAPEnum = 7 Then Call LEGAL7R
If isLEGALright = False Then Exit Sub
SPOT = SPOT + 1
For z = 0 To 299
    If SP(z) < 8 Then SP(z) = 0
Next z
If SHAPEnum = 1 Then Call SH1
If SHAPEnum = 2 Then Call SH2
If SHAPEnum = 3 Then Call SH3
If SHAPEnum = 4 Then Call SH4
If SHAPEnum = 5 Then Call SH5
If SHAPEnum = 6 Then Call SH6
If SHAPEnum = 7 Then Call SH7
Call GiveColor
End Sub
Public Sub GoSpin()
isLEGALspin = False
If SHAPEnum = 1 Then Call LEGAL1S
If SHAPEnum = 2 Then Call LEGAL2S
If SHAPEnum = 3 Then Call LEGAL3S
If SHAPEnum = 4 Then Call LEGAL4S
If SHAPEnum = 5 Then Call LEGAL5S
If SHAPEnum = 6 Then Call LEGAL6S
If SHAPEnum = 7 Then Call LEGAL7S
If isLEGALspin = False Then Exit Sub
SPIN = SPIN + 1
If SPIN = 5 Then SPIN = 1
For z = 0 To 299
    If SP(z) < 8 Then SP(z) = 0
Next z
If SHAPEnum = 1 Then Call SH1
If SHAPEnum = 2 Then Call SH2
If SHAPEnum = 3 Then Call SH3
If SHAPEnum = 4 Then Call SH4
If SHAPEnum = 5 Then Call SH5
If SHAPEnum = 6 Then Call SH6
If SHAPEnum = 7 Then Call SH7
Call GiveColor
End Sub
Public Sub GoSpin2()
isLEGALspin = False
If SHAPEnum = 1 Then Call LEGAL1S
If SHAPEnum = 2 Then Call LEGAL2S
If SHAPEnum = 3 Then Call LEGAL3S
If SHAPEnum = 4 Then Call LEGAL4S
If SHAPEnum = 5 Then Call LEGAL5S2
If SHAPEnum = 6 Then Call LEGAL6S2
If SHAPEnum = 7 Then Call LEGAL7S2
If isLEGALspin = False Then Exit Sub
SPIN = SPIN - 1
If SPIN = 0 Then SPIN = 4
For z = 0 To 299
    If SP(z) < 8 Then SP(z) = 0
Next z
If SHAPEnum = 1 Then Call SH1
If SHAPEnum = 2 Then Call SH2
If SHAPEnum = 3 Then Call SH3
If SHAPEnum = 4 Then Call SH4
If SHAPEnum = 5 Then Call SH5
If SHAPEnum = 6 Then Call SH6
If SHAPEnum = 7 Then Call SH7
Call GiveColor
End Sub
Public Sub GoDown()
isLEGAL = False
If SHAPEnum = 1 Then Call LEGAL1D
If SHAPEnum = 2 Then Call LEGAL2D
If SHAPEnum = 3 Then Call LEGAL3D
If SHAPEnum = 4 Then Call LEGAL4D
If SHAPEnum = 5 Then Call LEGAL5D
If SHAPEnum = 6 Then Call LEGAL6D
If SHAPEnum = 7 Then Call LEGAL7D
If isLEGAL = False Then Call StartShape: Exit Sub
SPOT = SPOT + 10
For z = 0 To 299
    If SP(z) < 8 Then SP(z) = 0
Next z
If SHAPEnum = 1 Then Call SH1
If SHAPEnum = 2 Then Call SH2
If SHAPEnum = 3 Then Call SH3
If SHAPEnum = 4 Then Call SH4
If SHAPEnum = 5 Then Call SH5
If SHAPEnum = 6 Then Call SH6
If SHAPEnum = 7 Then Call SH7
Call GiveColor
End Sub
Public Sub CheckLines() 'check to see if the player got a line/lines
addSCORE = 0
tmpLINN = 0
For z = 0 To 290 Step 10
    If SP(z) > 7 And SP(z + 1) > 7 And SP(z + 2) > 7 And SP(z + 3) > 7 And SP(z + 4) > 7 And SP(z + 5) > 7 And SP(z + 6) > 7 And SP(z + 7) > 7 And SP(z + 8) > 7 And SP(z + 9) > 7 Then
        LINN = LINN + 1
        tmpLINN = tmpLINN + 1
        If addSCORE > 0 Then addSCORE = addSCORE + 50 * (LEV + 1) 'multiline bonus
        addSCORE = addSCORE + 100 * (LEV + 1)
        For zz = z + 9 To 10 Step -1
            SP(zz) = SP(zz - 10)
        Next zz
        For zz = 0 To 9
            SP(zz) = 0
        Next zz
        Call GiveColor
        LEV = LINN \ 10
        INTY = 1000 - (20 * LEV)
        If INTY < 1 Then INTY = 1
        frmTetris.tmrA.Interval = INTY
        HOWmany = tmpLINN
    WHICHone = 0
    'Call DoGuyPic(a good idea gone bad...lol)
    SCORE = SCORE + addSCORE
    End If
Next z
frmTetris.lblScore.Caption = "Score: " & SCORE
frmTetris.lblLines.Caption = "Lines: " & LINN
frmTetris.lblLevel.Caption = "Level: " & LEV
End Sub
Public Sub GameOver() 'game over
tmpSCO = 10
For z = 4 To 0 Step -1
    If SCORE > SCO(z) Then tmpSCO = z
Next z
If tmpSCO = 10 Then GoTo NoHigh 'didnt get a high score
'did get a high score
For z = 3 To tmpSCO Step -1
    SCO(z + 1) = SCO(z)
    NAM(z + 1) = NAM(z)
Next z
SCO(tmpSCO) = SCORE
NoHigh:
frmTetris.tmrA.Interval = 0
If tmpSCO = 10 Then
    MENUsel = 5
Else
    MENUsel = tmpSCO + 6
    NAM(tmpSCO) = "_"
End If
MENUnum = 2
MENUview = True
Call ScoreLoad 'show high scores
frmTetris.tmrMENU.Interval = 100
End Sub
Public Sub SONGset() 'default settings for songs
On Error Resume Next
SONGname(0) = "Scariness"
SONGname(1) = "Flip-Out"
SONGname(2) = "Jump Around"
SONGname(3) = "A Forest"
SONGname(4) = "Lameness"
SONGtotal = 4
SONGpath(0) = App.Path & "\song1.mp3"
SONGpath(1) = App.Path & "\song2.mp3"
SONGpath(2) = App.Path & "\song3.mp3"
SONGpath(3) = App.Path & "\song4.mp3"
SONGpath(4) = App.Path & "\song5.mp3"
frmTetris.lblMUS(1).Caption = SONGname(0)
frmTetris2.MP.FileName = SONGpath(0)
End Sub
Public Sub OpenFile() 'open the high scores file
Open App.Path & "\tet.dat" For Input As #1
    Input #1, SCO(0)
    Input #1, NAM(0)
    Input #1, CHK(0)
    Input #1, SCO(1)
    Input #1, NAM(1)
    Input #1, CHK(1)
    Input #1, SCO(2)
    Input #1, NAM(2)
    Input #1, CHK(2)
    Input #1, SCO(3)
    Input #1, NAM(3)
    Input #1, CHK(3)
    Input #1, SCO(4)
    Input #1, NAM(4)
    Input #1, CHK(4)
Close #1
For z = 0 To 4
    If CHK(z) <> Trim(String(SCO(z) \ Asc(Left(NAM(z), 1)), "*")) Then Call WriteDefault
Next z
End Sub
Public Sub WriteFile() 'save the high scores
For z = 0 To 4
    CHK(z) = String(SCO(z) \ Asc(Left(NAM(z), 1)), "*")
Next z
Open App.Path & "\tet.dat" For Output As #1
    Print #1, SCO(0)
    Print #1, NAM(0)
    Print #1, CHK(0)
    Print #1, SCO(1)
    Print #1, NAM(1)
    Print #1, CHK(1)
    Print #1, SCO(2)
    Print #1, NAM(2)
    Print #1, CHK(2)
    Print #1, SCO(3)
    Print #1, NAM(3)
    Print #1, CHK(3)
    Print #1, SCO(4)
    Print #1, NAM(4)
    Print #1, CHK(4)
Close #1
End Sub
Public Sub WriteDefault() 'set up default high scores
SCO(0) = 300000
SCO(1) = 250000
SCO(2) = 150000
SCO(3) = 100000
SCO(4) = 50000
NAM(0) = "RyRy"
NAM(1) = "Mary"
NAM(2) = "Justin"
NAM(3) = "Evan"
NAM(4) = "Michelle"
For z = 0 To 4 'security check to prevent peeps from changing scores in high score file
    CHK(z) = String(SCO(z) \ Asc(Left(NAM(z), 1)), "*")
Next z
Open App.Path & "\tet.dat" For Output As #1
    Print #1, SCO(0)
    Print #1, NAM(0)
    Print #1, CHK(0)
    Print #1, SCO(1)
    Print #1, NAM(1)
    Print #1, CHK(1)
    Print #1, SCO(2)
    Print #1, NAM(2)
    Print #1, CHK(2)
    Print #1, SCO(3)
    Print #1, NAM(3)
    Print #1, CHK(3)
    Print #1, SCO(4)
    Print #1, NAM(4)
    Print #1, CHK(4)
Close #1
End Sub
Public Sub ScoreLoad() 'display high scores
For z = 0 To 4
    frmTetris.lblSC(z).Caption = SCO(z)
    frmTetris.lblSC(z + 6).Caption = NAM(z)
Next z
End Sub
Public Sub MichelleX()
frmTetris3.Show
End Sub
Public Sub DoGuyPic()
frmTetris4.Show
frmTetris4.Left = frmTetris.Left + frmTetris.Width
frmTetris4.Top = frmTetris.Top
End Sub

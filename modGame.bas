Attribute VB_Name = "modGame"
'36|Bomber v2.0 developed by Kevin Fleet
'20|DirectSound Class by Dave Katroski
'20|Graphics by Kurt Joseph Sage
'20|Copyright(R) 2002 KevCom

Declare Function GetTickCount Lib "kernel32" () As Long

Const AI_W = 300
Const AI_H = 50

Const sndExplo As Integer = 1
Const sndSplash As Integer = 2
Const sndShot As Integer = 3
Const sndPLShot As Integer = 4
Const sndsmExplo As Integer = 5

Const dRight = 0
Const dLeft = 1

Type Player
X As Long
Y As Long

XS As Long
YS As Long

Dir As Long

ReloadB As Long
ReloadS As Long

HP As Long
Lives As Long

Score As Long

Act As Long
End Type

Type Enemy
X As Long
Y As Long
Dir As Long
Act As Long
Reload As Long
End Type

Type Shot
X As Long
Y As Long
YS As Long
XS As Long

Life As Long
 
ExploThing As Long

Act As Long
End Type
 
Type Explosion
X As Long
Y As Long

XS As Long
YS As Long

W As Long
H As Long

Frame As Long
Frames As Long

T As Long

Act As Long
End Type

Type OsamaShip
X As Long

Act As Long
End Type

Type Powerup
X As Long
Y As Long

Type As Long
Count As Long

Act As Long
End Type

Public Player As Player
Public P(1 To 10) As Enemy
Public Sh(1 To 10) As Enemy
Public S(1 To 10) As Enemy
Public Ex(1 To 20) As Explosion
Public plShot(1 To 20) As Shot
Public ShShot(1 To 20) As Shot
Public B(1 To 10) As Shot
Public PShot(1 To 20) As Shot
Public E As New ClsEngine
Public Running As Boolean
Public Message As String
Public Box(1 To 10) As Shot

Type MenuItem
CAPTION As String
FUNC As String
End Type

Public M(1 To 5) As MenuItem
Public Selected As Long
Public Pause As Long
Public Osama As OsamaShip
Public Sound As New DS_Engine

Function MakeMenu()
Selected = 1
M(1).CAPTION = "New Game": M(1).FUNC = "new"
M(2).CAPTION = "Top Score": M(2).FUNC = "tscore"
M(3).CAPTION = "About": M(3).FUNC = "About"
M(5).CAPTION = "Exit": M(5).FUNC = "Exit"
M(4).CAPTION = "Tactics (Help)": M(4).FUNC = "Tact"
End Function

Function MakePlane()
Dim I As Long

For I = LBound(P()) To UBound(P())
If P(I).Act = False Then
P(I).Act = True
P(I).Reload = 0
P(I).Y = (frmGame.board.ScaleHeight * 0.2) + Int(Rnd * 50) + 1

Select Case Int(Rnd * 2)
Case 0
P(I).X = frmGame.board.ScaleWidth
P(I).Dir = dLeft
Case 1
P(I).X = -50
P(I).Dir = dRight
End Select
Exit For
End If
Next I
End Function

Function MakeShip()
Dim I As Long

For I = LBound(Sh()) To UBound(Sh())
If Sh(I).Act = False Then
Sh(I).Act = True
Sh(I).Reload = 0
Sh(I).Y = 203 - 20

Select Case Int(Rnd * 2)
Case 0
Sh(I).X = frmGame.board.ScaleWidth
Sh(I).Dir = dLeft
Case 1
Sh(I).X = -50
Sh(I).Dir = dRight
End Select
Exit For
End If
Next I
End Function

Function MakeSub()
Dim I As Long

For I = LBound(S()) To UBound(S())
If S(I).Act = False Then
S(I).Act = True
S(I).Y = 203 + 30 + (Int(Rnd * 70) + 1)

Select Case Int(Rnd * 2)
Case 0
S(I).X = frmGame.board.ScaleWidth
S(I).Dir = dLeft
Case 1
S(I).X = -50
S(I).Dir = dRight
End Select
Exit For
End If
Next I
End Function

Function MovePlayer()
Dim I As Long

If Player.HP <= 0 Then
Running = False
Message = "Game Over"
If Player.Score > Val(GetSetting("BOMBER", "TOPSCORE", "SCORE")) Then
SaveSetting "BOMBER", "TOPSCORE", "SCORE", Player.Score
SaveSetting "BOMBER", "TOPSCORE", "NAME", InputBox("You have beaten the TOP SCORE!" & vbCrLf & "Please enter your name for the records...", "You Won", "Player")
End If
End If

Player.ReloadS = Player.ReloadS - 1: If Player.ReloadS < 0 Then Player.ReloadS = 0
Player.ReloadB = Player.ReloadB - 1: If Player.ReloadB < 0 Then Player.ReloadB = 0

Player.X = Player.X + Player.XS
If Player.X < 0 Then Player.X = 0: Player.XS = 0
If Player.X > (frmGame.board.ScaleWidth - 50) Then Player.X = (frmGame.board.ScaleWidth - 50): Player.XS = 0

Player.Y = Player.Y + Player.YS
If Player.Y < 0 Then Player.Y = 0: Player.YS = 0
If Player.Y > (frmGame.board.ScaleHeight * 0.4) Then Player.Y = (frmGame.board.ScaleHeight * 0.4): Player.YS = 0

Select Case Player.XS
Case Is < 0
Player.XS = Player.XS + 1
Case Is > 0
Player.XS = Player.XS - 1
End Select

Select Case Player.YS
Case Is < 0
Player.YS = Player.YS + 1
Case Is > 0
Player.YS = Player.YS - 1
End Select

For I = 1 To 20
If plShot(I).Act = True And E.CollisionDetect(Player.X, Player.Y, 50, 30, Player.Dir * 50, 0, frmGame.playm.hdc, plShot(I).X, plShot(I).Y, 5, 5, 0, 0, frmGame.bulletm.hdc) = True Then
Player.HP = Player.HP - 2
plShot(I).Act = False
DoExplo plShot(I).X - 2, plShot(I).Y - 2, 3, 10, 10, 0, 0
E.PlaySound sndsmExplo
End If

If ShShot(I).Act = True And E.CollisionDetect(Player.X, Player.Y, 50, 30, Player.Dir * 50, 0, frmGame.playm.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, 0, 0, frmGame.bulletm.hdc) = True Then
Player.HP = Player.HP - 5
ShShot(I).Act = False
DoExplo ShShot(I).X - 2, ShShot(I).Y - 2, 3, 10, 10, 0, 0
E.PlaySound sndsmExplo
End If

If I < 11 Then
If Box(I).Act = True And E.CollisionDetect(Player.X, Player.Y, 50, 30, Player.Dir * 50, 0, frmGame.playm.hdc, Box(I).X, Box(I).Y, 20, 20, 0, 0, frmGame.boxM.hdc) = True Then
Player.HP = Player.HP + 10: If Player.HP > 50 Then Player.HP = 50
Player.Score = Player.Score + 200
Box(I).Act = False
End If
End If
Next I
End Function

Function ResetPlayer()
Player.Dir = dRight
Player.HP = 50
Player.ReloadB = 0
Player.ReloadS = 0
Player.X = 10
Player.Y = 10
Player.YS = 0
Player.XS = 10
End Function

Function NewGame()
Dim I As Long

ResetPlayer

Player.Act = True
Player.Score = 0

For I = 1 To 10
S(I).Act = False
Sh(I).Act = False
P(I).Act = False
B(I).Act = False
Box(I).Act = False
Next I

For I = 1 To 20
PShot(I).Act = False
ShShot(I).Act = False
plShot(I).Act = False
Next I

Running = True

frmGame.board.SetFocus
End Function

Function ShipShoot(Index)
If Sh(Index).Reload <> 0 Or Sh(Index).Act = False Then Exit Function

If Player.X > Sh(Index).X - AI_W And Player.X < Sh(Index).X + AI_W Then

Dim I As Long
For I = 1 To 20
If ShShot(I).Act = False Then
Sh(Index).Reload = 10
ShShot(I).Act = True
ShShot(I).Life = 70
ShShot(I).X = Sh(Index).X + 20
ShShot(I).Y = Sh(Index).Y + 5
ShShot(I).YS = -3
Exit For
End If
Next I

End If
End Function

Function PlaneShoot(Index)
If P(Index).Reload <> 0 Or P(Index).Act = False Then Exit Function

If Player.X > P(Index).X - AI_W And Player.X < P(Index).X + AI_W And Player.Y > P(Index).Y - AI_H And Player.Y < P(Index).Y + AI_H Then

Dim I As Long
For I = 1 To 20
If plShot(I).Act = False Then
E.PlaySound sndPLShot

plShot(I).Act = True
plShot(I).Life = 30
plShot(I).Y = P(Index).Y + 10

P(Index).Reload = 3

Select Case P(Index).Dir
Case dLeft
plShot(I).X = P(Index).X + 5
plShot(I).XS = -10
Case dRight
plShot(I).X = P(Index).X + 45
plShot(I).XS = 10
End Select
Exit For
End If
Next I

End If
End Function

Function PlayerShoot()
If Player.ReloadS <> 0 Or Player.Act = False Then Exit Function
Dim I As Long
For I = 1 To 20
If PShot(I).Act = False Then
E.PlaySound sndShot

PShot(I).Act = True
PShot(I).Life = 20
PShot(I).Y = Player.Y + 10

Player.ReloadS = 4

Select Case Player.Dir
Case dLeft
PShot(I).X = Player.X + 5
PShot(I).XS = -15
Case dRight
PShot(I).X = Player.X + 45
PShot(I).XS = 15
End Select
Exit For
End If
Next I
End Function

Function PlayerBomb()
If Player.ReloadB <> 0 Or Player.Act = False Then Exit Function
Dim I As Long
For I = 1 To 10
If B(I).Act = False Then
Player.ReloadB = 6
B(I).Act = True
B(I).Life = 90
B(I).Y = Player.Y + 5
B(I).YS = 5
B(I).ExploThing = False
B(I).X = Player.X + 15
Exit For
End If
Next I
End Function

Function MoveShots()
Dim I As Long, I2

For I = 1 To 20
If PShot(I).Act = True Then
PShot(I).Life = PShot(I).Life - 1: If PShot(I).Life <= 0 Then PShot(I).Act = False
PShot(I).X = PShot(I).X + PShot(I).XS
PShot(I).Y = PShot(I).Y + PShot(I).YS

For I2 = 1 To 10
If P(I2).Act = True And E.CollisionDetect(P(I2).X, P(I2).Y, 50, 30, P(I2).Dir * 50, 0, frmGame.eplanem.hdc, PShot(I).X, PShot(I).Y, 5, 5, 0, 0, frmGame.bulletm.hdc) = True Then
P(I2).Act = False
E.PlaySound sndExplo
DoExplo P(I2).X, P(I2).Y, 0, 30, 50, 0, 2
Player.Score = Player.Score + 100
End If
Next I2
End If

If I < 11 Then
If B(I).Act = True Then
B(I).Life = B(I).Life - 1: If B(I).Life <= 0 Then B(I).Act = False
B(I).X = B(I).X + B(I).XS
B(I).Y = B(I).Y + B(I).YS

If B(I).ExploThing = False And B(I).Y > 203 Then DoExplo B(I).X - 5, B(I).Y - 20, 1, 20, 20, 0, 0: B(I).ExploThing = True: B(I).YS = 2: E.PlaySound sndSplash

For I2 = 1 To 10
If S(I2).Act = True And E.CollisionDetect(S(I2).X, S(I2).Y, 50, 30, S(I2).Dir * 50, 0, frmGame.eSubm.hdc, B(I).X, B(I).Y, 5, 5, 0, 0, frmGame.bombm.hdc) = True Then
S(I2).Act = False
B(I).Act = False
E.PlaySound sndExplo
DoExplo S(I2).X, S(I2).Y, 2, 30, 50, 0, 2
Player.Score = Player.Score + 500
End If

If Sh(I2).Act = True And E.CollisionDetect(Sh(I2).X, Sh(I2).Y, 50, 30, Sh(I2).Dir * 50, 0, frmGame.eShipm.hdc, B(I).X, B(I).Y, 5, 5, 0, 0, frmGame.bombm.hdc) = True Then
Sh(I2).Act = False
B(I).Act = False
E.PlaySound sndExplo
Player.Score = Player.Score + 250
DoExplo Sh(I2).X, Sh(I2).Y, 0, 30, 50, 0, 1
End If

If P(I2).Act = True And E.CollisionDetect(P(I2).X, P(I2).Y, 50, 30, P(I2).Dir * 50, 0, frmGame.eplanem.hdc, B(I).X, B(I).Y, 5, 5, 0, 0, frmGame.bombm.hdc) = True Then
P(I2).Act = False
B(I).Act = False
E.PlaySound sndExplo
DoExplo P(I2).X, P(I2).Y, 0, 30, 50, 0, 2
Player.Score = Player.Score + 50
End If
Next I2

If Osama.Act = True And E.CollisionDetect(Osama.X, 203 - 45, 100, 50, 0, 0, frmGame.osamaM.hdc, B(I).X, B(I).Y, 5, 5, 0, 0, frmGame.bombm.hdc) = True Then
E.PlaySound sndExplo
Osama.Act = False
Osama.Act = False
B(I).Act = False
DoExplo Osama.X, 203 - 45, 5, 50, 100, 0, 2
End If
End If
End If

If plShot(I).Act = True Then
plShot(I).Life = plShot(I).Life - 1: If plShot(I).Life <= 0 Then plShot(I).Act = False
plShot(I).X = plShot(I).X + plShot(I).XS
plShot(I).Y = plShot(I).Y + plShot(I).YS
End If

If ShShot(I).Act = True Then
ShShot(I).Life = ShShot(I).Life - 1: If ShShot(I).Life <= 0 Then ShShot(I).Act = False
ShShot(I).X = ShShot(I).X + ShShot(I).XS
ShShot(I).Y = ShShot(I).Y + ShShot(I).YS
End If
Next I

For I = 1 To 10
If Box(I).Act = True Then
Box(I).Y = Box(I).Y + 2: If Box(I).Y > frmGame.board.ScaleHeight Then Box(I).Act = False
If Box(I).Y > 203 And Box(I).ExploThing = False Then DoExplo Box(I).X, Box(I).Y - 20, 1, 20, 20, 0, 0: E.PlaySound sndSplash: Box(I).ExploThing = True
End If
Next I
End Function

Function MoveEnemys()
Dim I As Long

For I = 1 To 10
If S(I).Act = True Then
Select Case S(I).Dir
Case dLeft
S(I).X = S(I).X - 5
Case dRight
S(I).X = S(I).X + 5
End Select

If S(I).X < -50 Then S(I).Act = False
If S(I).X > frmGame.board.ScaleWidth Then S(I).Act = False
End If

If P(I).Act = True Then
Select Case P(I).Dir
Case dLeft
P(I).X = P(I).X - 5
Case dRight
P(I).X = P(I).X + 5
End Select
P(I).Reload = P(I).Reload - 1: If P(I).Reload < 0 Then P(I).Reload = 0
If Int(Rnd * 5) = 3 Then PlaneShoot I

If Int(Rnd * 10) = Int(Rnd * 5) Then
If Player.Y < P(I).Y Then P(I).Y = P(I).Y - 5
If Player.Y > P(I).Y Then P(I).Y = P(I).Y + 5
End If

If Int(Rnd * 10) = Int(Rnd * 5) Then SwitchDir P(I)

If P(I).X < -50 Then P(I).Act = False
If P(I).X > frmGame.board.ScaleWidth Then P(I).Act = False
End If

If Sh(I).Act = True Then
Select Case Sh(I).Dir
Case dLeft
Sh(I).X = Sh(I).X - 5
Case dRight
Sh(I).X = Sh(I).X + 5
End Select
Sh(I).Reload = Sh(I).Reload - 1: If Sh(I).Reload < 0 Then Sh(I).Reload = 0
If Int(Rnd * 5) = 3 Then ShipShoot I

If Int(Rnd * 15) = Int(Rnd * 5) Then SwitchDir Sh(I)

If Sh(I).X < -50 Then Sh(I).Act = False
If Sh(I).X > frmGame.board.ScaleWidth Then Sh(I).Act = False
End If
Next I

If Osama.Act = True Then
Osama.X = Osama.X - 3: If Osama.X < -100 Then Osama.Act = False
End If
End Function

Function CheckPlayerInput()
If E.IsPressed(vbKeyLeft) = True Then Player.XS = -10: Player.Dir = dLeft
If E.IsPressed(vbKeyRight) = True Then Player.XS = 10: Player.Dir = dRight
If E.IsPressed(vbKeyUp) = True Then Player.YS = -5
If E.IsPressed(vbKeyDown) = True Then Player.YS = 5
If E.IsPressed(vbKeySpace) = True Then PlayerShoot
If E.IsPressed(vbKeyReturn) = True Then PlayerBomb
If E.IsPressed(vbKeyP) = True And Player.Score > 1000 Then MakeOsama
End Function

Function DoExplo(X, Y, T, H, W, Optional XS As Long, Optional YS As Long)
Dim I As Long, Frames As Long

Frames = frmGame.explo(T).ScaleWidth \ W

For I = 1 To 20
If Ex(I).Act = False Then
Ex(I).Act = True

Ex(I).T = T
Ex(I).Frame = 0
Ex(I).Frames = Frames
Ex(I).H = H
Ex(I).W = W
Ex(I).X = X
Ex(I).Y = Y
Ex(I).XS = XS
Ex(I).YS = YS
Exit For
End If
Next I
End Function

Function RunExplo()
Dim I As Long

For I = 1 To 20
If Ex(I).Act = True Then
Ex(I).X = Ex(I).X + Ex(I).XS
Ex(I).Y = Ex(I).Y + Ex(I).YS

Ex(I).Frame = Ex(I).Frame + 1: If Ex(I).Frame > Ex(I).Frames Then Ex(I).Act = False
End If
Next I
End Function

Function DrawMenu(board As PictureBox)
Dim Height, Width, I, cy, size
size = board.FontSize

board.FontSize = 10
Height = UBound(M()) * (10 + board.TextHeight("|"))
cy = board.ScaleHeight \ 2 - Height \ 2

For I = 1 To UBound(M())
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(M(I).CAPTION) \ 2
board.CurrentY = cy + ((I - 1) * (10 + board.TextHeight("|")))
board.ForeColor = vbBlack
If I = Selected Then board.ForeColor = vbRed
board.Print M(I).CAPTION
Next I
board.FontSize = size
End Function

Function DoItem()
Select Case UCase(M(Selected).FUNC)
Case "NEW"
frmGame.cmdNew_Click
Case "TSCORE"
frmGame.cmdTScore_Click
Case "EXIT"
frmGame.cmdEnd_Click
Case "ABOUT"
frmGame.cmdAbout_Click
Case "TACT"
If Running = True Then Pause = 9999999
Load frmTactics
End Select
End Function

Function MakeOsama()
If Osama.Act = True Then Exit Function
Osama.Act = True
Osama.X = frmGame.board.ScaleWidth
End Function

Function SwitchDir(obj As Enemy)
Select Case Player.X
Case Is < obj.X
obj.Dir = dLeft
Case Is > obj.X
obj.Dir = dRight
End Select
End Function

Function DropBox()
Dim I As Long
For I = 1 To 10
If Box(I).Act = False Then
Box(I).Act = True
Box(I).X = Int(Rnd * frmGame.board.ScaleWidth)
Box(I).ExploThing = False
If Box(I).X < 20 Then Box(I).X = 20
If Box(I).X > frmGame.board.ScaleWidth - 40 Then Box(I).X = frmGame.board.ScaleWidth - 40
Box(I).Y = -19
Exit For
End If
Next I
End Function

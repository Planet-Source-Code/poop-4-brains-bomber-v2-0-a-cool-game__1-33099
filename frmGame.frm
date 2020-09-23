VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bomber 2"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox boxM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9480
      Picture         =   "frmGame.frx":030A
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox boxS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9120
      Picture         =   "frmGame.frx":07FC
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   11160
      Picture         =   "frmGame.frx":0CEE
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   11640
      Picture         =   "frmGame.frx":188E0
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox osamas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   12240
      Picture         =   "frmGame.frx":304D2
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox osamaM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   10080
      Picture         =   "frmGame.frx":3445C
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":383E6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   4
      Left            =   9840
      Picture         =   "frmGame.frx":3CA78
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":4110A
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   3
      Left            =   11520
      Picture         =   "frmGame.frx":4173C
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":41D6E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   9840
      Picture         =   "frmGame.frx":498F8
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":51482
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":52C34
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox explom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":543E6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox explo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":58A78
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox eSubm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":5D10A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eSub 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":5F474
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShipm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":617DE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eShip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":63B48
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bombm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9960
      Picture         =   "frmGame.frx":65EB2
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox bomb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   9840
      Picture         =   "frmGame.frx":66034
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox eplanem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10320
      Picture         =   "frmGame.frx":661B6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox eplane 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   10200
      Picture         =   "frmGame.frx":68520
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox bulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9960
      Picture         =   "frmGame.frx":6A88A
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   9840
      Picture         =   "frmGame.frx":6A91C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox playm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9960
      Picture         =   "frmGame.frx":6A9AE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox play 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   9840
      Picture         =   "frmGame.frx":6CD18
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox board 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "frmGame.frx":6F082
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   -120
      Width           =   9000
      Begin VB.Timer tmr 
         Interval        =   1000
         Left            =   840
         Top             =   1200
      End
      Begin VB.PictureBox tscore 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   2640
         ScaleHeight     =   1335
         ScaleWidth      =   3495
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdOk 
            BackColor       =   &H00FF0000&
            Caption         =   "Ok"
            Height          =   255
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblScore 
            BackStyle       =   0  'Transparent
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1920
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Top Score:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Currently Held By:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Caption         =   "HotDogMan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdAbout 
         BackColor       =   &H000080FF&
         Caption         =   "About"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdTScore 
         BackColor       =   &H000080FF&
         Caption         =   "Top Score"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H000080FF&
         Caption         =   "New Game"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         BackColor       =   &H000080FF&
         Caption         =   "Exit"
         Height          =   255
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub cmdAbout_Click()
DoCredits
End Sub

Sub cmdEnd_Click()
E.EndSound
End
End Sub

Sub cmdNew_Click()
Randomize
NewGame
End Sub

Function MainLoop()
Dim c As Long, AddS, AddSh, AddPl, I, AddB

MakeMenu
E.StartSound
E.LoadSound App.Path + "\explo.wav", 1
E.LoadSound App.Path + "\splash.wav", 2
E.LoadSound App.Path + "\shot.wav", 3
E.LoadSound App.Path + "\shot.wav", 4
E.LoadSound App.Path + "\smlexplo.wav", 5

Do
If c > 1000 And GetTickCount > 500 Then
c = 0

If Pause > 0 Then
Pause = Pause - 1
Running = False
Message = "Paused"
If Pause = 1 Then Running = True
End If

If Running = True And Pause <= 0 Then

If AddS > 15 Then
AddS = 0
If Int(Rnd * 5) = 3 Then MakeSub
Else
AddS = AddS + 1
End If

If AddPl > 10 Then
AddPl = 0
If Int(Rnd * 5) = 3 Then MakePlane
Else
AddPl = AddPl + 1
End If

If AddB > 10 Then
AddB = 0
If Int(Rnd * 30) = Int(Rnd * 5) And Player.HP < (50 * 0.75) Then DropBox
Else
AddB = AddB + 1
End If


If AddSh > 15 Then
AddSh = 0
If Int(Rnd * 5) = 3 Then MakeShip
Else
AddSh = AddSh + 1
End If

CheckPlayerInput
MoveShots
MoveEnemys
MovePlayer
RunExplo

board.Cls

For I = 1 To 20
If PShot(I).Act = True Then
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, PShot(I).X, PShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If plShot(I).Act = True Then
E.DrawObj board.hdc, plShot(I).X, plShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, plShot(I).X, plShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If ShShot(I).Act = True Then
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bulletm.hdc, 0, 0, Mask
E.DrawObj board.hdc, ShShot(I).X, ShShot(I).Y, 5, 5, bullet.hdc, 0, 0, sprite
End If

If I < 11 Then
If B(I).Act = True Then
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bombm.hdc, 0, 0, Mask
E.DrawObj board.hdc, B(I).X, B(I).Y, 10, 10, bomb.hdc, 0, 0, sprite
End If
End If
Next I

For I = 1 To 10
If S(I).Act = True Then
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSubm.hdc, S(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, S(I).X, S(I).Y, 50, 30, eSub.hdc, S(I).Dir * 50, 0, sprite
End If

If P(I).Act = True Then
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplanem.hdc, P(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, P(I).X, P(I).Y, 50, 30, eplane.hdc, P(I).Dir * 50, 0, sprite
End If

If Sh(I).Act = True Then
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShipm.hdc, Sh(I).Dir * 50, 0, Mask
E.DrawObj board.hdc, Sh(I).X, Sh(I).Y, 50, 30, eShip.hdc, Sh(I).Dir * 50, 0, sprite
End If
Next I

If Osama.Act = True Then
E.DrawObj board.hdc, Osama.X, 203 - 45, 100, 50, osamaM.hdc, 0, 0, Mask
E.DrawObj board.hdc, Osama.X, 203 - 45, 100, 50, osamas.hdc, 0, 0, sprite
End If

E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, playm.hdc, Player.Dir * 50, 0, Mask
E.DrawObj board.hdc, Player.X, Player.Y, 50, 30, play.hdc, Player.Dir * 50, 0, sprite

For I = 1 To 10
If Box(I).Act = True Then
E.DrawObj board.hdc, Box(I).X, Box(I).Y, 20, 20, boxM.hdc, 0, 0, Mask
E.DrawObj board.hdc, Box(I).X, Box(I).Y, 20, 20, boxS.hdc, 0, 0, sprite
End If
Next I

For I = 1 To 20
If Ex(I).Act = True Then
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explom(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, Mask
E.DrawObj board.hdc, Ex(I).X, Ex(I).Y, Ex(I).W, Ex(I).H, explo(Ex(I).T).hdc, Ex(I).Frame * Ex(I).W, 0, sprite
End If
Next I

board.CurrentX = 10
board.CurrentY = 10
board.Print "Score: " & Player.Score
board.CurrentX = 10
board.CurrentY = 20 + board.TextHeight("|")
board.Print "HP: "
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + 50, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbRed, BF
board.Line (15 + board.TextWidth("HP: "), 20 + board.TextHeight("|"))-((15 + board.TextWidth("HP: ")) + Player.HP, (20 + board.TextHeight("|")) + board.TextHeight("|")), vbGreen, BF

Else
board.Cls

board.ForeColor = vbBlack
board.FontSize = 46
board.CurrentX = board.ScaleWidth \ 2 - board.TextWidth(Message) \ 2
board.CurrentY = board.ScaleHeight \ 2 - (board.TextHeight("|") * 2)
board.Print Message

board.FontSize = 10

If E.IsPressed(vbKeyDown) Then Selected = Selected + 1: If Selected > UBound(M()) Then Selected = UBound(M())
If E.IsPressed(vbKeyUp) Then Selected = Selected - 1: If Selected < LBound(M()) Then Selected = LBound(M())
If E.IsPressed(vbKeyControl) Then DoItem

DrawMenu board
End If
Else
c = c + 1
End If

If E.IsPressed(vbKeyF3) Then
Select Case Pause
Case Is <= 0: Pause = 9999999
Case Is > 0: Pause = 0: Running = True
End Select
End If

DoEvents
Loop
End Function

Private Sub cmdOk_Click()
tscore.Visible = False
End Sub

Sub cmdTScore_Click()
tscore.Visible = True
cmdOk.SetFocus
End Sub

Private Sub Form_Load()
Message = "Bomber"
Me.Visible = True
MainLoop
End Sub

Private Sub Form_Resize()
board.Left = 0
board.Top = 0
End Sub

Private Sub tmr_Timer()
lblScore.CAPTION = GetSetting("BOMBER", "TOPSCORE", "SCORE")
lblName.CAPTION = GetSetting("BOMBER", "TOPSCORE", "NAME")
End Sub

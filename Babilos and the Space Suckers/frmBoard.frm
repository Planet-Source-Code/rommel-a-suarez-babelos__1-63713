VERSION 5.00
Begin VB.Form FrmBoard 
   BackColor       =   &H80000007&
   Caption         =   "Babelos  and the Space Demon Suckers"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "frmBoard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Stat_checker 
      Interval        =   1
      Left            =   11760
      Top             =   8040
   End
   Begin VB.Timer bk_ground_Launcher 
      Interval        =   1000
      Left            =   9480
      Top             =   8400
   End
   Begin VB.Timer Enemy_Bullet_luncher 
      Interval        =   1000
      Left            =   12720
      Top             =   7560
   End
   Begin VB.Timer Enemy_Bullet_timer 
      Index           =   0
      Interval        =   1
      Left            =   10320
      Top             =   7200
   End
   Begin VB.Timer Bck_Ground_Timer 
      Index           =   0
      Interval        =   1
      Left            =   12720
      Top             =   6840
   End
   Begin VB.Timer Heroe_Gun_timer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   12480
      Top             =   7320
   End
   Begin VB.Timer Enemy_timer 
      Interval        =   1
      Left            =   10320
      Top             =   6720
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "Player:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13800
      TabIndex        =   7
      Top             =   10320
      Width           =   735
   End
   Begin VB.Label PlName 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   14640
      TabIndex        =   6
      Top             =   10320
      Width           =   615
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   29
      Left            =   9720
      Picture         =   "frmBoard.frx":0442
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   28
      Left            =   8760
      Picture         =   "frmBoard.frx":074C
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   27
      Left            =   7680
      Picture         =   "frmBoard.frx":0A56
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   26
      Left            =   6720
      Picture         =   "frmBoard.frx":0D60
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   25
      Left            =   5640
      Picture         =   "frmBoard.frx":106A
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   24
      Left            =   4560
      Picture         =   "frmBoard.frx":1374
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   23
      Left            =   3480
      Picture         =   "frmBoard.frx":167E
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   22
      Left            =   2400
      Picture         =   "frmBoard.frx":1988
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   21
      Left            =   1320
      Picture         =   "frmBoard.frx":1C92
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   20
      Left            =   240
      Picture         =   "frmBoard.frx":1F9C
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image EGun 
      Height          =   210
      Index           =   0
      Left            =   3360
      Picture         =   "frmBoard.frx":22A6
      Top             =   4080
      Width           =   210
   End
   Begin VB.Label LBLLevel 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   14640
      TabIndex        =   5
      Top             =   10080
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13800
      TabIndex        =   4
      Top             =   10080
      Width           =   735
   End
   Begin VB.Image bk_img_yellow 
      Height          =   255
      Index           =   0
      Left            =   3120
      Picture         =   "frmBoard.frx":2552
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image bk_Img_blue 
      Height          =   255
      Index           =   0
      Left            =   2160
      Picture         =   "frmBoard.frx":2908
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image bk_img 
      Height          =   255
      Index           =   0
      Left            =   1680
      Picture         =   "frmBoard.frx":2CBE
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Life:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13800
      TabIndex        =   3
      Top             =   9840
      Width           =   735
   End
   Begin VB.Label Life 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   14640
      TabIndex        =   2
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label LBLScore 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   14640
      TabIndex        =   1
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label xxx 
      BackColor       =   &H80000007&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   13800
      TabIndex        =   0
      Top             =   9600
      Width           =   975
   End
   Begin VB.Image Heroe_gun 
      Height          =   480
      Index           =   0
      Left            =   5160
      Picture         =   "frmBoard.frx":3074
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Heroe_Fleet 
      Height          =   480
      Left            =   5160
      Picture         =   "frmBoard.frx":34B6
      Top             =   8520
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   19
      Left            =   9720
      Picture         =   "frmBoard.frx":38F8
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   18
      Left            =   8760
      Picture         =   "frmBoard.frx":3C02
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   17
      Left            =   7680
      Picture         =   "frmBoard.frx":3F0C
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   16
      Left            =   6720
      Picture         =   "frmBoard.frx":4216
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   15
      Left            =   5640
      Picture         =   "frmBoard.frx":4520
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   14
      Left            =   4560
      Picture         =   "frmBoard.frx":482A
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   13
      Left            =   3480
      Picture         =   "frmBoard.frx":4B34
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   12
      Left            =   2400
      Picture         =   "frmBoard.frx":4E3E
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   11
      Left            =   1320
      Picture         =   "frmBoard.frx":5148
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   10
      Left            =   240
      Picture         =   "frmBoard.frx":5452
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   9
      Left            =   9720
      Picture         =   "frmBoard.frx":575C
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   8
      Left            =   8760
      Picture         =   "frmBoard.frx":5A66
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   7
      Left            =   7680
      Picture         =   "frmBoard.frx":5D70
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   6
      Left            =   6720
      Picture         =   "frmBoard.frx":607A
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   5
      Left            =   5640
      Picture         =   "frmBoard.frx":6384
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   4
      Left            =   4560
      Picture         =   "frmBoard.frx":668E
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   3
      Left            =   3480
      Picture         =   "frmBoard.frx":6998
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   2
      Left            =   2400
      Picture         =   "frmBoard.frx":6CA2
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   1
      Left            =   1320
      Picture         =   "frmBoard.frx":6FAC
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Enemy_Fleet 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmBoard.frx":72B6
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "FrmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////
'
'Programmed and designed by Rommel A. Suarez (HACKMEL)  December 9,2005 5:09 PM
'       Title : Babelos and the Space Demon Suckers
'  Description: A (GALAGA/GALAXIAN) like arcade
'               Demonstrate  basic collision detection on sprites/objects
'               Hope you'll learn something and have some fun on this mini arcade
'               Comments,questions and suggestions are freely accepted just send to Hackmel@yahoo.com
'
'   Dedication: To all people who inspires me and to those kids or even grown ups
'               who love the field of software development
'//////////////////////////////////////////////////////////////////////////////////////////


Private Sub Bck_Ground_Timer_Timer(Index As Integer)
Me.bk_img(Index).Top = Me.bk_img(Index).Top + 10
Me.bk_Img_blue(Index).Top = Me.bk_Img_blue(Index).Top + 10
Me.bk_img_yellow(Index).Top = Me.bk_img_yellow(Index).Top + 10
End Sub

Private Sub bk_ground_Launcher_Timer()
BK_STARS = BK_STARS + 1
Load Me.bk_img(BK_STARS)
Load Me.Bck_Ground_Timer(BK_STARS)
Me.bk_img(BK_STARS).Visible = True
Me.bk_img(BK_STARS).Top = Int((Me.ScaleHeight - 1 + 1) * Rnd + 1)
Me.bk_img(BK_STARS).Left = Int((Me.ScaleWidth - 1 + 1) * Rnd + 1)

Load Me.bk_Img_blue(BK_STARS)

Me.bk_Img_blue(BK_STARS).Visible = True
Me.bk_Img_blue(BK_STARS).Top = Int((Me.ScaleHeight - 1 + 1) * Rnd + 1)
Me.bk_Img_blue(BK_STARS).Left = Int((Me.ScaleWidth - 1 + 1) * Rnd + 1)


Load Me.bk_img_yellow(BK_STARS)

Me.bk_img_yellow(BK_STARS).Visible = True
Me.bk_img_yellow(BK_STARS).Top = Int((Me.ScaleHeight - 1 + 1) * Rnd + 1)
Me.bk_img_yellow(BK_STARS).Left = Int((Me.ScaleWidth - 1 + 1) * Rnd + 1)

End Sub

Private Sub Enemy_Bullet_luncher_Timer()
EGUN_INDEX = EGUN_INDEX + 1
Load Me.EGun(EGUN_INDEX)
Load Me.Enemy_Bullet_timer(EGUN_INDEX)
Me.EGun(EGUN_INDEX).Visible = True
Me.EGun(EGUN_INDEX).Left = Int((Me.ScaleWidth - 1 + 1) * Rnd + 1)
Me.EGun(EGUN_INDEX).Top = 1320
End Sub

Private Sub Enemy_Bullet_timer_Timer(Index As Integer)
Me.EGun(Index).Top = Me.EGun(Index).Top + 10

Me.Life = HLIFE
If isHeroe_Hit(Me.EGun(Index)) Then
HLIFE = HLIFE - 1
End If
End Sub

Private Sub Enemy_timer_Timer()
Dim i As Integer
If Me.Enemy_Fleet(0).Left < 0 Then
    DIR = RIGHT_DIR
End If
 
If Me.Enemy_Fleet(29).Left > Me.ScaleWidth - 1220 Then
    DIR = LEFT_DIR
End If
 
If Me.Enemy_Fleet(0).Top <= 600 Then
    MOVEMENT = FORWARD
End If
 
If Me.Enemy_Fleet(29).Top >= 8400 Then
    MOVEMENT = DOWNWARD
End If

   
For i = 0 To ENEMY_COUNT
 If DIR = LEFT_DIR Then
   Me.Enemy_Fleet(i).Left = Me.Enemy_Fleet(i).Left - 100
 Else
   Me.Enemy_Fleet(i).Left = Me.Enemy_Fleet(i).Left + 100
 End If
 
 If MOVEMENT = FORWARD Then
      Me.Enemy_Fleet(i).Top = Me.Enemy_Fleet(i).Top + 10
 Else
      Me.Enemy_Fleet(i).Top = Me.Enemy_Fleet(i).Top - 10
 End If
 
 If isHeroe_Hit(Me.Enemy_Fleet(i)) Then
   HLIFE = HLIFE - 1
 End If
Next
End Sub

Private Sub Form_Activate()
FrmBoard.Heroe_Fleet.Visible = True
FrmBoard.Stat_checker.Enabled = True
FrmBoard.Enemy_Bullet_luncher.Enabled = True
refreshGame
Me.PlName = Player_Name
objScore.Player = Player_Name

If objScore.isPlayerExist And Level = 1 Then
   If MsgBox("The Player name you entered already exist,do you want to overwrite?", vbQuestion + vbYesNo) = vbYes Then
    objScore.EraseScore
   Else
    Unload Me
    FrmSplash.Show vbModal
   End If
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 37 Then
   Me.Heroe_Fleet.Left = Me.Heroe_Fleet.Left - 100
ElseIf KeyCode = 39 Then
   Me.Heroe_Fleet.Left = Me.Heroe_Fleet.Left + 100
ElseIf KeyCode = 38 Then
    Me.Heroe_Fleet.Top = Me.Heroe_Fleet.Top - 100
ElseIf KeyCode = 40 Then
    Me.Heroe_Fleet.Top = Me.Heroe_Fleet.Top + 100
ElseIf KeyCode = 17 Then
   HGUN_INDEX = HGUN_INDEX + 1
   Load Heroe_gun(HGUN_INDEX)
   Load Heroe_Gun_timer(HGUN_INDEX)
   
   Me.Heroe_gun(HGUN_INDEX).Top = Me.Heroe_Fleet.Top + 100
   Me.Heroe_gun(HGUN_INDEX).Left = Me.Heroe_Fleet.Left
   Heroe_gun(HGUN_INDEX).Visible = True
   Heroe_Gun_timer(HGUN_INDEX).Enabled = True
   
End If
End Sub

Private Sub Form_Load()
EGUN_INDEX = 0
HLIFE = 3
Level = 0
End Sub



Function isEnemy_Hit(Gun As Image) As Boolean
Dim i As Integer
Dim x As Integer
Dim deltaX As Integer
Dim deltay As Integer

isEnemy_Hit = False

 If Gun.Visible Then
    For x = 0 To ENEMY_COUNT
        deltaX = Abs(Me.Enemy_Fleet(x).Left - Gun.Left)
        deltay = Abs(Me.Enemy_Fleet(x).Top - Gun.Top)
       
       If Me.Enemy_Fleet(x).Visible Then
            If deltaX <= 480 / 2 And deltay <= 480 / 2 Then
                 Me.Enemy_Fleet(x).Visible = False
                 Me.explode Me.Enemy_Fleet(x).Left, Me.Enemy_Fleet(x).Top
                 Gun.Visible = False
                 isEnemy_Hit = True
                 Exit For
            End If
       End If
    deltaX = 0
    deltay = 0
    
 Next
 End If
End Function


Sub explode(x As Integer, y As Integer)
Dim i As Integer
For i = 0 To 1000
  Me.ForeColor = i
  Me.Circle (x, y), i
Next
For i = 0 To 1000
  Me.ForeColor = vblack
  Me.Circle (x, y), i
Next
End Sub


Function isHeroe_Hit(Gun As Image) As Boolean
Dim i As Integer
Dim x As Integer
Dim deltaX As Long
Dim deltay As Long

isHeroe_Hit = False
 If Gun.Visible Then
    
        deltaX = Abs(Me.Heroe_Fleet.Left - Gun.Left)
        deltay = Abs(Me.Heroe_Fleet.Top - Gun.Top)
       
           If Me.Heroe_Fleet.Visible Then
            If deltaX < (480 + Gun.Width / 2) And deltay < (Gun.Width / 2) Then
                Me.Heroe_Fleet.Visible = False
                explode Me.Heroe_Fleet.Left, Me.Heroe_Fleet.Top
                Gun.Visible = False
                isHeroe_Hit = True
                Me.Heroe_Fleet.Top = 8520
                Me.Heroe_Fleet.Left = Me.ScaleWidth / 2
                Me.Heroe_Fleet.Visible = True
                
            End If
           End If
 End If
 
End Function


Private Sub Heroe_Gun_timer_Timer(Index As Integer)
Me.Heroe_gun(Index).Top = Me.Heroe_gun(Index).Top - 10

If Heroe_gun(Index).Top > 0 Then
  If isEnemy_Hit(Heroe_gun(Index)) Then
      Score = Score + 1
      TOTAL_SCORE = TOTAL_SCORE + 1
      Me.LBLScore = TOTAL_SCORE
  End If
End If
End Sub

Private Sub Stat_checker_Timer()
If HLIFE <= 0 Then
  Stat_checker.Enabled = False
  Me.Enemy_Bullet_luncher.Enabled = False
  Me.Heroe_Fleet.Visible = False
  frmStat.Label1.Caption = "You Lost"
  frmStat.Show vbModal
Else
   If Score >= 30 Then
    Score = 0
    Stat_checker.Enabled = False
    Me.Heroe_Fleet.Visible = False
    Me.Enemy_Bullet_luncher.Enabled = False
    frmStat.Label1.Caption = "You Won"
    frmStat.Show vbModal
   End If
End If

Me.LBLLevel = Level
End Sub




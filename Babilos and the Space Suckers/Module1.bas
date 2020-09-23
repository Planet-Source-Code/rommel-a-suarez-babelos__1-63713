Attribute VB_Name = "BabelosModule"
'///////////////////////////////////////////////////////////////////////////////////
'/  Programmed and designed by Rommel A. Suarez (HACKMEL)  December 9,2005 5:09 PM
'/      Title :     Babelos and the Space Demon Suckers
'/  Description:    Written here the public declarations of variables,objects,methods and procedures
'/
'//////////////////////////////////////////////////////////////////////////////////////////



Public Level As Integer
Public DIR As Integer, MOVEMENT As Integer
Public EGUN_INDEX As Integer, HGUN_INDEX As Integer, BK_STARS As Integer
Public Const RIGHT_DIR = 1
Public Const LEFT_DIR = 2
Public Const FORWARD = 1
Public Const DOWNWARD = 0
Public Const ENEMY_COUNT = 29
Public Score As Integer, TOTAL_SCORE As Integer, HLIFE As Integer
Public Player_Name As String
Public objScore As New ScoreLogger
Public INPLAY As Integer



Sub Enemy_Revive()
Dim i As Integer
For i = 0 To ENEMY_COUNT
   FrmBoard.Enemy_Fleet(i).Visible = True
Next
End Sub


Sub refreshEnemyBullets()
Dim i As Long
For i = 1 To EGUN_INDEX
   Unload FrmBoard.Enemy_Bullet_timer(i)
   Unload FrmBoard.EGun(i)
Next
EGUN_INDEX = 0
End Sub

Sub refreshHeroeBullets()
Dim i As Long
For i = 1 To HGUN_INDEX
   Unload FrmBoard.Heroe_gun(i)
   Unload FrmBoard.Heroe_Gun_timer(i)
Next
HGUN_INDEX = 0
End Sub

Sub refreshAsteroids()
Dim i As Long
For i = 1 To BK_STARS
   Unload FrmBoard.Bck_Ground_Timer(i)
   Unload FrmBoard.bk_img(i)
   Unload FrmBoard.bk_Img_blue(i)
   Unload FrmBoard.bk_img_yellow(i)
   
Next
BK_STARS = 0
End Sub

Sub refreshGame()
    refreshHeroeBullets
    refreshEnemyBullets
    refreshAsteroids
    Enemy_Revive
    If Level = 0 Then Level = 1
    
End Sub

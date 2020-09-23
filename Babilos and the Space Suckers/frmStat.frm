VERSION 5.00
Begin VB.Form frmStat 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ScoreTab 
      Caption         =   "Scores"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton End 
      BackColor       =   &H80000007&
      Caption         =   "&End"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton continue 
      BackColor       =   &H80000007&
      Caption         =   "&Continue"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Programmed and Designed by Rommel A. Suarez"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   7800
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Babelos and the Space Demon Suckers ver.1.0"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   36
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000007&
      BorderColor     =   &H00000080&
      Height          =   1455
      Left            =   240
      Top             =   5040
      Width           =   6255
   End
   Begin VB.Label LBLLife 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "Life:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label LBLScores 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      Caption         =   "Scores:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000007&
      BorderColor     =   &H00000080&
      Height          =   3015
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   6015
   End
End
Attribute VB_Name = "frmStat"
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

Dim NextLevel As Boolean
Private Sub continue_Click()
If NextLevel Then
    Select Case Level
      Case 2
        FrmBoard.Enemy_Bullet_luncher.Interval = 900
      Case 3
        FrmBoard.Enemy_Bullet_luncher.Interval = 800
      Case 4
        FrmBoard.Enemy_Bullet_luncher.Interval = 700
      Case 5
        FrmBoard.Enemy_Bullet_luncher.Interval = 600
      Case 6
        FrmBoard.Enemy_Bullet_luncher.Interval = 500
      Case 7
        FrmBoard.Enemy_Bullet_luncher.Interval = 400
      Case 8
        FrmBoard.Enemy_Bullet_luncher.Interval = 300
      Case 9
        FrmBoard.Enemy_Bullet_luncher.Interval = 200
      Case 10
        FrmBoard.Enemy_Bullet_luncher.Interval = 100
    End Select
Else
   FrmBoard.Enemy_Bullet_luncher.Interval = 1000
   TOTAL_SCORE = 0
   HLIFE = 3
   Level = 0
End If
  
Unload Me
'continue_Click
End Sub

Private Sub End_Click()
End
End Sub

Private Sub Form_Activate()
If Me.Label1.Caption = "You Won" Then
   NextLevel = True
   
   
   
   Level = Level + 1
   If Level < 10 Then
      NextLevel = True
      Label2.Caption = "Proceed to Level " & Level
   Else
      NextLevel = False
      Label2.Caption = "Congratulations you reached " & Level
   End If
Else
NextLevel = False
Me.Label2.Caption = "Sorry Try again"
End If
Me.LBLScores = TOTAL_SCORE
Me.LBLLife = HLIFE

objScore.writeScores TOTAL_SCORE, Level
End Sub



Private Sub ScoreTab_Click()
FrmScoresTab.Show vbModal
End Sub

VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Babelos and the Space Demon Suckers "
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ScoreTab 
      Caption         =   "Scores"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Pname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Player Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4920
      Width           =   6735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Press [ESCAPE] to End"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   6360
      Width           =   6495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Press [Enter] to Continue"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   21.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   5760
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Hackmel@yahoo.com"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "December 9,2005 5:09 PM"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      Height          =   3735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7215
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
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Programmed and Designed by Rommel A. Suarez"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   2640
      Width           =   3495
   End
End
Attribute VB_Name = "FrmSplash"
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


Private Sub Form_Activate()
Me.Pname.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   Player_Name = Me.Pname.Text
   Unload Me
   FrmBoard.Show vbModal
   
End If
End Sub

Private Sub Pname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Form_KeyPress 13
Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub ScoreTab_Click()
FrmScoresTab.Show vbModal
End Sub

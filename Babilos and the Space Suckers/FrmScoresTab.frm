VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmScoresTab 
   Caption         =   "Score Tabulations"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5820
   Icon            =   "FrmScoresTab.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7435
      SortKey         =   1
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Player Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Score"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Reached Level"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FrmScoresTab"
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


Private Sub Form_Load()
objScore.FillListWithScoresLog Me.ListView1
End Sub

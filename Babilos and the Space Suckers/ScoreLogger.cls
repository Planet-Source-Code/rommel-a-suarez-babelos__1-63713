VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScoreLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''///////////////////////////////////////////////////////////////////////////////////
'/  Programmed and designed by Rommel A. Suarez (HACKMEL)  December 9,2005 5:09 PM
'/      Title :     Babelos and the Space Demon Suckers
'/  Description:    A class module for writing,retrieving,deleting records of a player such as Name,Scores and
'/                  date of play in a text file using random access file method
'//////////////////////////////////////////////////////////////////////////////////////////


Dim PlayerName As String

Private Type PlayerRecord
    Player_Name As String * 10
    Score As String * 10
    Date As String * 30
    Level As String * 10
End Type

Dim Player_Rec As PlayerRecord


Public Sub FillListWithScoresLog(L As ListView)
'////////////////////////////////////////////
'/  Displays the scores from the Scores.txt
'/  into a list view
'///////////////////////////////////////////

Dim pathname: pathname = App.Path & "\Scores.txt"
Dim i As Integer
Dim iFileNo As Integer
Dim item As ListItem

iFileNo = FreeFile

Open pathname For Random As #iFileNo Len = Len(Player_Rec)
L.ListItems.Clear
i = 1
Do While Not EOF(iFileNo)
   Get #iFileNo, i, Player_Rec
    
    If Trim(Player_Rec.Player_Name) <> "          " Then
        Set item = L.ListItems.Add(, , Trim(Player_Rec.Player_Name))
        item.SubItems(1) = Trim(Player_Rec.Score)
        item.SubItems(2) = Trim(Player_Rec.Date)
        item.SubItems(3) = Trim(Player_Rec.Level)
        i = i + 1
    End If
Loop

Close #iFileNo
End Sub



Public Function isPlayerExist() As Boolean
'////////////////////////////////////////////
'/  Checks wether a player's name exists
'/  on the Scores.txt
'///////////////////////////////////////////
Dim pathname: pathname = App.Path & "\Scores.txt"
Dim iFileNo As Integer
Dim i As Integer

iFileNo = FreeFile
isPlayerExist = False

Open pathname For Random As #iFileNo Len = Len(Player_Rec)
i = 1
Do Until EOF(iFileNo)
    Get #iFileNo, i, Player_Rec
    If Trim(Player_Rec.Player_Name) = Trim(Player) Then
      isPlayerExist = True
      Close #iFileNo
      Exit Function
    End If
    i = i + 1
Loop

Close #iFileNo
End Function




Public Sub writeScores(ByVal Score As Integer, Level As Integer)
'////////////////////////////////////////////
'/  Write the scores into the Scores.txt
'/
'///////////////////////////////////////////
Dim pathname: pathname = App.Path & "\Scores.txt"
Dim iFileNo As Integer
Dim i As Integer
i = 1
iFileNo = FreeFile
isUpdate = False

Open pathname For Random As #iFileNo Len = Len(Player_Rec)
     Do Until EOF(iFileNo)
            Get #iFileNo, i, Player_Rec
            If Trim(Player_Rec.Player_Name) = Trim(Player) Then
               Player_Rec.Score = RTrim(LTrim(CStr(Score)))
               Player_Rec.Level = RTrim(LTrim(CStr(Level)))
               Player_Rec.Date = Trim(CStr(Now))
               Put #iFileNo, i, Player_Rec
               Close #iFileNo
               
               Exit Sub
            End If
            i = i + 1
    Loop
   Player_Rec.Player_Name = Trim(Player)
   Player_Rec.Score = RTrim(LTrim(CStr(Score)))
   Player_Rec.Date = Trim(CStr(Now))
   Player_Rec.Level = RTrim(LTrim(CStr(Level)))
   Put #iFileNo, , Player_Rec


Close #iFileNo


End Sub


Public Sub EraseScore()
'////////////////////////////////////////////
'/  Erase the scores from the Scores.txt
'/
'///////////////////////////////////////////
Dim pathname: pathname = App.Path & "\Scores.txt"
Dim iFileNo As Integer
Dim i As Integer
i = 1
iFileNo = FreeFile


Open pathname For Random As #iFileNo Len = Len(Player_Rec)
     Do Until EOF(iFileNo)
            Get #iFileNo, i, Player_Rec
            If Trim(Player_Rec.Player_Name) = Trim(Player) Then
               
               Player_Rec.Score = "0"
               Player_Rec.Level = "0"
               
               Player_Rec.Date = CStr(Now)
               Put #iFileNo, i, Player_Rec
               Close #iFileNo
               Exit Sub
            End If
            i = i + 1
    Loop
End Sub

'////////////////////////////////////////////
'/  ScoreLogger Class Propertites
'/
'///////////////////////////////////////////

Public Property Get Player() As Variant
'////////////////////////////////////////////
'/  Returns the player name Property value
'///////////////////////////////////////////
Player = PlayerName
End Property

Public Property Let Player(ByVal vNewValue As Variant)
'////////////////////////////////////////////
'/  Saves new player name Property value
'///////////////////////////////////////////
PlayerName = vNewValue
End Property



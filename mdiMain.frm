VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Othello"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9840
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStartGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuDif 
      Caption         =   "Difficulty"
      Begin VB.Menu mnuEasy 
         Caption         =   "Easy"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuHard 
         Caption         =   "Hard"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuEasy_Click()
      bDiff = False
      mnuEasy.Checked = True
      mnuHard.Checked = False
End Sub

Private Sub mnuExit_Click()
      Unload mdiMain
End Sub

Private Sub mnuHard_Click()
      bDiff = True
      mnuHard.Checked = True
      mnuEasy.Checked = False
End Sub

Private Sub mnuStartGame_Click()
      ResetBoard
End Sub

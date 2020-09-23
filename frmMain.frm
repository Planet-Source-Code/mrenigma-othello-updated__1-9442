VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBoard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4560
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   11100
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   345
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   5670
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   405
      Index           =   1
      Left            =   10650
      Picture         =   "frmMain.frx":065E
      ScaleHeight     =   345
      ScaleWidth      =   330
      TabIndex        =   0
      Top             =   5670
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3960
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   6985
      _Version        =   393216
      Rows            =   9
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   405
      BackColor       =   32768
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
      AllowUserResizing=   3
      PictureType     =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   "    |^  1|^  2|^  3|^  4|^  5|^  6|^  7|^  8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
      ResetBoard
      Me.Top = 0
      Me.Left = 0
      
End Sub



Private Sub Form_Resize()
      frmMessages.SetSize
End Sub

Private Sub Grid_Click()
      If Grid.Row = 0 Or Grid.Col = 0 Then
         Exit Sub
      End If
      frmMessages.lblError = ""
      If GetPos(Grid.Row, Grid.Col) = False Then
         frmMessages.lblError = "Invalid Move - Try Again"
         Exit Sub
      End If
      TestCounters
      ChangePlayer
      OtherPlayerGo
End Sub

Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim iR As Integer
Dim iC As Integer
Dim iTotalCounters As Integer

      ReDim bMoves(iBoardSize) As Boolean

      iR = Grid.MouseRow
      iC = Grid.MouseCol

      If TestMove(iR, iC, iTotalCounters, bMoves()) Then

         frmMessages.lblHelp = "Valid Move - Counters Possible :" & iTotalCounters
      Else
         frmMessages.lblHelp = "Invalid Move"
      End If
End Sub


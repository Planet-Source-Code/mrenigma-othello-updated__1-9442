VERSION 5.00
Begin VB.Form frmMessages 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Information"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Restart"
      Height          =   495
      Left            =   5730
      TabIndex        =   9
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   525
      Left            =   5760
      TabIndex        =   8
      Top             =   90
      Width           =   1095
   End
   Begin VB.CommandButton cmdNoMove 
      Caption         =   "NoMove"
      Height          =   495
      Left            =   5790
      TabIndex        =   7
      Top             =   780
      Width           =   1095
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label lblSize 
      Caption         =   "Set Board Size"
      Height          =   285
      Left            =   30
      TabIndex        =   6
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblPlayer1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3225
   End
   Begin VB.Label lblPlayer2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   3225
   End
   Begin VB.Label lblError 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   2010
      Width           =   4275
   End
   Begin VB.Label lblHelp 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2670
      Width           =   6555
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSize_Click()
      iBoardSize = cboSize.List(cboSize.ListIndex)
      ResetBoard
End Sub
Private Sub cmdCancel_Click()
      Unload mdiMain
End Sub
Private Sub cmdNoMove_Click()
      ChangePlayer
End Sub

Private Sub Command1_Click()
      Call OtherPlayerGo
End Sub

Private Sub Form_Load()
      cboSize.Clear
      cboSize.AddItem 8
      cboSize.AddItem 10
      cboSize.AddItem 12
      cboSize.AddItem 14
      cboSize.ListIndex = 0
End Sub
Public Sub SetSize()
      Me.Left = frmBoard.Left + frmBoard.Width + 5
      Me.Top = frmBoard.Top
End Sub

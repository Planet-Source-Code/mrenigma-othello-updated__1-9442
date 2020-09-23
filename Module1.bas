Attribute VB_Name = "Module1"
Option Explicit

Public cPlayers(2) As Long
Public iPlayer As Integer
Public iBoardSize As Integer
Public bGameon As Boolean
Public bDiff As Boolean



Private Sub main()
      cPlayers(1) = &HFFFFFF
      cPlayers(2) = &H1&
      Load mdiMain
      Load frmBoard
      Load frmMessages
      mdiMain.Show
      
End Sub


Public Sub ResetBoard()
      bGameon = True
      iPlayer = 2
      iBoardSize = frmMessages.cboSize.Text
      SetupBoard
      TestCounters
      DisplayPlayer
End Sub
Public Sub DisplayPlayer()
      If iPlayer = 2 Then
         frmMessages.lblMessage = "Black Player Move"
      Else
         frmMessages.lblMessage = "White Player Move"
      End If
End Sub
Public Sub TestCounters()
Dim ply1 As Integer
Dim ply2 As Integer
Dim sMessage As String

      Call CountCounters(ply1, ply2)
      frmMessages.lblPlayer1.Caption = "Black Player as " & ply2 & " Counters"
      frmMessages.lblPlayer2.Caption = "White Player as " & ply1 & " Counters"
      If ply1 + ply2 = (iBoardSize * iBoardSize) Or ply1 = 0 Or ply2 = 0 Then
         sMessage = "Game Over - "
         If ply1 > ply2 Then
            sMessage = sMessage & "White Player Wins"
         Else
            sMessage = sMessage & "Black Player Wins"
         End If
         MsgBox sMessage
         bGameon = False
      End If

End Sub
Public Sub ChangePlayer()
      If iPlayer = 1 Then
         iPlayer = 2
      ElseIf iPlayer = 2 Then
         iPlayer = 1
      End If
      DisplayPlayer
End Sub
Public Sub CountCounters(iPlayer1 As Integer, iPlayer2 As Integer)
Dim iRow As Integer
Dim iCol As Integer

      iPlayer1 = 0
      iPlayer2 = 0

      With frmBoard.Grid
         For iCol = 1 To iBoardSize
            For iRow = 1 To iBoardSize
               If .TextMatrix(iRow, iCol) = "1" Then
                  iPlayer1 = iPlayer1 + 1
               End If
               If .TextMatrix(iRow, iCol) = "2" Then
                  iPlayer2 = iPlayer2 + 1
               End If
            Next

         Next
      End With
End Sub

Public Function GetPos(iRow As Integer, iCol As Integer) As Boolean
Dim sOtherPlayer As String
Dim Count As Integer
Dim i As Integer
Dim sCounter As String
Dim sCur As String
Dim iTemp As Integer
Dim iTotalCounters As Integer

      ReDim bMoves(iBoardSize) As Boolean

      GetPos = False
      
      With frmBoard.Grid
         If .TextMatrix(iRow, iCol) <> "" Then
            GetPos = False
            Exit Function
         End If
         If iPlayer = 1 Then
            sOtherPlayer = "2"
         Else
            sOtherPlayer = "1"
         End If
         On Error Resume Next
         
         ' First lets test that move is valid
         
         If TestMove(iRow, iCol, iTotalCounters, bMoves()) Then
         
            ' Change states of pieces
            Debug.Print "Total Counters = " & iTotalCounters
            GetPos = True
            Call SetPos(iPlayer, iRow, iCol)
         
            ' Set the up counters
            
            If bMoves(0) Then
               ' up
               Count = 0
               For i = iRow - 1 To 1 Step -1
                  Count = Count + 1
                  sCur = .TextMatrix(i, iCol)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iCol)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(1) Then
               ' down
               Count = 0
               For i = iRow + 1 To iBoardSize
                  Count = Count + 1
                  sCur = .TextMatrix(i, iCol)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iCol)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(2) Then
               ' left
               Count = 0
               For i = iCol - 1 To 1 Step -1
                  Count = Count + 1
                  sCur = .TextMatrix(iRow, i)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, iRow, i)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(3) Then
               ' right
               Count = 0
               For i = iCol + 1 To iBoardSize
                  Count = Count + 1
                  sCur = .TextMatrix(iRow, i)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, iRow, i)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(4) Then
               ' upleft
               iTemp = iCol
               Count = 0
               For i = iRow - 1 To 1 Step -1
                  Count = Count + 1
                  iTemp = iTemp - 1
                  sCur = .TextMatrix(i, iTemp)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iTemp)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(5) Then
               ' upright
               iTemp = iCol
               Count = 0
               For i = iRow - 1 To 1 Step -1
                  Count = Count + 1
                  iTemp = iTemp + 1
                  sCur = .TextMatrix(i, iTemp)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iTemp)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(6) Then
               ' downleft
               iTemp = iCol
               Count = 0
               For i = iRow + 1 To iBoardSize
                  Count = Count + 1
                  iTemp = iTemp - 1
                  sCur = .TextMatrix(i, iTemp)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iTemp)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
            If bMoves(7) Then
               ' downright
               iTemp = iCol
               Count = 0
               For i = iRow + 1 To iBoardSize
                  Count = Count + 1
                  iTemp = iTemp + 1
                  sCur = .TextMatrix(i, iTemp)
                  If sCur = sOtherPlayer Then
                     Call SetPos(iPlayer, i, iTemp)
                  End If
                  If sCur = CStr(iPlayer) And Count <> 1 Then
                     Exit For
                  End If
               Next
            End If
         End If
      End With

End Function

Public Sub SetPos(iPlayer As Integer, iRow As Integer, iCol As Integer)
      With frmBoard.Grid
         .Col = iCol
         .Row = iRow
         .Text = CStr(iPlayer)
         .CellBackColor = cPlayers(iPlayer)
         .CellForeColor = cPlayers(iPlayer)
         .RowHeightMin = frmBoard.Picture1(iPlayer).Height
         
      End With
End Sub

Public Function TestMove(iRow As Integer, iCol As Integer, ByRef iTotalCounters As Integer, bMoves() As Boolean) As Boolean
Dim sOtherPlayer As String
Dim Count As Integer
Dim i As Integer
Dim sCounter As String
Dim sResult As String
Dim iTemp As Integer

      TestMove = False
      
      iTotalCounters = 0
      
      With frmBoard.Grid
         If .TextMatrix(iRow, iCol) <> "" Then
            Exit Function
         End If
         If iPlayer = 1 Then
            sOtherPlayer = "2"
         Else
            sOtherPlayer = "1"
         End If
         
         On Error Resume Next
         
         ' First lets test that move is valid
10       Count = 0
         ' Is there a valid move Up?
         For i = iRow - 1 To 1 Step -1
            Count = Count + 1
            sCounter = .TextMatrix(i, iCol)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(0) = False
               Else
                  bMoves(0) = True
               End If
               GoTo 20
            End If
            If sCounter = sOtherPlayer Then
               bMoves(0) = True
               GoTo 15
            End If
            If sCounter = "" Then
               bMoves(0) = False
               GoTo 20
            End If
15       Next
         bMoves(0) = False
   
20       If bMoves(0) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0
         ' Is there a valid move Down?
         For i = iRow + 1 To iBoardSize
            Count = Count + 1
            sCounter = .TextMatrix(i, iCol)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(1) = False
               Else
                  bMoves(1) = True
               End If
               GoTo 30
            End If
            If sCounter = sOtherPlayer Then
               bMoves(1) = True
               GoTo 25
            End If
            If sCounter = "" Then
               bMoves(1) = False
               GoTo 30
            End If
25       Next
         bMoves(1) = False
    
30       If bMoves(1) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move left?
         For i = iCol - 1 To 1 Step -1
            Count = Count + 1
            sCounter = .TextMatrix(iRow, i)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(2) = False
               Else
                  bMoves(2) = True
               End If
               GoTo 40
            End If
            If sCounter = sOtherPlayer Then
               bMoves(2) = True
               GoTo 35
            End If
            If sCounter = "" Then
               bMoves(2) = False
               GoTo 40
            End If
35       Next
         bMoves(2) = False
         
40       If bMoves(2) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move right?
         For i = iCol + 1 To iBoardSize
            Count = Count + 1
            sCounter = .TextMatrix(iRow, i)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(3) = False
               Else
                  bMoves(3) = True
               End If
               GoTo 50
            End If
            If sCounter = sOtherPlayer Then
               bMoves(3) = True
               GoTo 45
            End If
            If sCounter = "" Then
               bMoves(3) = False
               GoTo 50
            End If
45       Next
         bMoves(3) = False
             
50       If bMoves(3) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move upleft?
         iTemp = iCol
         For i = iRow - 1 To 1 Step -1
            Count = Count + 1
            iTemp = iTemp - 1
            sCounter = .TextMatrix(i, iTemp)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(4) = False
               Else
                  bMoves(4) = True
               End If
               GoTo 60
            End If
            If sCounter = sOtherPlayer Then
               bMoves(4) = True
               GoTo 55
            End If
            If sCounter = "" Then
               bMoves(4) = False
               GoTo 60
            End If
55       Next
         bMoves(4) = False
         
60       If bMoves(4) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move upright?
         iTemp = iCol
         For i = iRow - 1 To 1 Step -1
            Count = Count + 1
            iTemp = iTemp + 1
            sCounter = .TextMatrix(i, iTemp)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(5) = False
               Else
                  bMoves(5) = True
               End If
               GoTo 70
            End If
            If sCounter = sOtherPlayer Then
               bMoves(5) = True
               GoTo 65
            End If
            If sCounter = "" Then
               bMoves(5) = False
               GoTo 70
            End If
65       Next
         bMoves(5) = False
           
70       If bMoves(5) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move downleft?
         iTemp = iCol
         For i = iRow + 1 To iBoardSize
            Count = Count + 1
            iTemp = iTemp - 1
            sCounter = .TextMatrix(i, iTemp)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(6) = False
               Else
                  bMoves(6) = True
               End If
               GoTo 80
            End If
            If sCounter = sOtherPlayer Then
               bMoves(6) = True
               GoTo 75
            End If
            If sCounter = "" Then
               bMoves(6) = False
               GoTo 80
            End If
75       Next
         bMoves(6) = False
         
80       If bMoves(6) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         Count = 0

         ' Is there a valid move downright?
         iTemp = iCol
         For i = iRow + 1 To iBoardSize
            Count = Count + 1
            iTemp = iTemp + 1
            sCounter = .TextMatrix(i, iTemp)
            If sCounter = CStr(iPlayer) Then
               If Count = 1 Then
                  bMoves(7) = False
               Else
                  bMoves(7) = True
               End If
               GoTo 90
            End If
            If sCounter = sOtherPlayer Then
               bMoves(7) = True
               GoTo 85
            End If
            If sCounter = "" Then
               bMoves(7) = False
               GoTo 90
            End If
85       Next
         bMoves(7) = False
         
90       If bMoves(7) Then
            iTotalCounters = iTotalCounters + Count - 1
         End If
         
         For i = 0 To 7
            If bMoves(i) = True Then
               TestMove = True
               Exit For
            End If
         Next
      End With
      Exit Function
End Function
Public Function SetBack(iRow As Integer, iCol As Integer, bHighlight As Boolean)
      With frmBoard.Grid
         .Row = iRow
         .Col = iCol
         If bHighlight Then
            .CellBackColor = &H80&
         Else
            .CellBackColor = &H8000&
         End If
      End With
End Function
Public Sub SetupBoard()
Dim i As Integer
Dim sTmp As String

      With frmBoard.Grid
         .RowHeightMin = frmBoard.Picture1(iPlayer).Height
         .Clear
         sTmp = "    "
         For i = 1 To iBoardSize
            sTmp = sTmp & "|^  " & i
         Next
         .FormatString = sTmp
         .Rows = iBoardSize + 1
         .Cols = iBoardSize + 1
         For i = 1 To iBoardSize
            .TextMatrix(i, 0) = i & ""
            .TextMatrix(0, i) = i & ""
            .Col = 0
            .Row = i
            .ColWidth(i) = frmBoard.Picture1(1).Width
            .CellBackColor = &H8000000F
         Next
         .Col = 1
         .Row = 1
         .Height = (.CellHeight * (iBoardSize + 2)) + 173
         .Width = (.CellWidth * (iBoardSize + 2)) + 173
         Call SetPos(1, (iBoardSize / 2), (iBoardSize / 2))
         Call SetPos(1, (iBoardSize / 2) + 1, (iBoardSize / 2) + 1)
         Call SetPos(2, (iBoardSize / 2), (iBoardSize / 2) + 1)
         Call SetPos(2, (iBoardSize / 2) + 1, (iBoardSize / 2))
         .ColWidth(0) = frmBoard.Picture1(1).Width
         .Top = 100
         .Left = 100
         frmBoard.Width = .Width
         frmBoard.Height = .Height + 200
      End With
End Sub
Public Sub OtherPlayerGo()
Dim R As Integer
Dim C As Integer
Dim iCounters As Integer
Dim iCount As Integer
Dim sTmp As String
Dim asTmp() As String
Dim asTemp() As String

      If bGameon Then
         ReDim iMoves(1 To iBoardSize, 1 To iBoardSize) As Integer
         ReDim sMoves(0) As String
         ReDim bMoves(iBoardSize) As Boolean
      
         For R = 1 To iBoardSize
            For C = 1 To iBoardSize
               If TestMove(R, C, iCounters, bMoves) Then
                  iMoves(R, C) = iCounters
               End If
            Next
         Next
      
         For R = 1 To iBoardSize
            For C = 1 To iBoardSize
               If iMoves(R, C) > 0 Then
                  ReDim Preserve sMoves(iCount) As String
                  sMoves(iCount) = iMoves(R, C) & "," & R & "," & C
                  ' SetBack R, C, True
                  iCount = iCount + 1
                  ' SetBack R, C, False
               End If
            Next
         Next
         If iCount > 2 Then
            For iCount = 1 To UBound(sMoves)
redo:
               If Val(Left$(sMoves(iCount), InStr(1, sMoves(iCount), ","))) > Val(Left$(sMoves(iCount - 1), InStr(1, sMoves(iCount - 1), ","))) Then
                  ' Sort
                  sTmp = sMoves(iCount)
                  sMoves(iCount) = sMoves(iCount - 1)
                  sMoves(iCount - 1) = sTmp
                  If iCount > 1 Then
                     iCount = iCount - 1
                  End If
                  GoTo redo
               End If
            Next
         End If
                  
         If UBound(sMoves) >= 0 And sMoves(0) <> "" Then
            asTmp = Split(sMoves(0), ",")
         
            iCount = 0
            For R = 1 To UBound(sMoves)
               asTemp = Split(sMoves(R), ",")
               If asTemp(0) = asTmp(0) Then
                  iCount = iCount + 1
               Else
                  Exit For
               End If
            Next
            
            If bDiff Then
               ' Difficulty is hard
            Else
               ' Difficulty is easy
               iCount = UBound(sMoves)
            End If
            
            If iCount > 0 Then
               ' found more than 1 of best number of counters
               ' So left pick a random one
               R = Int(Rnd * (iCount + 1))
               asTmp = Split(sMoves(R), ",")
            End If
         
            If TestMove(CInt(asTmp(1)), CInt(asTmp(2)), iCounters, bMoves) Then
               GetPos CInt(asTmp(1)), CInt(asTmp(2))
            End If
         Else
            frmMessages.lblError = "Could Not Move"
         End If
         TestCounters
         ChangePlayer
      End If
End Sub


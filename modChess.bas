Attribute VB_Name = "modChess"
Option Explicit
Option Base 0

Public Const EMPTY_SQUARE As Byte = 0
Public Const PAWN As Byte = 1
Public Const ROOK As Byte = 2
Public Const KNIGHT As Byte = 4
Public Const BISHOP As Byte = 8
Public Const QUEEN As Byte = 16
Public Const KING As Byte = 32
Public Const WHITE_PEICE As Byte = 0
Public Const BLACK_PEICE As Byte = 128

Private Const AI_IN As Double = 1 / 255
Private Const AI_OUT As Double = 1 / 64

Public CurTurn As Byte
Public AI As cNN 'v1.0.2 used
Public MoveList() As Variant 'Array of moves, each move is Peice, OriginX, OriginY, DestX, DestY
Public Board(1 To 8, 1 To 8) As Byte '1 -- 8, A -- H
            '
            'H |_|#|_|#|_|#|_|#|
            'G |#|_|#|_|#|_|#|_|
            'F |_|#|_|#|_|#|_|#|
            'E |#|_|#|_|#|_|#|_|
            'D |_|#|_|#|_|#|_|#|
            'C |#|_|#|_|#|_|#|_|
            'B |_|#|_|#|_|#|_|#|
            'A |#|_|#|_|#|_|#|_|
            '   1 2 3 4 5 6 7 8
            

Public Sub Main()
  Debug.Print "Startup: ", Now
  Set AI = New cNN
  AI.InitializeNN Array(65, 504, 128, 48, 2)
  InitializeBoard
  Debug.Print "Init Done: ", Now
  Do
    Call AI_Move
    'If (CheckMate) Then InitalizeBoard
  Loop
End Sub

Public Sub InitializeBoard()
  'Okay, okay, so the way I do this is not necessary, but it was fun and it's interesting!
  Dim i As Long, j As Long, k As Byte
  
  CurTurn = WHITE_PEICE
  ReDim MoveList(0)
  MoveList(0) = Array(CByte(0), 0, 0, 0, 0)
  
  'Clear center
  For i = 1 To 8
    For j = 3 To 6
      Board(i, j) = EMPTY_SQUARE
    Next
  Next
  
  'Place Pawns
  k = 0
  For i = 1 To 8
    For j = 2 To 7 Step 5
      Board(i, j) = PAWN + k
      k = k Xor BLACK_PEICE
    Next
  Next
  
  'Place Rooks
  k = 0
  For i = 1 To 8 Step 7
    For j = 1 To 8 Step 7
      Board(i, j) = ROOK + k
      k = k Xor BLACK_PEICE
    Next
  Next
  
  'Place Knights
  k = 0
  For i = 2 To 7 Step 5
    For j = 1 To 8 Step 7
      Board(i, j) = KNIGHT + k
      k = k Xor BLACK_PEICE
    Next
  Next
  
  'Place Bishops
  k = 0
  For i = 3 To 6 Step 3
    For j = 1 To 8 Step 7
      Board(i, j) = BISHOP + k
      k = k Xor BLACK_PEICE
    Next
  Next
  
  'Place Queens
  k = 0
  i = 4
  For j = 1 To 8 Step 7
    Board(i, j) = QUEEN + k
    k = k Xor BLACK_PEICE
  Next
  
  'Place Kings
  k = 0
  i = 5
  For j = 1 To 8 Step 7
    Board(i, j) = KING + k
    k = k Xor BLACK_PEICE
  Next
End Sub

Private Function EvalSquare(Peice As Byte, DestX, DestY) As Long
  'Determins if peice can be placed in destination
  'Returns: 0 for square occupied by one of your peices, -1 for empty square, and -2 for enemy in square
  If (Board(DestX, DestY) = 0) Then EvalSquare = -1 Else If ((Peice And BLACK_PEICE) Xor (Board(DestX, DestY) And BLACK_PEICE)) Then EvalSquare = -2 Else EvalSquare = 0
End Function

Private Function EvalPath(Peice As Byte, OriginX, OriginY, DestX, DestY) As Long
  'Determins if path of peice movement is clear
  'Returns: 0 for path occupied any peice, -1 for empty path+destination, and -2 for enemy in destination square
  Dim a As Long, b As Long
  Dim i As Long, j As Long
  
  If _
    ((OriginX = DestX) And (OriginY <> DestY)) _
    Or _
    ((OriginX <> DestX) And (OriginY = DestY)) _
  Then
    'Check squares along Horzontal and Vertical paths
    b = 0
    For i = OriginX To DestX Step ((CInt(OriginX > DestX) * 2) + 1)
      For j = OriginY To DestY Step ((CInt(OriginY > DestY) * 2) + 1)
        If (Not ((i = OriginX) And (j = OriginY))) Then
          a = EvalSquare(Peice, i, j)
          If (i = DestX) And (j = DestY) Then
            b = a
          ElseIf (a < -1) Then
            a = 0
          End If
          If Not a Then
            EvalPath = 0
            Exit Function
          End If
        End If
      Next
    Next
    EvalPath = b
  ElseIf _
    ( _
      ((OriginX <> DestX) And (OriginY <> DestY)) _
      And _
      ((Abs(OriginX - DestX) - Abs(OriginY - DestY)) = 0) _
    ) _
  Then
    'Check squares along diagonal paths
    b = 0
    For i = OriginX To DestX Step ((CInt(OriginX > DestX) * 2) + 1)
      For j = OriginY To DestY Step ((CInt(OriginY > DestY) * 2) + 1)
        If ((Abs(i) - Abs(j)) = 0) Then
          If (Not ((OriginX = DestX) And (OriginY = DestY)) And (Not ((i = OriginX) And (j = OriginY)))) Then
            a = EvalSquare(Peice, i, j)
            If (i = DestX) And (j = DestY) Then
              b = a
            ElseIf (a < -1) Then
              a = 0
            End If
            If Not a Then
              EvalPath = 0
              Exit Function
            End If
          End If
        End If
      Next
    Next
    EvalPath = b
  Else
    'Not a horziontal, vertical, or diagonal path
    EvalPath = 0
  End If
End Function

Public Function EvalMove(Peice As Byte, OriginX, OriginY, DestX, DestY) As Long
  'Checks if peice movement is valid
  'Returns: 0 for path occupied any peice, -1 for empty path+destination, and -2 for enemy in destination square
  Dim a As Byte, b As Long
  a = Peice And BLACK_PEICE
  If _
    ((OriginX = DestX) And (OriginY = DestY)) _
    Or _
    (a <> CurTurn) _
  Then
    EvalMove = 0 'No movement, just same square or wrong color peice
  Else
    Select Case (Peice And Not BLACK_PEICE) 'Remove color from peice
      Case PAWN
          'Peice movement one space in proper direction
          'OR
          'Peice movement from starting row 2 spaces in proper direction with no sideways movement
          b = (Abs(OriginX - DestX) < 2) _
              And _
              ( _
                (OriginY - DestY = (1 + ((a <> BLACK_PEICE) * 2))) _
                Or _
                ( _
                  (OriginY - DestY = (2 + ((a <> BLACK_PEICE) * 4))) _
                  And _
                  (OriginY = 2 + Abs(a = BLACK_PEICE) * 5) _
                  And _
                  (OriginX = DestX) _
                ) _
              )
                
          'b
          'AND
          '  No sideways movement
          '  OR
          '    Attack enemy peice
          '    OR
          '    En Passant
          If _
            b _
            And _
            ( _
              (OriginX = DestX) _
              Or _
              ( _
                (EvalSquare(Peice, DestX, DestY) = -2) _
                Or _
                ( _
                  (EvalSquare(Peice, DestX, DestY) = -1) _
                  And _
                  (MoveList(UBound(MoveList))(1) = DestX) _
                  And _
                  (MoveList(UBound(MoveList))(2) = DestY + (((a <> BLACK_PEICE) * 2) + 1))) _
                  And _
                  ((MoveList(UBound(MoveList))(0) And Not BLACK_PEICE) = PAWN) _
                  And _
                  (((MoveList(UBound(MoveList))(0) And Not PAWN) Xor a) = BLACK_PEICE) _
                ) _
              ) _
          Then
            EvalMove = EvalSquare(Peice, DestX, DestY)
          Else
            EvalMove = 0
          End If
      Case ROOK
          'Only horizontal and vertical moves
          If ((OriginX = DestX) And (OriginY <> DestY)) Or ((OriginX <> DestX) And (OriginY = DestY)) Then
            EvalMove = EvalPath(Peice, OriginX, OriginY, DestX, DestY)
          Else
            EvalMove = 0
          End If
      Case KNIGHT
          'Only hops 2 out and 1 to the side
          If _
            (Abs(OriginX - DestX) = 1) And (Abs(OriginY - DestY) = 2) _
            Or _
            (Abs(OriginX - DestX) = 2) And (Abs(OriginY - DestY) = 1) _
          Then
            EvalMove = EvalSquare(Peice, DestX, DestY)
          Else
            EvalMove = 0
          End If
      Case BISHOP
          'Only diagonal moves
          If ((OriginX <> DestX) And (OriginY <> DestY)) Then
            EvalMove = EvalPath(Peice, OriginX, OriginY, DestX, DestY)
          Else
            EvalMove = 0
          End If
      Case QUEEN
          EvalMove = EvalPath(Peice, OriginX, OriginY, DestX, DestY)
      Case KING
          If _
            (CheckSquare(Peice, DestX, DestY) = False) _
            And _
            ( _
              (Abs(OriginX - DestX) = 1) And (Abs(OriginY - DestY) = 1) _
              Or _
              ( _
                (Abs(OriginX - DestX) = 2) And (OriginY = DestY) _
                And _
                (OriginY = (((a = BLACK_PEICE) * 7) + 8)) _
                And _
                (OriginX = 5) _
                And _
                (CheckSquare(Peice, OriginX + ((OriginX > DestX) * 2 + 1), DestY) = False) _
              ) _
            ) _
          Then
            EvalMove = EvalPath(Peice, OriginX, OriginY, DestX, DestY)
          Else
            EvalMove = 0
          End If
    End Select
  End If
End Function

Private Function CheckSquare(Peice As Byte, DestX, DestY) As Boolean
  'Checks if a square can be attacked by another peice (ie: a king moving into a "Checked" square)
  'Returns: True if square can be attacked
  Dim a As Long
  Dim i As Long, j As Long
  
  For i = 1 To 8
    For j = 1 To 8
      'An Enemy peice
      'And
      'Ignore the square in question
      If _
        ((Board(i, j) And BLACK_PEICE) <> (Peice And BLACK_PEICE)) _
        And _
        Not ((i = DestX) And (j = DestY)) _
      Then
        'enemy king within one space of square is check
        '(Don't use EvalMove for king distance because of infinate stack calling)
        If ((Board(i, j) And Not BLACK_PEICE) = KING) Then
          If ((Abs(DestX - i) = 1) And (Abs(DestY - j) = 1)) Then a = True
        Else
          a = EvalMove(Board(i, j), i, j, DestX, DestY)
        End If
        If a Then
          CheckSquare = True
          Exit Function
        End If
      End If
    Next
  Next
End Function

Public Sub MakeMove(Peice As Byte, OriginX, OriginY, DestX, DestY)
  'Record the move
  ReDim Preserve MoveList(1 To UBound(MoveList) + 1)
  MoveList(UBound(MoveList)) = Array(Peice, CLng(OriginX), CLng(OriginY), CLng(DestX), CLng(DestY))
  
  'Update Board
  Board(DestX, DestY) = Board(OriginX, OriginY)
  Board(OriginX, OriginY) = EMPTY_SQUARE
  
  'Update Meshes
  
  
  'Change CurTurn
  If (CurTurn = WHITE_PEICE) Then CurTurn = BLACK_PEICE Else CurTurn = whitepeice
End Sub

Public Sub AI_Train()
  'Teach the AI all the possible moves for the current board state
  Dim i As Long, j As Long, k As Long, l As Long, f As Long
  
  f = 0
  For i = 1 To 8
    For j = 1 To 8
      If ((Board(i, j) And BLACK_PEICE) = (Board(i, j) And CurTurn)) Then
        For k = 1 To 8
          For l = 1 To 8
            If (EvalMove(Board(i, j), i, j, k, l)) Then
              f = f + 1
              AI.Refresh
              DoEvents
              AI.Train Array((i + j * 8) * AI_OUT, (k + l * 8) * AI_OUT)
              AI.ExportNN App.Path & "\AI.net"
              DoEvents
              Debug.Print "Training Progress... " & Now, f & " (" & i & "," & j & ")(" & k & "," & l & ")", CInt(AI.GetOutput(1) / AI_OUT), CInt(AI.GetOutput(2) / AI_OUT)
            End If
          Next
        Next
      End If
    Next
  Next
End Sub
    
Sub AI_Move()
  Dim a As Variant, b As Variant, c As Variant
  Dim i As Long, j As Long
  
  c = Array(AI_IN * CDbl(Board(1, 1)), AI_IN * CDbl(Board(2, 1)), AI_IN * CDbl(Board(3, 1)), AI_IN * CDbl(Board(4, 1)), AI_IN * CDbl(Board(5, 1)), AI_IN * CDbl(Board(6, 1)), AI_IN * CDbl(Board(7, 1)), AI_IN * CDbl(Board(8, 1)), _
            AI_IN * CDbl(Board(1, 2)), AI_IN * CDbl(Board(2, 2)), AI_IN * CDbl(Board(3, 2)), AI_IN * CDbl(Board(4, 2)), AI_IN * CDbl(Board(5, 2)), AI_IN * CDbl(Board(6, 2)), AI_IN * CDbl(Board(7, 2)), AI_IN * CDbl(Board(8, 2)), _
            AI_IN * CDbl(Board(1, 3)), AI_IN * CDbl(Board(2, 3)), AI_IN * CDbl(Board(3, 3)), AI_IN * CDbl(Board(4, 3)), AI_IN * CDbl(Board(5, 3)), AI_IN * CDbl(Board(6, 3)), AI_IN * CDbl(Board(7, 3)), AI_IN * CDbl(Board(8, 3)), _
            AI_IN * CDbl(Board(1, 4)), AI_IN * CDbl(Board(2, 4)), AI_IN * CDbl(Board(3, 4)), AI_IN * CDbl(Board(4, 4)), AI_IN * CDbl(Board(5, 4)), AI_IN * CDbl(Board(6, 4)), AI_IN * CDbl(Board(7, 4)), AI_IN * CDbl(Board(8, 4)), _
            AI_IN * CDbl(Board(1, 5)), AI_IN * CDbl(Board(2, 5)), AI_IN * CDbl(Board(3, 5)), AI_IN * CDbl(Board(4, 5)), AI_IN * CDbl(Board(5, 5)), AI_IN * CDbl(Board(6, 5)), AI_IN * CDbl(Board(7, 5)), AI_IN * CDbl(Board(8, 5)), _
            AI_IN * CDbl(Board(1, 6)), AI_IN * CDbl(Board(2, 6)), AI_IN * CDbl(Board(3, 6)), AI_IN * CDbl(Board(4, 6)), AI_IN * CDbl(Board(5, 6)), AI_IN * CDbl(Board(6, 6)), AI_IN * CDbl(Board(7, 6)), AI_IN * CDbl(Board(8, 6)), _
            AI_IN * CDbl(Board(1, 7)), AI_IN * CDbl(Board(2, 7)), AI_IN * CDbl(Board(3, 7)), AI_IN * CDbl(Board(4, 7)), AI_IN * CDbl(Board(5, 7)), AI_IN * CDbl(Board(6, 7)), AI_IN * CDbl(Board(7, 7)), AI_IN * CDbl(Board(8, 7)), _
            AI_IN * CDbl(Board(1, 8)), AI_IN * CDbl(Board(2, 8)), AI_IN * CDbl(Board(3, 8)), AI_IN * CDbl(Board(4, 8)), AI_IN * CDbl(Board(5, 8)), AI_IN * CDbl(Board(6, 8)), AI_IN * CDbl(Board(7, 8)), AI_IN * CDbl(Board(8, 8)), _
            Abs(CurTurn = BLACK_PEICE) _
          )
  AI.SetInput c
  Do
    AI.Refresh
    DoEvents
    a = CInt(AI.GetOutput(1) / AI_OUT)
    b = CInt(AI.GetOutput(2) / AI_OUT)
    If (a = 0) Then a = 1
    If (b = 0) Then b = 1
    For i = 1 To 8
      For j = 1 To 8
        If (Not IsArray(a)) Then If (i + ((j - 1) * 8) = a) Then a = Array(i, j)
        If (Not IsArray(b)) Then If (i + ((j - 1) * 8) = b) Then b = Array(i, j)
      Next
      If (IsArray(a) And IsArray(b)) Then Exit For
    Next
    If Not (IsArray(a) And IsArray(b)) Then
      Debug.Print "Move attempt failed... Begining Training."
      AI_Train
    ElseIf (EvalMove(Board(a(0), a(1)), a(0), a(1), b(0), b(1))) Then
      MakeMove Board(a(0), a(1)), a(0), a(1), b(0), b(1)
      Debug.Print ListMove(UBound(MoveList))
      Exit Do
    Else
      Debug.Print "Move attempt failed... Begining Training."
      AI_Train
    End If
  Loop
End Sub

Public Function ListMove(MoveNumber) As String
  Dim a As Variant
  
  a = Peice And BLACK_PEICE
  Select Case Peice And Not a
    Case PAWN:   ListMove = "P"
    Case ROOK:   ListMove = "R"
    Case KNIGHT: ListMove = "N"
    Case BISHOP: ListMove = "B"
    Case QUEEN:  ListMove = "Q"
    Case KING:   ListMove = "K"
  End Select
  If (a = BLACK_PEICE) Then ListMove = "Blk: " & ListMove Else ListMove = "Wht: " & ListMove
  a = Array("A", "B", "C", "D", "E", "F", "G", "H")
  ListMove = ListMove & MoveList(MoveNumber)(1) & a(MoveList(MoveNumber)(2)) & ":" & MoveList(MoveNumber)(3) & a(MoveList(MoveNumber)(4))
End Function

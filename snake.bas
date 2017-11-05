Public power As Boolean
Public dir As Integer
Public speed As Integer
Public play As Boolean
Public score As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + DurationMS
    
    Do While EndTime > GetTickCount
        DoEvents
        Sleep 1
    Loop
End Sub

Sub MoveUp()
    dir = 1
End Sub

Sub MoveDown()
    dir = 2
End Sub

Sub MoveLeft()
    dir = 3
End Sub

Sub MoveRight()
    dir = 4
End Sub

Sub ClearField()
    Range("C3:J10").Clear
    With Range("B2:K11")
        .Interior.ColorIndex = 0
        .Value = vbNullString
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub CheckTable()
    With Worksheets("Table")
        If .Range("B11").Value < score Then
            .Range("B11").Value = score
            .Range("C11").Value = InputBox("Name:", "New Record: " & score, "")
            .Range("B2:C11").Sort key1:=.Range("B2:B11"), order1:=xlDescending
            
            Range("B2:K11").Interior.ColorIndex = 1
            ActiveWorkbook.Save
        End If
    End With
End Sub

Sub KillSnake(snake() As Integer)
    Wait 50
    
    play = False
    score = score * speed
    
    CheckTable
    InitializeConsoleMenu
End Sub

Function CheckSnake(snake() As Integer) As Boolean
    CheckSnake = False
    
    For I = 0 To UBound(snake)
        For j = I + 1 To UBound(snake)
            If (snake(I, 0) = snake(j, 0)) And (snake(I, 1) = snake(j, 1)) Then
                KillSnake snake
                Exit Function
            End If
        Next j
    Next I
    
    If Application.Intersect(Range("B2:K11"), Cells(snake(0, 0), snake(0, 1))) Is Nothing Then
        KillSnake snake
    Else
        CheckSnake = True
    End If
End Function

Sub IncreaseLength(snake() As Integer, tail() As Integer)
    Dim newSnake() As Integer
    ReDim newSnake(UBound(snake) + 1, 1)
    
    For I = 0 To UBound(snake)
        newSnake(I, 0) = snake(I, 0)
        newSnake(I, 1) = snake(I, 1)
    Next I
    newSnake(UBound(newSnake), 0) = tail(0, 0)
    newSnake(UBound(newSnake), 1) = tail(0, 1)
    
    snake = newSnake
End Sub

Sub GenerateFood()
    Dim randomCell As Long
    randomCell = Int(Rnd * Range("B2:K11").Cells.Count) + 1
    
    Do While Range("B2:K11").Cells(randomCell).Interior.ColorIndex > 0
        randomCell = Int(Rnd * Range("B2:K11").Cells.Count) + 1
    Loop
    
    Range("B2:K11").Cells(randomCell).Value = "o"
End Sub

Sub DyeSnake(snake() As Integer)
    If Not Application.Intersect(Range("B2:K11"), Cells(snake(0, 0), snake(0, 1))) Is Nothing Then
        Cells(snake(0, 0), snake(0, 1)).Interior.ColorIndex = 1
    End If
    
    For I = 1 To UBound(snake)
        Cells(snake(I, 0), snake(I, 1)).Interior.ColorIndex = 10
    Next I
End Sub

Sub MakeStep(snake() As Integer)
    Dim tail(0, 1) As Integer
    tail(0, 0) = snake(UBound(snake), 0)
    tail(0, 1) = snake(UBound(snake), 1)
    
    For I = UBound(snake) To 1 Step -1
        snake(I, 0) = snake(I - 1, 0)
        snake(I, 1) = snake(I - 1, 1)
    Next I
    
    If Cells(snake(0, 0), snake(0, 1)).Value <> vbNullString Then
        IncreaseLength snake, tail
        GenerateFood
        Cells(snake(0, 0), snake(0, 1)).Value = vbNullString
    Else
        Cells(tail(0, 0), tail(0, 1)).Interior.ColorIndex = 0
    End If
    
    If dir = 1 Then
        snake(0, 0) = snake(0, 0) - 1
    Else
        If dir = 2 Then
            snake(0, 0) = snake(0, 0) + 1
        Else
            If dir = 3 Then
                snake(0, 1) = snake(0, 1) - 1
            Else
                If dir = 4 Then
                    snake(0, 1) = snake(0, 1) + 1
                End If
            End If
        End If
    End If
    
    DyeSnake snake
    score = score + UBound(snake)
End Sub

Sub InitializeGame(snake() As Integer)
    ClearField
    dir = 4
    play = True
    score = 0
    
    ReDim snake(2, 1)
    snake(0, 0) = 10
    snake(0, 1) = 5
    snake(1, 0) = 10
    snake(1, 1) = 4
    snake(2, 0) = 10
    snake(2, 1) = 3
    
    DyeSnake snake
    GenerateFood
    Wait 500
End Sub

Sub Start()
    Dim snake() As Integer
    InitializeGame snake
    
    Do While power And play
        MakeStep snake
        If CheckSnake(snake) = False Then
            Exit Do
        End If
        
        Wait 800 / 2 ^ (speed - 1)
    Loop
End Sub

Sub InitializeConsoleMenu()
    ClearField
    Range("E5:H8").BorderAround xlContinuous
    
    If score <> 0 Then
        With Range("E3:H3")
            .Merge
            .Borders(xlEdgeTop).LineStyle = xlDouble
            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Value = "SCORE: " & score
            .Font.Name = "Consolas"
            .Font.Size = 12
            .Font.Bold = True
        End With
    End If
    
    With Range("E5:H6")
        .Merge
        .Value = "START"
        .Font.Name = "Consolas"
        .Font.Size = 20
    End With
    With Range("E7:H8")
        .Merge
        .Value = "SPEED"
        .Font.Name = "Consolas"
        .Font.Size = 20
        .Interior.ColorIndex = 17
    End With
    With Range("D7:D8")
        .Merge
        .Value = "<"
        .Font.Name = "Consolas"
        .Font.Size = 20
    End With
    With Range("I7:I8")
        .Merge
        .Value = ">"
        .Font.Name = "Consolas"
        .Font.Size = 20
    End With
    With Range("F9:G10")
        .Merge
        .Value = speed
        .Font.Name = "Consolas"
        .Font.Size = 20
    End With
End Sub

Sub ConsolePower()
    dir = 0
    speed = 1
    play = False
    score = 0
    
    If power = False Then
        InitializeConsoleMenu
        power = True
    Else
        ClearField
        
        With Range("B2:K11")
            .Clear
            .BorderAround xlContinuous, xlMedium
            .Interior.ColorIndex = 1
        End With
        
        power = False
    End If
    
    Do While power = True
        If dir = 3 And speed > 1 Then
            speed = speed - 1
            Range("F9:G10").Value = speed
        Else
            If dir = 4 And speed < 5 Then
                speed = speed + 1
                Range("F9:G10").Value = speed
            Else
                If dir = 1 Then
                    Range("E5:H6").Interior.ColorIndex = 17
                    Range("E7:H8").Interior.ColorIndex = 2
                    
                    Wait 50
                    Start
                End If
            End If
        End If
        dir = 0
        
        DoEvents
        Sleep 1
    Loop
End Sub
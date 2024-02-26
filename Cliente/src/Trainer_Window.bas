Attribute VB_Name = "Trainer_Window"
Public Sub DrawTrainer()
Dim i As Long, YPos As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Name
        YPos = .Y + 70
        RenderText Font_Default, Trim$(Player(MyIndex).Name), .X + 96, YPos, White
        
        '//Level
        YPos = YPos + 23
        RenderText Font_Default, "Lv " & Trim$(Player(MyIndex).Level), .X + 96 + (15 - (GetTextWidth(Font_Default, "Lv " & Trim$(Player(MyIndex).Level)) / 2)), YPos, White
        
        '//Exp
        YPos = YPos + 23
        RenderText Font_Default, Player(MyIndex).CurExp & "/" & GetLevelNextExp(Player(MyIndex).Level), .X + 96, YPos, White
        
        '//Money
        YPos = YPos + 23
        RenderTexture Tex_Item(IDMoney), .X + 60, YPos, 0, 0, 17, 17, 24, 24
        RenderText Font_Default, (FormatarValor(Player(MyIndex).Money)), .X + 96, YPos, White
        
        '//Cash
        YPos = YPos + 23
        RenderTexture Tex_Item(IDCash), .X + 60, YPos, 0, 0, 17, 17, 24, 24
        RenderText Font_Default, (FormatarValor(Player(MyIndex).Cash)), .X + 96, YPos, White
        
        '//Jornada Iniciada
        YPos = YPos + 23
        RenderText Font_Default, (Player(MyIndex).Started), .X + 96, YPos, White
        
        '//Tempo Jogado
        YPos = YPos + 23
        RenderText Font_Default, SecondsToHMS(Player(MyIndex).TimePlay), .X + 96, YPos, White
        
        '//PvP
        YPos = YPos + 53
        RenderText Font_Default, (Player(MyIndex).win), .X + 96, YPos, White
        
        YPos = YPos + 24
        RenderText Font_Default, (Player(MyIndex).Lose), .X + 96, YPos, White
        
        YPos = YPos + 24
        RenderText Font_Default, (Player(MyIndex).Tie), .X + 96, YPos, White
        
    End With
End Sub

Function FormatarValor(ByVal valor As Long) As String
    If Len(CStr(valor)) > 3 Then
        FormatarValor = Format(valor, "0,###")
    Else
        FormatarValor = CStr(valor)
    End If
End Function

' ***************
' ** Trainer **
' ***************
Public Sub TrainerMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_TRAINER
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub TrainerMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRAINER Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Public Sub TrainerMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRAINER Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Trainer_Close
                                If GUI(GuiEnum.GUI_TRAINER).Visible = True Then
                                    GuiState GUI_TRAINER, False
                                End If
                            Case ButtonEnum.Trainer_Badge
                                If GUI(GuiEnum.GUI_BADGE).Visible = False Then
                                    GuiState GUI_BADGE, True
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

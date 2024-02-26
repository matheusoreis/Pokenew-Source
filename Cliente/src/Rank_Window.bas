Attribute VB_Name = "Rank_Window"
Public Sub DrawRank()
    Dim i As Long
    Dim rankIndex As Long
    Dim DrawX As Integer, DrawY As Integer

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Buttons
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollDown
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        '//Scroll
        If RankingHighIndex > RankingScrollViewLine Then
            RenderTexture Tex_Gui(.Pic), .X + 227, .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY), 260, 60, 19, 35, 19, 35
        End If
        
        '//ShowRank
        For i = RankingViewCount To (RankingViewCount) + (RankingScrollViewLine - 1)
            If i >= 0 And i < RankingHighIndex Then
                rankIndex = i + 1
                
                DrawX = .X + 8
                DrawY = .Y + (40 + (31 * ((((rankIndex) - (RankingViewCount)) - 1))))
                
                Select Case rankIndex
                    Case 1
                        RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 7, 316, 213, 28, 213, 28
                    Case 2
                        RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 7, 344, 213, 28, 213, 28
                    Case 3
                        RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 7, 372, 213, 28, 213, 28
                    Case Else
                        RenderTexture Tex_Gui(.Pic), DrawX, DrawY, 7, 400, 213, 28, 213, 28
                End Select
                
                RenderText Font_Default, rankIndex & ": " & Trim$(Rank(rankIndex).Name) & " Lv" & Rank(rankIndex).Level, DrawX + 5, DrawY + 3, White

                'If rankIndex >= 1 And rankIndex <= 3 Then
                '    RenderTexture Tex_Item(528 - 1 + rankIndex), DrawX, DrawY, 0, 0, 24, 24, 24, 24
                '    RenderText Font_Default, Trim$(Rank(rankIndex).Name) & " Lv" & Rank(rankIndex).Level, DrawX + 15, DrawY, Dark
                'Else
                '    RenderText Font_Default, rankIndex & ": " & Trim$(Rank(rankIndex).Name) & " Lv" & Rank(rankIndex).Level, DrawX + 5, DrawY, Dark
                'End If
            End If
        Next
    End With
End Sub

' **********
' ** Rank **
' **********
Public Sub RankMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim CanHover As Boolean, MoveNum As Long, MN As Long
    Dim x2 As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Set to top most
        UpdateGuiOrder GUI_RANK

        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick

                        Select Case i
                        Case ButtonEnum.Rank_ScrollUp
                            RankingScrollUp = True
                            RankingScrollDown = False
                            RankingScrollTimer = GetTickCount
                        Case ButtonEnum.Rank_ScrollDown
                            RankingScrollUp = False
                            RankingScrollDown = True
                            RankingScrollTimer = GetTickCount
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for scroll
        If CursorX >= .X + 227 And CursorX <= .X + 227 + 19 And CursorY >= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) And CursorY <= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize Then
            RankingScrollHold = True
        End If

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub RankMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpX As Long, tmpY As Long
    Dim i As Long
    Dim CanHover As Boolean, MoveNum As Long, MN As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RANK Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover

                        IsHovering = True
                        MouseIcon = 1    '//Select
                    End If
                End If
            End If
        Next

        '//Check for scroll
        If RankingHighIndex > RankingScrollViewLine Then
            If CursorX >= .X + 227 And CursorX <= .X + 227 + 19 And CursorY >= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) And CursorY <= .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize Then
                IsHovering = True
                MouseIcon = 1    '//Select
            End If

            '//Scroll moving
            If RankingScrollHold Then
                '//Upward
                If CursorY < .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + (RankingScrollSize / 2) Then
                    If RankingScrollY < RankingScrollEndY - RankingScrollSize Then
                        RankingScrollY = (CursorY - (.Y + RankingScrollStartY + (RankingScrollEndY - RankingScrollSize)) - (RankingScrollSize / 2)) * -1
                        If RankingScrollY >= RankingScrollEndY - RankingScrollSize Then RankingScrollY = RankingScrollEndY - RankingScrollSize
                    End If
                End If
                '//Downward
                If CursorY > .Y + RankingScrollStartY + ((RankingScrollEndY - RankingScrollSize) - RankingScrollY) + RankingScrollSize - (RankingScrollSize / 2) Then
                    If RankingScrollY > 0 Then
                        RankingScrollY = (CursorY - (.Y + RankingScrollStartY + (RankingScrollEndY - RankingScrollSize)) - RankingScrollSize + (RankingScrollSize / 2)) * -1
                        If RankingScrollY <= 0 Then RankingScrollY = 0
                    End If
                End If

                RankingScrollCount = (RankingScrollLength - RankingScrollY)
                RankingViewCount = ((RankingScrollCount / RankingMaxViewLine) / (RankingScrollLength / RankingMaxViewLine)) * RankingMaxViewLine
            End If
        End If

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

Public Sub RankMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RANK Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Rank_Close
                                If GUI(GuiEnum.GUI_RANK).Visible Then
                                    GuiState GUI_RANK, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Ranking Scroll
        RankingScrollHold = False

        '//Check for dragging
        .InDrag = False
    End With
End Sub

Attribute VB_Name = "GlobalMenu_Window"
Public Sub DrawGlobalMenu()
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(200, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                Select Case i
                    Case 13
                        ' Desenha o botão de retornar
                        RenderText Font_Default, TextUIGlobalMenuReturn, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIGlobalMenuReturn) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                    Case 14
                        ' Desenha o botão de opções
                        RenderText Font_Default, TextUIGlobalMenuOptions, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIGlobalMenuOptions) / 2) - 4, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                    Case 15
                        ' Desenha o botão de voltar o menu principal
                        RenderText Font_Default, TextUIGlobalMenuReturnMenu, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIGlobalMenuReturnMenu) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                    Case 16
                        ' Desenha o botão de sair
                        RenderText Font_Default, TextUIGlobalMenuExit, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIGlobalMenuExit) / 2) - 4, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                End Select
                
            End If
        Next
    End With
End Sub

' ***************
' ** GlobalMenu **
' ***************
Public Sub GlobalMenuMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_GLOBALMENU
        
        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub GlobalMenuMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
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
    End With
End Sub

Public Sub GlobalMenuMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.GlobalMenu_Return
                                GuiState GUI_GLOBALMENU, False
                            Case ButtonEnum.GlobalMenu_Option
                                GuiState GUI_GLOBALMENU, False
                                GuiState GUI_OPTION, True
                                InitSettingConfiguration
                            Case ButtonEnum.GlobalMenu_Back
                                If GameState = GameStateEnum.InMenu Then
                                    GuiState GUI_GLOBALMENU, False
                                ElseIf GameState = GameStateEnum.InGame Then
                                    GuiState GUI_GLOBALMENU, False
                                    OpenChoiceBox TextUIChoiceReturnMainMenu, CB_RETURNMENU
                                End If
                            Case ButtonEnum.GlobalMenu_Exit
                                GuiState GUI_GLOBALMENU, False
                                OpenChoiceBox TextUIChoiceExit, CB_EXIT
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

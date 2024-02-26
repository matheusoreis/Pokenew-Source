Attribute VB_Name = "Characters_Windows"
' Padding lateral
Dim PaddingLeft As Long

' Padding da barra superior
Dim PaddingTop As Long

Public Sub DrawCharacterSelect()
Dim i As Long
Dim CharNameText As String
Dim Sprite As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
    
        ' Certifica que está vísivel
        If Not .Visible Then Exit Sub
        
        ' Importa a tradução
        Language
        
        ' Incia o valor do padding lateral
        PaddingLeft = 13
        
        ' Incia o valor do padding superior
        PaddingTop = 32
        
        ' Desenha a janela
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        Select Case RandBackPlayer
            Case 1
                BackPlayerX = 2
                BackPlayerY = 247
            Case 2
                BackPlayerX = 2
                BackPlayerY = 432
            Case 3
                BackPlayerX = 2
                BackPlayerY = 617
            Case 4
                BackPlayerX = 2
                BackPlayerY = 802
            Case 5
                BackPlayerX = 203
                BackPlayerY = 247
            Case 6
                BackPlayerX = 203
                BackPlayerY = 432
            Case 7
                BackPlayerX = 203
                BackPlayerY = 617
            Case 8
                BackPlayerX = 203
                BackPlayerY = 802
        End Select
    
        ' Desenha o Background
        RenderTexture Tex_Gui(.Pic), .X + 2, .Y + 32, BackPlayerX, BackPlayerY, 199, 183, 199, 183
        
        ' Botões
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        ' Nome do Personagem
        If pCharInUsed(CurChar) Then
            CharNameText = Trim$(pCharName(CurChar))
            
            ' Desenha o texto do botão usar personagem
            RenderText Font_Default, TextUICharactersUse, .X + PaddingLeft + (125 / 2) - (GetTextWidth(Font_Default, TextUICharactersUse) / 2), (.Y + PaddingTop) + 146, White, , 255
            
            ' Desenha o texto do botão apagar personagem
            RenderText Font_Default, TextUICharactersDelete, .X + 141 + (50 / 2) - (GetTextWidth(Font_Default, TextUICharactersDelete) / 2) - 2, (.Y + PaddingTop) + 146, White, , 255
        Else
            CharNameText = TextUICharactersNone
            
            ' Desenha o texto do botão novo personagem
            RenderText Font_Default, TextUICharactersNew, .X + PaddingLeft + (178 / 2) - (GetTextWidth(Font_Default, TextUICharactersNew) / 2) - 2, (.Y + PaddingTop) + 146, White, , 255
        End If
        
        RenderText Font_Default, CharNameText, .X + PaddingLeft + (178 / 2) - ((GetTextWidth(Font_Default, CharNameText) / 2)) - 2, .Y + 37, D3DColorARGB(180, 250, 250, 250), False
        
        ' Sprite do personagem
        If pCharInUsed(CurChar) Then
            Sprite = pCharSprite(CurChar)
            If Sprite > 0 Then
                RenderTexture Tex_Character(Sprite), .X + ((.Width / 2) - (((GetPicWidth(Tex_Character(Sprite)) / 3) * 2) / 2)) - 2, .Y + 75, (GetPicWidth(Tex_Character(Sprite)) / 3), 0, (GetPicWidth(Tex_Character(Sprite)) / 3) * 2, (GetPicHeight(Tex_Character(Sprite)) / 4) * 2, GetPicWidth(Tex_Character(Sprite)) / 3, GetPicHeight(Tex_Character(Sprite)) / 4, D3DColorARGB(255, 255, 255, 255)
            End If
        End If
        
    
    End With
End Sub

Public Sub CharacterSelectMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        ' Ordena a janela ao ser clicada
        UpdateGuiOrder GUI_CHARACTERSELECT
        
        ' Verifica todos os itens
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        ' Verifica se foi movido
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub CharacterSelectMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
        
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERSELECT Then Exit Sub
        
        IsHovering = False
        
        ' Verifica todos os itens
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
                        
                        IsHovering = True
                        MouseIcon = 1
                    End If
                End If
            End If
        Next

        ' Verifica se foi movido
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Public Sub CharacterSelectMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERSELECT Then Exit Sub
        
        ' Verifica todos os itens
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Character_SwitchLeft
                                If CurChar > 1 Then
                                    RandBackPlayer = Format(RandomNumBetween(1, 8))
                                    CurChar = CurChar - 1
                                End If
                            Case ButtonEnum.Character_SwitchRight
                                If CurChar < MAX_PLAYERCHAR Then
                                    CurChar = CurChar + 1
                                    RandBackPlayer = Format(RandomNumBetween(1, 8))
                                End If
                            Case ButtonEnum.Character_New
                                CharName = vbNullString
                                SelGender = 0

                                GuiState GUI_CHARACTERSELECT, False
                                GuiState GUI_CHARACTERCREATE, True
                            Case ButtonEnum.Character_Use
                                If WaitTimer > GetTickCount Then
                                    AddAlert "Please wait for a few second before trying again", White
                                Else
                                    Menu_State MENU_STATE_USECHAR
                                    '//Prevent Spamming
                                    WaitTimer = GetTickCount + 5000
                                End If
                            Case ButtonEnum.Character_Delete
                                OpenChoiceBox TextUIChoiceDeleteCharacter, CB_CHARDEL
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

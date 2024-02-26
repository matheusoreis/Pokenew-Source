Attribute VB_Name = "CreateCharacter_Window"
' Padding lateral
Dim PaddingLeft As Long

' Padding da barra superior
Dim PaddingTop As Long

' Tamanho do texto na Caixa
Dim TextUIBoxSize As Long

Public Sub DrawCharacterCreate()
Dim i As Long
Dim Sprite As Long
Dim Width As Long, Height As Long

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
    
        ' Importa a tradução
        Language
        
        ' Certifica que é vísivel
        If Not .Visible Then Exit Sub
        
        ' Incia o valor do padding lateral
        PaddingLeft = 23
        
        ' Incia o valor do padding superior
        PaddingTop = 32
        
        ' Inicia o valor do texto na caixa
        TextUIBoxSize = 104
        
        ' Desenha a janela
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        ' Botões
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        ' Desenha o texto do username
        RenderText Font_Default, TextUICreateCharacterUsername, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUICreateCharacterUsername) / 2) - 2, (.Y + PaddingTop) + 13, White, , 255
        
        ' Desenha o texto do botão criar
        RenderText Font_Default, TextUICreateCharacterCreateButton, .X + 10 + (254 / 2) - (GetTextWidth(Font_Default, TextUICreateCharacterCreateButton) / 2) - 2, .Y + (PaddingTop + 142) + 34 / 2 - 12, White, , 255
        
        ' Textbox
        RenderText Font_Default, UpdateChatText(Font_Default, CharName, 110) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 13), Dark, , 255
        
        ' Seleção do gênero
        If SelGender = GENDER_MALE Then
            RenderTexture Tex_Gui(.Pic), .X + 26, .Y + 78, 288, 114, 107, 87, 107, 87
            RenderTexture Tex_Gui(.Pic), .X + 142, .Y + 78, 396, 114, 107, 87, 107, 87
            
            Sprite = 1
            Width = (GetPicWidth(Tex_Character(Sprite)) / 3) * 2
            Height = (GetPicHeight(Tex_Character(Sprite)) / 4) * 2
            RenderTexture Tex_Character(Sprite), .X + 26 + ((114 / 2) - (Width / 2)) - 2, .Y + 78 + ((107 / 2) - (Height / 2)) - 20, (Width / 2) * GenderAnim, 0, Width, Height, Width / 2, Height / 2
            Sprite = 2
            Width = (GetPicWidth(Tex_Character(Sprite)) / 3) * 2
            Height = (GetPicHeight(Tex_Character(Sprite)) / 4) * 2
            RenderTexture Tex_Character(Sprite), .X + 142 + ((114 / 2) - (Width / 2)) - 2, .Y + 78 + ((107 / 2) - (Height / 2)) - 20, (Width / 2), 0, Width, Height, Width / 2, Height / 2
        ElseIf SelGender = GENDER_FEMALE Then
            RenderTexture Tex_Gui(.Pic), .X + 26, .Y + 78, 396, 114, 107, 87, 107, 87
            RenderTexture Tex_Gui(.Pic), .X + 142, .Y + 78, 288, 114, 107, 87, 107, 87
            
            Sprite = 1
            Width = (GetPicWidth(Tex_Character(Sprite)) / 3) * 2
            Height = (GetPicHeight(Tex_Character(Sprite)) / 4) * 2
            RenderTexture Tex_Character(Sprite), .X + 26 + ((114 / 2) - (Width / 2)) - 2, .Y + 78 + ((107 / 2) - (Height / 2)) - 20, (Width / 2), 0, Width, Height, Width / 2, Height / 2
            Sprite = 2
            Width = (GetPicWidth(Tex_Character(Sprite)) / 3) * 2
            Height = (GetPicHeight(Tex_Character(Sprite)) / 4) * 2
            RenderTexture Tex_Character(Sprite), .X + 142 + ((114 / 2) - (Width / 2)) - 2, .Y + 78 + ((107 / 2) - (Height / 2)) - 20, (Width / 2) * GenderAnim, 0, Width, Height, Width / 2, Height / 2
        End If

    End With
End Sub

Public Sub CharacterCreateKeyPress(KeyAscii As Integer)
    
    ' Certifica que está visível
    If Not GUI(GuiEnum.GUI_CHARACTERCREATE).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        FoundError = False
                                    
        If WaitTimer > GetTickCount Then
            AddAlert TextUIWait, White
            FoundError = True
        End If
    
        If Not FoundError And Not CheckNameInput(CharName, False, (NAME_LENGTH - 1)) Then
            CurTextbox = 1
            AddAlert TextUICreateCharacterUsernameLenght, White
            FoundError = True
        End If
                                    
        ' Algum bug
        If CurChar = 0 Then
            GuiState GUI_CHARACTERCREATE, False
            GuiState GUI_CHARACTERSELECT, True
        End If
        
        ' Não encontrou erro
        If Not FoundError Then
            ' Envia os dados
            Menu_State MENU_STATE_ADDCHAR
            
            ' Previne o spam
            WaitTimer = GetTickCount + 5000
        End If
    End If
    
    If (isNameLegal(KeyAscii, True) And Len(CharName) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then CharName = InputText(CharName, KeyAscii)
End Sub

Public Sub CharacterCreateMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
        
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        ' Ordena a janela ao ser clicada
        UpdateGuiOrder GUI_CHARACTERCREATE
        
        ' Verifica todos os itens
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        ' Gênero
        If CursorY >= .Y + 78 And CursorY <= .Y + 78 + 87 Then
            If CursorX >= .X + 26 And CursorX <= .X + 26 + 114 Then
                SelGender = GENDER_MALE
            ElseIf CursorX >= .X + 142 And CursorX <= .X + 142 + 114 Then
                SelGender = GENDER_FEMALE
            End If
        End If
        
        ' Verifica se foi movido
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub CharacterCreateMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
        
        IsHovering = False
        
        ' Verifica todos os itens
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
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
        
        ' Gênero
        If CursorY >= .Y + 78 And CursorY <= .Y + 78 + 87 Then
            If CursorX >= .X + 26 And CursorX <= .X + 26 + 114 Then
                IsHovering = True
                MouseIcon = 1
            ElseIf CursorX >= .X + 142 And CursorX <= .X + 142 + 114 Then
                IsHovering = True
                MouseIcon = 1
            End If
        End If

        ' Passa o mouse pela textbox
        If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 120 And CursorY >= .Y + (PaddingTop + 13) And CursorY <= .Y + (PaddingTop + 13) + 23 Then
            IsHovering = True
            MouseIcon = 2
        End If

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


Public Sub CharacterCreateMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
        
        ' Verifica todos os itens
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.CharCreate_Confirm
                                FoundError = False
                                
                                If WaitTimer > GetTickCount Then
                                    AddAlert TextUIWait, White
                                    FoundError = True
                                End If

                                If Not FoundError And Not CheckNameInput(CharName, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 1
                                    AddAlert TextUICreateCharacterUsernameLenght, White
                                    FoundError = True
                                End If
                                
                                ' Algum erro
                                If CurChar = 0 Then
                                    GuiState GUI_CHARACTERCREATE, False
                                    GuiState GUI_CHARACTERSELECT, True
                                End If
    
                                ' Não encontrou nenhum erro
                                If Not FoundError Then
                                    ' Envia os dados
                                    Menu_State MENU_STATE_ADDCHAR
                                    
                                    ' Previne o spam
                                    WaitTimer = GetTickCount + 5000
                                End If
                            Case ButtonEnum.CharCreate_Close
                                GuiState GUI_CHARACTERCREATE, False
                                GuiState GUI_CHARACTERSELECT, True
                        End Select
                    End If
                End If
            End If
        Next
        
        ' Verifica se foi movido
        .InDrag = False
    End With
End Sub


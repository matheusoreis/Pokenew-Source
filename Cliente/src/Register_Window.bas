Attribute VB_Name = "Register_Window"
' Padding lateral
Dim PaddingLeft As Long

' Padding da barra superior
Dim PaddingTop As Long

' Tamanho do texto na Caixa
Dim TextUIBoxSize As Long

Public Sub DrawRegister()
Dim i As Byte
Dim PassColor As Long

    With GUI(GuiEnum.GUI_REGISTER)
        
        ' Importa a tradução
        Language
        
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        ' Incia o valor do padding lateral
        PaddingLeft = 23
        
        ' Incia o valor do padding superior
        PaddingTop = 32
        
        ' Inicia o valor do texto na caixa
        TextUIBoxSize = 104
        
        ' Desenha a janela
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        ' Desenha o texto do username
        RenderText Font_Default, TextUIRegisterUsername, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUIRegisterUsername) / 2) - 2, (.Y + PaddingTop) + 13, White, , 255
        
        ' Desenha o texto da password
        RenderText Font_Default, TextUIRegisterPassword, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUIRegisterPassword) / 2) - 2, (.Y + PaddingTop) + 55, White, , 255
        
        ' Desenha o texto da checkbox
        RenderText Font_Default, TextUIRegisterCheckBox, .X + (PaddingLeft + 19), (.Y + PaddingTop) + 107, White, , 255
        
        ' Desenha o texto da password
        RenderText Font_Default, TextUIRegisterEmail, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUIRegisterEmail) / 2) - 2, (.Y + PaddingTop) + 135, White, , 255
        
        ' Desenha os botões
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                ' Desenha o Botão
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        ' Desenha o texto do botão
        RenderText Font_Default, TextUIRegisterConfirm, .X + 23 + (254 / 2) - (GetTextWidth(Font_Default, TextUIRegisterConfirm) / 2) - 2, (.Y + 200) + 34 / 2 - 12, White, , 255
                
        ' Verificar se a confirmação de senha é a mesma
        If Len(Pass2) > 0 Then
            If Pass <> Pass2 Then
                PassColor = BrightRed
            Else
                PassColor = Dark
            End If
        Else
            PassColor = Dark
        End If
        
        ' Desenha as textbox
        Select Case CurTextbox
            Case 1 ' User
                RenderText Font_Default, UpdateChatText(Font_Default, User, 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 14), Dark
                If ShowPass = YES Then
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass2, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                Else
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass2), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                End If
                RenderText Font_Default, UpdateChatText(Font_Default, Email, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 135), Dark
            Case 2 ' Pass
                RenderText Font_Default, UpdateChatText(Font_Default, User, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 14), Dark
                If ShowPass = YES Then
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass, 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass2, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                Else
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass2), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                End If
                RenderText Font_Default, UpdateChatText(Font_Default, Email, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 135), Dark
            Case 3 ' Retype Pass
                RenderText Font_Default, UpdateChatText(Font_Default, User, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 14), Dark
                If ShowPass = YES Then
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass2, 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                Else
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass2), 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                End If
                RenderText Font_Default, UpdateChatText(Font_Default, Email, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 135), Dark
            Case 4 ' Email
                RenderText Font_Default, UpdateChatText(Font_Default, User, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 14), Dark
                If ShowPass = YES Then
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, Pass2, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                Else
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 47), PassColor
                    RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass2), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 68), PassColor
                End If
                RenderText Font_Default, UpdateChatText(Font_Default, Email, 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 135), Dark
        End Select
        
        ' Desenha as barras de senha forte
        If Len(Pass) > 0 Then
            If Len(Pass) >= 0 And Len(Pass) < ((NAME_LENGTH - 1) / 4) Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 237, 28, 36)
            ElseIf Len(Pass) >= ((NAME_LENGTH - 1) / 4) And Len(Pass) < ((NAME_LENGTH - 1) / 4) * 2 Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 198, 153, 0)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 198, 153, 0)
            ElseIf Len(Pass) >= ((NAME_LENGTH - 1) / 4) * 2 And Len(Pass) < ((NAME_LENGTH - 1) / 4) * 3 Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 53, 165, 51)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 53, 165, 51)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38 * 2, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 53, 165, 51)
            ElseIf Len(Pass) >= ((NAME_LENGTH - 1) / 4) * 3 And Len(Pass) < ((NAME_LENGTH - 1) + 5) Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 0, 162, 232)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 0, 162, 232)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38 * 2, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 0, 162, 232)
                RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize) + 38 * 3, .Y + (PaddingTop + 93), 0, 8, 36, 4, 1, 1, D3DColorARGB(255, 0, 162, 232)
            End If
        End If
        
        ' Checkbox para mostrar a senha
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125 + 17, 17, 17, 17, 17
        Else
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125, 17, 17, 17, 17
        End If
        
        If ShowPass = YES Then
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125 + 34, 17, 17, 17, 17
        End If
    End With
End Sub

' Método dos cliques na janela
Public Sub RegisterMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    With GUI(GuiEnum.GUI_REGISTER)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        ' Ordena a janela ao ser clicada
        UpdateGuiOrder GUI_REGISTER
        
        ' Verifica todos os itens
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        ' Clique na textbox
        If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 13) And CursorY <= .Y + (PaddingTop + 13) + 21 Then
            CurTextbox = 1
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 46) And CursorY <= .Y + (PaddingTop + 46) + 21 Then
            CurTextbox = 2
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 67) And CursorY <= .Y + (PaddingTop + 67) + 21 Then
            CurTextbox = 3
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 134) And CursorY <= .Y + (PaddingTop + 134) + 21 Then
            CurTextbox = 4
        End If
        
        ' Clique na checkbox
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 Then
            If ShowPass = YES Then
                ShowPass = NO
            Else
                ShowPass = YES
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

' Método ao passar o mouse por cima dos itens
Public Sub RegisterMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_REGISTER)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub
        
        IsHovering = False
        
        ' Verifica todos os itens
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
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
        
        ' Passar o mouse por cima das textbox
        If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 13) And CursorY <= .Y + (PaddingTop + 11) + 23 Then
            IsHovering = True
            MouseIcon = 2
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 46) And CursorY <= .Y + (PaddingTop + 44) + 23 Then
            IsHovering = True
            MouseIcon = 2
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 67) And CursorY <= .Y + (PaddingTop + 65) + 23 Then
            IsHovering = True
            MouseIcon = 2
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 148 And CursorY >= .Y + (PaddingTop + 134) And CursorY <= .Y + (PaddingTop + 132) + 23 Then
            IsHovering = True
            MouseIcon = 2
        End If
        
        ' Passar o mouse por cima da checkbox
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 Then
            IsHovering = True
            MouseIcon = 1
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

' Método para verificar entrar pressionando Return
Public Sub RegisterKeyPress(KeyAscii As Integer)
Dim FoundError As Boolean

    ' Certifica que está visível
    If Not GUI(GuiEnum.GUI_REGISTER).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub

    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        If CurTextbox = 1 Then
            CurTextbox = 2
        ElseIf CurTextbox = 2 Then
            CurTextbox = 3
        ElseIf CurTextbox = 3 Then
            CurTextbox = 4
        ElseIf CurTextbox = 4 Then
            If KeyAscii = vbKeyReturn Then
                FoundError = False
                
                If WaitTimer > GetTickCount Then
                    AddAlert TextUIWait, White
                    FoundError = True
                End If
                
                ' Erro no usuário
                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 1
                    AddAlert TextUIRegisterUsernameLenght, White
                    FoundError = True
                End If
                
                ' Erro na senha 1
                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 2
                    AddAlert TextUIRegisterPasswordMatch, White
                    FoundError = True
                End If
                
                ' Erro na senha 2
                If Not FoundError And (Pass <> Pass2) Then
                    CurTextbox = 3
                    AddAlert TextUIRegisterPasswordMatch, White
                    FoundError = True
                End If
                
                ' Erro no email
                If Not FoundError And Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
                    CurTextbox = 4
                    AddAlert TExtUIRegisterInvalidEmail, White
                    FoundError = True
                End If
                
                ' Não encontrou erro
                If Not FoundError Then
                    'Envia as informações de registro
                    Menu_State MENU_STATE_REGISTER
                    
                    ' Previni o spam
                    WaitTimer = GetTickCount + 5000
                End If
            Else
                CurTextbox = 1
            End If
        End If
    End If
    
    Select Case CurTextbox
        Case 1: If (isNameLegal(KeyAscii, True) And Len(User) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then User = InputText(User, KeyAscii)
        Case 2: If (isNameLegal(KeyAscii, True) And Len(Pass) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Pass = InputText(Pass, KeyAscii)
        Case 3: If (isNameLegal(KeyAscii, True) And Len(Pass2) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Pass2 = InputText(Pass2, KeyAscii)
        Case 4: If (isStringLegal(KeyAscii, True) And Len(Email) < (TEXT_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Email = InputText(Email, KeyAscii)
    End Select
End Sub

Public Sub RegisterMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_REGISTER)
    
        'Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub
        
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        
                        Select Case i
                            Case ButtonEnum.Register_Confirm
                                FoundError = False
                                
                                If WaitTimer > GetTickCount Then
                                    AddAlert TextUIWait, White
                                    FoundError = True
                                End If
                                
                                ' Erro no usuário
                                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 1
                                    AddAlert TextUIRegisterUsernameLenght, White
                                    FoundError = True
                                End If
                                
                                ' Erro na senha 1
                                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 2
                                    AddAlert TextUIRegisterPasswordLenght, White
                                    FoundError = True
                                End If
                                
                                ' Erro na senha 2
                                If Not FoundError And (Pass <> Pass2) Then
                                    CurTextbox = 3
                                    AddAlert TextUIRegisterPasswordMatch, White
                                    FoundError = True
                                End If
                                
                                ' Erro no email
                                If Not FoundError And Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
                                    CurTextbox = 4
                                    AddAlert TExtUIRegisterInvalidEmail, White
                                    FoundError = True
                                End If
                                
                                ' Não encontrou erro
                                If Not FoundError Then
                                    'Envia as informações de registro
                                    Menu_State MENU_STATE_REGISTER
                                    
                                    ' Previni o spam
                                    WaitTimer = GetTickCount + 5000
                                End If
                                
                            Case ButtonEnum.Register_Close
                                GuiState GUI_REGISTER, False
                                GuiState GUI_LOGIN, True, True
                                CurTextbox = 1
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

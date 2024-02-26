Attribute VB_Name = "Window_Login"
' Padding lateral
Dim PaddingLeft As Long

' Padding da barra superior
Dim PaddingTop As Long

' Tamanho do texto na Caixa
Dim TextUIBoxSize As Long

' Método que desenha a janela
Public Sub DrawLogin()
    Dim i As Long
    Dim X As Long
    Dim SString As String

    With GUI(GuiEnum.GUI_LOGIN)

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
        RenderText Font_Default, TextUILoginUsername, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUILoginUsername) / 2) - 2, (.Y + PaddingTop) + 13, White, , 255

        ' Desenha o texto da password
        RenderText Font_Default, TextUILoginPassword, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUILoginPassword) / 2) - 2, (.Y + PaddingTop) + 46, White, , 255

        ' Desenha o texto do server list
        RenderText Font_Default, TextUILoginServerList, (.X + PaddingLeft) + TextUIBoxSize / 2 - (GetTextWidth(Font_Default, TextUILoginServerList) / 2) - 2, (.Y + PaddingTop) + 79, White

        ' Desenha o texto da checkbox
        RenderText Font_Default, TextUILoginCheckBox, .X + (PaddingLeft + 19), (.Y + PaddingTop) + 106, White

        ' Desenha os botões
        For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
            If CanShowButton(i) Then
                ' Desenha o botão
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next
        
        ' Botão Registrar
        'If Not CreditVisible = True Then
        '    If CursorX >= Column / 2 - GetTextWidth(Font_Default, TextUIFooterCreateAccount) / 2 And CursorX <= Column / 2 - GetTextWidth(Font_Default, TextUIFooterCreateAccount) / 2 + GetTextWidth(Font_Default, TextUIFooterCreateAccount) And CursorY >= textY And CursorY <= textY + 40 Then
        '        If Not GUI(GuiEnum.GUI_REGISTER).Visible Then
        '            GuiState GUI_LOGIN, False
        '            GuiState GUI_REGISTER, True, True
        '            CurTextbox = 1
        '            User = vbNullString
        '            Pass = vbNullString
        '            Pass2 = vbNullString
        '            Email = vbNullString
        '        End If
        '    End If
        'End If

        ' Desenha o texto do botão
        RenderText Font_Default, TextUILoginEntryButton, .X + 23 + (254 / 2) - (GetTextWidth(Font_Default, TextUILoginEntryButton) / 2) - 2, (.Y + 161) + 34 / 2 - 11, White, , 255

        ' Desenha as textbox
        Select Case CurTextbox
        Case 1    ' User
            RenderText Font_Default, UpdateChatText(Font_Default, User, 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 13), Dark
            RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 46), Dark
        Case 2    ' Pass
            RenderText Font_Default, UpdateChatText(Font_Default, User, 130), .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 13), Dark
            RenderText Font_Default, UpdateChatText(Font_Default, CensorWord(Pass), 130) & TextLine, .X + (PaddingLeft + TextUIBoxSize), .Y + (PaddingTop + 46), Dark
        End Select

        ' Desenha a checkbox
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125 + 17, 17, 17, 17, 17
        Else
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125, 17, 17, 17, 17
        End If

        ' Desenha a lista de servidores
        If ShowServerList Then
            If CurServerList > 0 Then
                RenderText Font_Default, ServerName(CurServerList), (.X + PaddingLeft) + TextUIBoxSize + 2, .Y + (PaddingTop + 79), Dark
            End If

            If ServerList Then
                For X = 1 To MAX_SERVER_LIST
                    If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 140 And CursorY >= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((X - 2) * 20)) And CursorY <= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((X - 2) * 20)) + 20 Then
                        RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + 98 + ((20 * MAX_SERVER_LIST) + ((X - 2) * 20)), 0, 8, 150, 20, 1, 1, D3DColorARGB(255, 146, 146, 146)
                    Else
                        RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + TextUIBoxSize), .Y + 98 + ((20 * MAX_SERVER_LIST) + ((X - 2) * 20)), 0, 8, 150, 20, 1, 1, D3DColorARGB(240, 41, 41, 41)
                    End If
                    RenderText Font_Default, ServerName(X), .X + (PaddingLeft + TextUIBoxSize) + 4, .Y + 98 + ((20 * MAX_SERVER_LIST) + ((X - 2) * 20)) + 2, White
                Next
            End If

            ' Desenha a quantidade de jogadores neste servidor!
            'SString = Replace$(SString, ColourChar, vbNullString)
            SString = "Status:" & ColourChar & ServerInfo(CurServerList).Colour & Space(1) & ServerInfo(CurServerList).Status
            SString = SString & ColourChar & Yellow & " Players:" & ColourChar & ServerInfo(CurServerList).Colour & Space(1) & ServerInfo(CurServerList).Player
            ' Degrade
            RenderTexture Tex_Gui(12), 0, 5, 59, 241, (GetTextWidth(Font_Default, SString)) - 70, 20, 165, 1
            RenderText Font_Default, SString, 0, 5, Yellow
        End If


        ' Deseenha a checkbox
        If GameSetting.SavePass = YES Then
            RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + (PaddingTop + 107), 319, 125 + 34, 17, 17, 17, 17
        End If
    End With
End Sub

' Método dos cliques na janela
Public Sub LoginMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    With GUI(GuiEnum.GUI_LOGIN)
        
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        ' Ordena a janela ao ser clicada
        UpdateGuiOrder GUI_LOGIN
        
        ' Verifica todos os itens
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            Next
        End If
        
        ' Clique na textbox
        If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 12) And CursorY <= .Y + (PaddingTop + 12) + 23 Then
            CurTextbox = 1
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 45) And CursorY <= .Y + (PaddingTop + 45) + 23 Then
            CurTextbox = 2
        End If
        
        ' Clique na checkbox
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 Then
            If GameSetting.SavePass = YES Then
                GameSetting.SavePass = NO
            Else
                GameSetting.SavePass = YES
            End If
        End If
        
        ' Clique na lista de servidores
        If ServerList Then
            For i = 1 To MAX_SERVER_LIST
                If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 140 And CursorY >= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) And CursorY <= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) + 20 Then
                    CurServerList = i
                    LoadServerList CurServerList
                    '//Solicitação de informações de jogadores!
                    RequestServerInfo
                    Exit For
                End If
            Next
            ServerList = False
        Else
            If CursorX >= (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 79) And CursorY <= .Y + (PaddingTop + 79) + 23 Then
                ServerList = True
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
Public Sub LoginMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_LOGIN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
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
        End If
        
        ' Passar o mouse por cima das textbox
        If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 12) And CursorY <= .Y + (PaddingTop + 12) + 23 Then
            IsHovering = True
            MouseIcon = 2
        ElseIf CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 45) And CursorY <= .Y + (PaddingTop + 45) + 23 Then
            IsHovering = True
            MouseIcon = 2
        End If
        
        ' Passar o mouse por cima da checkbox
        If CursorX >= .X + PaddingLeft And CursorX <= .X + (PaddingLeft + 17) And CursorY >= .Y + (PaddingTop + 107) And CursorY <= .Y + (PaddingTop + 107) + 17 Then
            IsHovering = True
            MouseIcon = 1
        End If
        
        ' Passar o mouse por cima da lista de servidores
        If ServerList Then
            For i = 1 To MAX_SERVER_LIST
                If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 140 And CursorY >= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) And CursorY <= .Y + 98 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) + 20 Then
                    IsHovering = True
                    MouseIcon = 1
                    Exit For
                End If
            Next
            IsHovering = False
        Else
            If CursorX >= .X + (PaddingLeft + TextUIBoxSize) And CursorX <= .X + (PaddingLeft + TextUIBoxSize) + 150 And CursorY >= .Y + (PaddingTop + 79) And CursorY <= .Y + (PaddingTop + 79) + 23 Then
                IsHovering = True
                MouseIcon = 1
            End If
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
Public Sub LoginKeyPress(KeyAscii As Integer)
Dim FoundError As Boolean

    ' Certifica que está visível
    If Not GUI(GuiEnum.GUI_LOGIN).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        If CurTextbox = 1 Then
            CurTextbox = 2
        ElseIf CurTextbox = 2 Then
            If KeyAscii = vbKeyReturn Then
                FoundError = False
                
                ' Mensagem para esperar
                If WaitTimer > GetTickCount Then
                    AddAlert TextUIWait, White
                    FoundError = True
                End If
                
                ' Usuário inválido
                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 1
                    AddAlert TextUILoginInvalidUsername, White
                    FoundError = True
                End If
                
                ' Senha inválida
                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 2
                    AddAlert TextUILoginInvalidPassword, White
                    FoundError = True
                End If
                
                ' Se não encontra nenhum erro
                If Not FoundError Then
                    ' Envia as informações de login
                    Menu_State MENU_STATE_LOGIN
                    
                    ' Previne Spam
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
    End Select
End Sub

Public Sub LoginMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_LOGIN)
    
        ' Certifica que está visível
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub
        
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            
                            Select Case i
                                Case ButtonEnum.Login_Confirm
                                    FoundError = False
                                    
                                    ' Mensagem para esperar
                                    If WaitTimer > GetTickCount Then
                                        AddAlert TextUIWait, White
                                        FoundError = True
                                    End If
                                    
                                    ' Usuário inválido
                                    If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                                        CurTextbox = 1
                                        AddAlert TextUILoginInvalidUsername, White
                                        FoundError = True
                                    End If
                                    
                                    ' Senha inválida
                                    If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                                        CurTextbox = 2
                                        AddAlert TextUILoginInvalidPassword, White
                                        FoundError = True
                                    End If
                                    
                                    ' Se não encontra nenhum erro
                                    If Not FoundError Then
                                        ' Envia as informações de login
                                        Menu_State MENU_STATE_LOGIN
                                        
                                        ' Previne Spam
                                        WaitTimer = GetTickCount + 5000
                                    End If
                            End Select
                        End If
                    End If
                End If
            Next
        End If
        
        .InDrag = False
    End With
End Sub

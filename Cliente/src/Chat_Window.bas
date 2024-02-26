Attribute VB_Name = "Chat_Window"
Public Sub DrawChatbox()
    Dim i As Long

    With GUI(GuiEnum.GUI_CHATBOX)

        'If ChatOn Then

        ' Certifica que está visível
        If Not .Visible Then Exit Sub

        ' Importa a tradução
        'Language

        If ReInit Then Exit Sub
        
        If ChatMinimize Then Exit Sub

        ' Desenha a janela
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        ' Chatbox Text
        RenderChatTextBuffer

        ' Botões
        For i = ButtonEnum.Chatbox_ScrollUp To ButtonEnum.Chatbox_Minimize
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        ' Textbox
        If ChatOn Then
            RenderText Font_Default, UpdateChatText(Font_Default, MyChat & TextLine, 306), .X + 62, .Y + 143, White
        Else
            RenderText Font_Default, TextEnterToChat, .X + 62, .Y + 143, White
        End If

        If EditTab Then
            RenderText Font_Default, UpdateChatText(Font_Default, ChatTab & TextLine, 38), .X + 6, .Y + 143, White
        Else
            RenderText Font_Default, PreviewChatText(Font_Default, ChatTab, 35), .X + 6, .Y + 143, White
        End If

        If totalChatLines > MaxChatLine Then
            '//Scrollbar
            RenderTexture Tex_Gui(.Pic), .X + chatScrollX, .Y + chatScrollTop + (chatScrollL - chatScrollY), 467, 165, chatScrollW, chatScrollH, chatScrollW, chatScrollH
        End If
        ' End If
    End With
End Sub

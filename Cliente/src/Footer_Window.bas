Attribute VB_Name = "Footer_Window"

Public Sub DrawFooter()
Dim Width As Long, Height As Long
Dim i As Long
    
    ' Importa a tradução
    Language
    
    ' Declaração de variáveis para armazenar o tamanho da janela
    Y = Screen_Height
    X = Screen_Width
    textY = Y - 40
    
    ' Declaração da quantidade de colunas
    Column = X / 3
    
    ' Desenha o Footer
    RenderTexture Tex_System(gSystemEnum.UserInterface), 0, Screen_Height - 41, 0, 36, Screen_Width, 41, 41, 41
    
    ' Desenhar o texto do desenvolvedor
    RenderText Font_Default, TextUIFooterDeveloper, Column + (Column / 2) - GetTextWidth(Font_Default, TextUIFooterDeveloper) / 2, Screen_Height - 26, White
    
    ' Desenha o texto dos créditos
    If Not GUI(GuiEnum.GUI_CHARACTERSELECT).Visible = True Then
        
        ' Ao passar o mouse sobre o texto de créditos
        If CursorX >= Column * 2 + (Column / 2) - GetTextWidth(Font_Default, TextUIFooterCredits) / 2 And CursorX <= (Column * 2) + Column / 2 - GetTextWidth(Font_Default, TextUIFooterCredits) / 2 + GetTextWidth(Font_Default, TextUIFooterCredits) And CursorY >= Screen_Height - 40 And CursorY <= (Screen_Height - 40) + 40 Then
            IsHovering = True
            MouseIcon = 1
            colorHoverCredits = BrightGreen
        Else
            colorHoverCredits = White
        End If
        
        ' Ao passar o mouse sobre o texto criar uma nova conta
        If CursorX >= Column / 2 - GetTextWidth(Font_Default, TextUIFooterCreateAccount) / 2 And CursorX <= Column / 2 - GetTextWidth(Font_Default, TextUIFooterCreateAccount) / 2 + GetTextWidth(Font_Default, TextUIFooterCreateAccount) And CursorY >= textY And CursorY <= textY + 40 Then
            If CreditVisible = True Then
                IsHovering = False
            Else
                IsHovering = True
                MouseIcon = 1
            End If
            colorHoverCreateAccount = BrightGreen
        Else
            colorHoverCreateAccount = White
        End If
        
        ' Desenhar o texto para criar a conta
        If Not CreditVisible = True Then
            RenderText Font_Default, TextUIFooterCreateAccount, Column / 2 - GetTextWidth(Font_Default, TextUIFooterCreateAccount) / 2, Screen_Height - 26, colorHoverCreateAccount
        End If
        
        ' Desenhar o texto de créditos
        RenderText Font_Default, TextUIFooterCredits, Column * 2 + (Column / 2) - GetTextWidth(Font_Default, TextUIFooterCredits) / 2, Y - 26, colorHoverCredits
    Else
        ' Hover ao passar o mouse sobre o texto de trocar a senha
        If CursorX >= Column * 2 + (Column / 2) - GetTextWidth(Font_Default, TextUIFooterCredits) / 2 And CursorX <= (Column * 2) + Column / 2 - GetTextWidth(Font_Default, TextUIFooterCredits) / 2 + GetTextWidth(Font_Default, TextUIFooterCredits) And CursorY >= Screen_Height - 40 And CursorY <= (Screen_Height - 40) + 40 Then
            IsHovering = True
            MouseIcon = 1
            colorHoverChangePassword = BrightGreen
        Else
            colorHoverChangePassword = White
        End If
        
        ' Desenhar o texto de troca de senha
        RenderText Font_Default, TextUIFooterChangePassword, Column * 2 + (Column / 2) - GetTextWidth(Font_Default, TextUIFooterChangePassword) / 2, Y - 26, colorHoverChangePassword
        
    End If
    
    DrawCredit
End Sub


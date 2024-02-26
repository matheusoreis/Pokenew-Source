Attribute VB_Name = "Credit_Window"

Public Sub DrawCredit()
Dim i As Long
Dim showText As String
Dim Color As Long
Dim CreditAlpha As Long

    If Not CreditVisible Then Exit Sub
    RenderTexture Tex_System(gSystemEnum.UserInterface), 0, (Screen_Height - 40) - CreditOffset, 0, 8, Screen_Width, CreditOffset, 1, 1, D3DColorARGB(90, 0, 0, 0)
    
    If Not CreditState = 0 Then Exit Sub
    
    '//Render Credit text
    If CreditTextCount > 0 And (CreditOffset >= (Screen_Height - 40)) Then
        For i = 0 To CreditTextCount
            showText = Trim$(Credit(i).Text)
            Color = White
            
            If Left$(showText, 2) = "#h" Then
                showText = Mid$(showText, 3, Len(showText) - 2)
                Color = Yellow
            End If
                    
            If Credit(i).Y >= -32 And Credit(i).Y <= (Screen_Height - 40) + 32 Then
                '//Alpha Fading
                If Credit(i).Y >= -32 And Credit(i).Y <= 255 Then
                    CreditAlpha = Credit(i).Y
                ElseIf Credit(i).Y >= (Screen_Height - 40) - 255 And Credit(i).Y <= (Screen_Height - 40) + 32 Then
                    CreditAlpha = (Screen_Height - 40) - Credit(i).Y
                Else
                    CreditAlpha = 255
                End If
                        
                If CreditAlpha >= 0 And CreditAlpha <= 255 Then
                    RenderText Font_Default, showText, (Screen_Width / 2) - (GetTextWidth(Font_Default, showText) / 2), Credit(i).Y, Color, True, CreditAlpha
                End If
            End If
        Next
    
    End If
End Sub

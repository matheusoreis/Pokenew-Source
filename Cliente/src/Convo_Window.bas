Attribute VB_Name = "Convo_Window"
' ***************
' ** Convo **
' ***************

Public Sub DrawConvo()
Dim i As Long
Dim Sprite As Long
Dim spriteWidth As Long, spriteHeight As Long
Dim scaleWidth As Long, scaleHeight As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Render black alpha
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(150, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Error
        If ConvoNum <= 0 Or ConvoData <= 0 Then Exit Sub
        
        '//Render Sprite
        If ConvoNpcNum > 0 Then
            Sprite = Npc(ConvoNpcNum).Sprite
            If Sprite > 0 Then
                spriteWidth = GetPicWidth(Tex_Character(Sprite)) / 3
                spriteHeight = GetPicHeight(Tex_Character(Sprite)) / 4
                scaleWidth = spriteWidth * 4
                scaleHeight = spriteHeight * 4
                RenderTexture Tex_Character(Sprite), .X + ((.Width / 2) - (scaleWidth / 2)), .Y - scaleHeight + 10, spriteWidth, 0, scaleWidth, scaleHeight, spriteWidth, spriteHeight
            End If
        End If
        
        RenderArrayText Font_Default, ConvoRenderText, .X + 25, .Y + 25, 400, White, , True
        
        If Len(ConvoText) > ConvoDrawTextLen Then
            RenderTexture Tex_System(gSystemEnum.CursorLoad), .X + 425, .Y + 115, 15 * CursorLoadAnim, 0, 15, 15, 15, 15
        End If
        
        '//Convo Reply
        If ConvoShowButton Then
            '//Render black alpha
            RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(150, 0, 0, 0)
        
            '//Buttons
            Dim ButtonText As String, DrawText As Boolean
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    RenderTexture Tex_Gui(.Pic), Button(i).X, Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                    '//Render Button Text
                    Select Case i
                        Case ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                            If Len(Trim$(ConvoReply((i + 1) - ButtonEnum.Convo_Reply1))) > 0 Then
                                ButtonText = ((i + 1) - ButtonEnum.Convo_Reply1) & ": " & Trim$(ConvoReply((i + 1) - ButtonEnum.Convo_Reply1))
                                DrawText = True
                            End If
                        Case Else: DrawText = False
                    End Select
                    If DrawText Then
                        Select Case Button(i).State
                            Case ButtonState.StateNormal: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 229, 229, 229), False
                            Case ButtonState.StateHover: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5, D3DColorARGB(255, 255, 255, 255), False
                            Case ButtonState.StateClick: RenderText Ui_Default, ButtonText, (Button(i).X) + 5, (Button(i).Y) + ((Button(i).Height / 2) - (8)) - 5 + 3, D3DColorARGB(255, 255, 255, 255), False
                        End Select
                    End If
                End If
            Next
        End If
    End With
End Sub


Public Sub ConvoMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            Next
        End If
        
        '//Skip Scrolling Text
        If Not ConvoShowButton Then
            If ConvoNum > 0 Then
                If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
                    If Len(ConvoText) > ConvoDrawTextLen Then
                        ConvoDrawTextLen = Len(ConvoText)
                        ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                    Else
                        '//Proceed to next convo
                        If ConvoNoReply = YES Then
                            '//Proceed to next
                            SendProcessConvo
                        Else
                            '//Show Choice
                            ConvoShowButton = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub ConvoMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
                
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            Next
        End If
        
        If Not ConvoShowButton Then
            If ConvoNum > 0 Then
                If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
            End If
        End If
    End With
End Sub

Public Sub ConvoMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                                    SendProcessConvo ((i + 1) - ButtonEnum.Convo_Reply1)
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub

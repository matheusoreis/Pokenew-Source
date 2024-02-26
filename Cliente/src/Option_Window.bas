Attribute VB_Name = "Option_Window"
' Padding lateral
Dim PaddingLeft As Long

' Padding da barra superior
Dim PaddingTop As Long

' Tamanho do texto na Caixa
Dim TextUIBoxSize As Long

Public Sub DrawOption()
    Dim i As Long
    Dim tmpX As Long, tmpY As Long
    Dim Count As Long
    Dim X As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        Language

        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        ' Incia o valor do padding lateral
        PaddingLeft = 152

        ' Incia o valor do padding superior
        PaddingTop = 42

        ' Inicia o valor do texto na caixa
        TextUIBoxSize = 104

        '//Buttons
        'Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow = i Then Button(i).State = ButtonState.StateClick

            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height

                Select Case i

                Case 18
                    RenderText Font_Default, TextUIOptionVideoButton, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIOptionVideoButton) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                Case 19
                    RenderText Font_Default, TextUIOptionSoundButton, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIOptionSoundButton) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                Case 20
                    RenderText Font_Default, TextUIOptionGameButton, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIOptionGameButton) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                Case 21
                    RenderText Font_Default, TextUIOptionControlButton, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIOptionControlButton) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                End Select

            End If
        Next

        Select Case setWindow
        Case ButtonEnum.Option_Video

            '//Fullscreen
            tmpX = 105: tmpY = 45
            RenderText Font_Default, "Fullscreen", .X + PaddingLeft + 25, .Y + tmpY, D3DColorARGB(255, 229, 229, 229), False

            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + tmpY, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + tmpY, 52, 403, 17, 17, 17, 17
            End If

            If isFullscreen = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + tmpY, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha o texto da resolução
            RenderText Font_Default, "Janela: ", .X + PaddingLeft, .Y + tmpY + 32, White, , 255

            ' Desenha o fundo
            RenderTexture Tex_System(gSystemEnum.UserInterface), .X + (PaddingLeft + 61), .Y + tmpY + 32, 0, 8, 180 - 27, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)

            ' Desenha a lista de resoluções
            If ShowResolutionList Then

                If CurResolutionList > 0 Then
                    RenderText Font_Default, ResolutionName(CurResolutionList), (.X + PaddingLeft) + 64, .Y + tmpY + 32, White
                Else
                    RenderText Font_Default, GameSetting.Width & "x" & GameSetting.Height, (.X + PaddingLeft) + 64, .Y + tmpY + 32, White
                End If

                If ResolutionList Then
                    ' Desenha uma arrow /\ nas resoluções, apenas detalhe, nada demais.
                    RenderTexture Tex_Gui(.Pic), .X + PaddingLeft + 198, .Y + tmpY + 31, Button(ButtonEnum.Option_cTabUp).StartX(Button(ButtonEnum.Option_cTabUp).State), Button(ButtonEnum.Option_cTabUp).StartY(Button(ButtonEnum.Option_cTabUp).State), Button(ButtonEnum.Option_cTabUp).Width - 3, Button(ButtonEnum.Option_cTabUp).Height - 3, Button(ButtonEnum.Option_cTabUp).Width, Button(ButtonEnum.Option_cTabUp).Height
                    For X = 1 To MAX_RESOLUTION_LIST
                        If CursorX >= .X + PaddingLeft + 64 And CursorX <= .X + PaddingLeft + 64 + 140 And CursorY >= .Y + ((20 * MAX_RESOLUTION_LIST) + ((X - 2) * 20)) And CursorY <= .Y + ((20 * MAX_RESOLUTION_LIST) + ((X - 2) * 20)) + 20 Then
                            RenderTexture Tex_System(gSystemEnum.UserInterface), .X + PaddingLeft + 64, .Y + ((20 * MAX_RESOLUTION_LIST) + ((X - 2) * 20)), 0, 8, 150, 20, 1, 1, D3DColorARGB(100, 0, 0, 0)
                        Else
                            RenderTexture Tex_System(gSystemEnum.UserInterface), .X + PaddingLeft + 64, .Y + ((20 * MAX_RESOLUTION_LIST) + ((X - 2) * 20)), 0, 8, 150, 20, 1, 1, D3DColorARGB(100, 0, 0, 0)
                        End If
                        RenderText Font_Default, ResolutionName(X), .X + PaddingLeft + 64 + 4, .Y + ((20 * MAX_RESOLUTION_LIST) + ((X - 2) * 20)) + 2, White
                    Next
                Else
                    ' Desenha uma arrow \/ nas resoluções, apenas detalhe, nada demais.
                    RenderTexture Tex_Gui(.Pic), .X + PaddingLeft + 198, .Y + tmpY + 31, Button(ButtonEnum.Option_cTabDown).StartX(Button(ButtonEnum.Option_cTabDown).State), Button(ButtonEnum.Option_cTabDown).StartY(Button(ButtonEnum.Option_cTabDown).State), Button(ButtonEnum.Option_cTabDown).Width - 3, Button(ButtonEnum.Option_cTabDown).Height - 3, Button(ButtonEnum.Option_cTabDown).Width, Button(ButtonEnum.Option_cTabDown).Height
                End If
            End If

        Case ButtonEnum.Option_Sound
            RenderText Font_Default, TextUIOptionMusic, .X + 152, .Y + 45, D3DColorARGB(255, 229, 229, 229), False
            RenderText Font_Default, TextUIOptionSound, .X + 152, .Y + 27 + 45, D3DColorARGB(255, 229, 229, 229), False
            For i = 1 To MAX_VOLUME
                If BGVolume >= i Then
                    RenderTexture Tex_System(gSystemEnum.UserInterface), .X + 152 + 162 + ((8 + 3) * (i - 1)), .Y + 45, 0, 8, 9, 20, 1, 1, D3DColorARGB(255, 94, 177, 94)
                Else
                    RenderTexture Tex_System(gSystemEnum.UserInterface), .X + 152 + 162 + ((8 + 3) * (i - 1)), .Y + 45, 0, 8, 9, 20, 1, 1, D3DColorARGB(255, 100, 100, 100)
                End If
                If SEVolume >= i Then
                    RenderTexture Tex_System(gSystemEnum.UserInterface), .X + 152 + 162 + ((8 + 3) * (i - 1)), .Y + 27 + 45, 0, 8, 9, 20, 1, 1, D3DColorARGB(255, 94, 177, 94)
                Else
                    RenderTexture Tex_System(gSystemEnum.UserInterface), .X + 152 + 162 + ((8 + 3) * (i - 1)), .Y + 27 + 45, 0, 8, 9, 20, 1, 1, D3DColorARGB(255, 100, 100, 100)
                End If
            Next

        Case ButtonEnum.Option_Game
            'Desenha a gui
            RenderText Font_Default, TextUIOptionPath, .X + PaddingLeft, .Y + 45, D3DColorARGB(255, 229, 229, 229), False
            RenderTexture Tex_System(gSystemEnum.UserInterface), .X + PaddingLeft + 80, .Y + 45, 0, 8, 180, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)

            If GuiPathEdit Then
                RenderText Font_Default, GuiPath & TextLine, .X + PaddingLeft + 84, .Y + 46, D3DColorARGB(255, 229, 229, 229), False
            Else
                RenderText Font_Default, "...\" & GuiPath & "\", .X + PaddingLeft + 84, .Y + 46, D3DColorARGB(255, 229, 229, 229), False
            End If

            ' Desenha o fps
            RenderText Font_Default, TextUIOptionsFps, .X + PaddingLeft + 25, .Y + 70, D3DColorARGB(255, 229, 229, 229), False
            If CursorX >= .X + tmpX And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + 70 And CursorY <= .Y + 70 + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 70, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 70, 52, 403, 17, 17, 17, 17
            End If
            If FPSvisible = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 70, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha o ping
            RenderText Font_Default, TextUIOptionsPing, .X + PaddingLeft + 25, .Y + 90, D3DColorARGB(255, 229, 229, 229), False
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + 90 And CursorY <= .Y + 90 + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 90, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 90, 52, 403, 17, 17, 17, 17
            End If
            If PingVisible = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 90, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha o ínicio rápido
            RenderText Font_Default, TextUIOptionsFast, .X + PaddingLeft + 25, .Y + 110, D3DColorARGB(255, 229, 229, 229), False
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + 110 And CursorY <= .Y + 110 + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 110, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 110, 52, 403, 17, 17, 17, 17
            End If
            If tSkipBootUp = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 110, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha o nome
            RenderText Font_Default, TextUIOptionName, .X + PaddingLeft + 25, .Y + 130, D3DColorARGB(255, 229, 229, 229), False
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + 130 And CursorY <= .Y + 130 + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 130, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 130, 52, 403, 17, 17, 17, 17
            End If
            If Namevisible = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 130, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha o PPBar
            RenderText Font_Default, TextUIOptionPP, .X + PaddingLeft + 25, .Y + 150, D3DColorARGB(255, 229, 229, 229), False
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + 150 And CursorY <= .Y + 150 + 17 And GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 150, 52, 403 + 17, 17, 17, 17, 17
            Else
                RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 150, 52, 403, 17, 17, 17, 17
            End If
            If PPBarvisible = YES Then RenderTexture Tex_Gui(.Pic), .X + PaddingLeft, .Y + 150, 52, 403 + 34, 17, 17, 17, 17

            ' Desenha a tradução
            RenderText Font_Default, TextUIOptionLanguage, .X + PaddingLeft, .Y + 190 + 4, D3DColorARGB(255, 229, 229, 229), False
            For i = 1 To MAX_LANGUAGE
                If (tmpCurLanguage + 1) = i Then
                    RenderTexture Tex_Misc(Misc_Language), .X + 80 + PaddingLeft + ((i - 1) * 55), .Y + 190, 0, 25 * (i - 1), 45, 25, 45, 25, D3DColorARGB(255, 255, 255, 255)
                Else
                    RenderTexture Tex_Misc(Misc_Language), .X + 80 + PaddingLeft + ((i - 1) * 55), .Y + 190, 0, 25 * (i - 1), 45, 25, 45, 25, D3DColorARGB(100, 255, 255, 255)
                End If
            Next

            ' Desenha a Opção de Controles
        Case ButtonEnum.Option_Control

            ' Desenha o Scroll
            RenderTexture Tex_Gui(.Pic), .X + 414, .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY), 328, 310, 19, 35, 19, 35

            ' Quantidade dos controles
            For i = 1 To ControlScrollViewLine
                Count = ControlViewCount + (i)
                If Count > 0 And Count <= ControlEnum.Control_Count - 1 Then
                    RenderTexture Tex_Gui(.Pic), .X + 152, .Y + 43 + ((25 + 5) * (i - 1)), 93, 389, 254, 28, 254, 28

                    ' Desenha o Nome
                    RenderText Font_Default, ControlKey(Count).keyName, .X + 158, .Y + 43 + ((25 + 5) * (i - 1)) + 3, White

                    If editKey = Count Then
                        RenderTexture Tex_Gui(.Pic), .X + 290, .Y + 43 + ((25 + 5) * (i - 1)), 231, 417, 116, 28, 116, 28
                    End If

                    ' Desenha a Key
                    RenderText Font_Default, GetKeyCodeName(TmpKey(Count)), .X + 295, .Y + 43 + ((25 + 5) * (i - 1)) + 3, Dark
                End If
            Next
        End Select
    End With
End Sub

' ***************
' ** Option **
' ***************
Public Sub OptionKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_OPTION).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then Exit Sub
    
    If GuiPathEdit Then
        GuiPath = InputText(GuiPath, KeyAscii)
        setDidChange = True
    End If
End Sub

Public Sub OptionKeyUp(KeyCode As Integer, Shift As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_OPTION).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then Exit Sub
    
    If editKey > 0 And editKey <= ControlEnum.Control_Count - 1 Then
        If Not InvalidInput(KeyCode) Then
            If Not CheckSameKey(KeyCode) Then
                TmpKey(editKey) = KeyCode
                '//Exit editing
                editKey = 0
                setDidChange = True
            Else
                AddAlert "Key Input already in used", White
            End If
        Else
            AddAlert "Invalid Key Input", White
        End If
    End If
End Sub

Public Sub OptionMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim tmpX As Long, tmpY As Long
    Dim Count As Long
    Dim curResolution As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Set to top most
        UpdateGuiOrder GUI_OPTION

        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                            Select Case i
                            Case ButtonEnum.Option_cTabUp
                                ControlScrollUp = True
                                ControlScrollDown = False
                                ControlScrollTimer = GetTickCount
                                If ControlViewCount > 0 Then
                                    ControlViewCount = ControlViewCount - 1
                                    ControlScrollY = (ControlViewCount * ControlScrollLength) \ ControlMaxViewLine
                                    ControlScrollY = (ControlScrollLength - ControlScrollY)
                                End If
                            Case ButtonEnum.Option_cTabDown
                                ControlScrollUp = False
                                ControlScrollDown = True
                                ControlScrollTimer = GetTickCount
                                If ControlViewCount + (ControlScrollViewLine) < ControlEnum.Control_Count - 1 Then
                                    ControlViewCount = ControlViewCount + 1
                                    ControlScrollY = (ControlViewCount * ControlScrollLength)
                                    ControlScrollY = (ControlScrollY \ ControlMaxViewLine)
                                    ControlScrollY = (ControlScrollLength - ControlScrollY)
                                End If
                            End Select
                        End If
                    End If
                End If
            End If
        Next

        GuiPathEdit = False

        '//Window
        Select Case setWindow
        Case ButtonEnum.Option_Video
            '//Fullscreen
            tmpX = 105: tmpY = 45
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If isFullscreen = YES Then
                    isFullscreen = NO
                Else
                    isFullscreen = YES
                End If
                setDidChange = True
            End If

            ' Clique sobre a lista de resolução
            If ResolutionList Then
                For i = 1 To MAX_RESOLUTION_LIST
                    If CursorX >= .X + PaddingLeft + 64 And CursorX <= .X + PaddingLeft + 64 + 140 And CursorY >= .Y + ((20 * MAX_RESOLUTION_LIST) + ((i - 2) * 20)) And CursorY <= .Y + ((20 * MAX_RESOLUTION_LIST) + ((i - 2) * 20)) + 20 Then
                        CurResolutionList = i

                        ' Verifica a resolução atual e seta os novos valores
                        Select Case CurResolutionList
                        Case 1
                            WidthSize = "800"
                            HeightSize = "608"
                        Case 2
                            WidthSize = "1280"
                            HeightSize = "704"
                        Case 3
                            WidthSize = "1344"
                            HeightSize = "704"
                        Case 4
                            WidthSize = "1600"
                            HeightSize = "832"
                        Case 5
                            WidthSize = "1856"
                            HeightSize = "960"
                        Case 6
                            WidthSize = "2432"
                            HeightSize = "960"
                        End Select

                        setDidChange = True
                        Exit For
                    End If
                Next
                ResolutionList = False
            Else
                If CursorX >= .X + PaddingLeft + 64 And CursorX <= .X + PaddingLeft + 64 + 140 And CursorY >= .Y + tmpY + 32 And CursorY <= .Y + tmpY + 32 + 23 Then
                    ResolutionList = True
                End If
            End If

        Case ButtonEnum.Option_Sound
            tmpX = 152: tmpY = 45
            For i = 1 To MAX_VOLUME
                If CursorX >= .X + tmpX + 162 + ((8 + 3) * (i - 1)) And CursorX <= .X + tmpX + 162 + ((8 + 3) * (i - 1)) + 9 Then
                    '//Background Music
                    If CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 20 Then
                        BGVolume = i
                        setDidChange = True
                    ElseIf CursorY >= .Y + 27 + tmpY And CursorY <= .Y + 27 + tmpY + 20 Then
                        '//Sound Effect
                        SEVolume = i
                        setDidChange = True
                    End If
                End If
            Next
        Case ButtonEnum.Option_Game
            '//Show Fps
            tmpX = 105: tmpY = 70
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If FPSvisible = YES Then
                    FPSvisible = NO
                Else
                    FPSvisible = YES
                End If
                setDidChange = True
            End If

            '//Show Ping
            tmpX = 105: tmpY = 90
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If PingVisible = YES Then
                    PingVisible = NO
                Else
                    PingVisible = YES
                End If
                setDidChange = True
            End If

            '//Skip Boot Up
            tmpX = 105: tmpY = 110
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If tSkipBootUp = YES Then
                    tSkipBootUp = NO
                Else
                    tSkipBootUp = YES
                End If
                setDidChange = True
            End If

            '//Name Visible
            tmpX = 105: tmpY = 130
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If Namevisible = YES Then
                    Namevisible = NO
                Else
                    Namevisible = YES
                End If
                setDidChange = True
            End If

            '//PP Bar
            tmpX = 105: tmpY = 150
            If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                If PPBarvisible = YES Then
                    PPBarvisible = NO
                Else
                    PPBarvisible = YES
                End If
                setDidChange = True
            End If

        Case ButtonEnum.Option_Control
            For i = 1 To ControlScrollViewLine
                Count = ControlViewCount + (i)
                If Count > 0 And Count <= ControlEnum.Control_Count - 1 Then
                    If CursorX >= .X + 290 And CursorX <= .X + 290 + 114 And CursorY >= .Y + 44 + ((25 + 5) * (i - 1)) And CursorY <= .Y + 44 + ((25 + 5) * (i - 1)) + 24 Then
                        editKey = Count
                    End If
                End If
            Next

            If CursorX >= .X + 414 And CursorX <= .X + 414 + 19 And CursorY >= .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) And CursorY <= .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) + ControlScrollSize Then
                ControlScrollHold = True
            End If

        End Select

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub OptionMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long
Dim Count As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
                
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Window
        Select Case setWindow
            Case ButtonEnum.Option_Video
                '//Fullscreen
                tmpX = 105: tmpY = 45
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                If ResolutionList Then
                    For i = 1 To MAX_RESOLUTION_LIST
                        If CursorX >= .X + PaddingLeft + 64 And CursorX <= .X + PaddingLeft + 64 + 140 And CursorY >= .Y + ((20 * MAX_RESOLUTION_LIST) + ((i - 2) * 20)) And CursorY <= .Y + ((20 * MAX_RESOLUTION_LIST) + ((i - 2) * 20)) + 20 Then
                            IsHovering = True
                            MouseIcon = 1
                            Exit For
                        End If
                    Next
                    IsHovering = False
                Else
                    If CursorX >= .X + PaddingLeft + 64 And CursorX <= .X + PaddingLeft + 64 + 140 And CursorY >= .Y + tmpY + 32 And CursorY <= .Y + tmpY + 32 + 23 Then
                        IsHovering = True
                        MouseIcon = 1
                    End If
                End If
        
            Case ButtonEnum.Option_Sound
                tmpX = 152: tmpY = 45
                For i = 1 To MAX_VOLUME
                    If CursorX >= .X + tmpX + 162 + ((8 + 3) * (i - 1)) And CursorX <= .X + tmpX + 162 + ((8 + 3) * (i - 1)) + 9 Then
                        '//Background Music
                        If CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 20 Then
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        ElseIf CursorY >= .Y + 27 + tmpY And CursorY <= .Y + 27 + tmpY + 20 Then
                        '//Sound Effect
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                Next
            Case ButtonEnum.Option_Game
                '//Gui Path
                If CursorX >= .X + 165 And CursorX <= .X + 165 + 180 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 18 Then
                    IsHovering = True
                    MouseIcon = 2 '//I-Beam
                End If
                
                '//FPS
                tmpX = 105: tmpY = 70
                If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Ping
                tmpX = 105: tmpY = 90
                If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Skip BootUp
                tmpX = 105: tmpY = 110
                If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Show Name
                tmpX = 105: tmpY = 130
                If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Show PP
                tmpX = 105: tmpY = 150
                If CursorX >= .X + PaddingLeft And CursorX <= .X + PaddingLeft + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Language
                tmpY = 190
                For i = 1 To MAX_LANGUAGE
                    If CursorX >= .X + PaddingLeft + 80 + ((i - 1) * 55) And CursorX <= .X + PaddingLeft + 80 + ((i - 1) * 55) + 45 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 25 Then
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                Next
            Case ButtonEnum.Option_Control
                '//Control Key
                For i = 1 To ControlScrollViewLine
                    Count = ControlViewCount + (i)
                    If Count > 0 And Count <= ControlEnum.Control_Count - 1 Then
                        If CursorX >= .X + 290 And CursorX <= .X + 290 + 114 And CursorY >= .Y + 44 + ((25 + 5) * (i - 1)) And CursorY <= .Y + 44 + ((25 + 5) * (i - 1)) + 24 Then
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                Next
                
                If CursorX >= .X + 414 And CursorX <= .X + 414 + 19 And CursorY >= .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) And CursorY <= .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) + ControlScrollSize Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Scroll moving
                If ControlScrollHold Then
                
                    '//Upward
                    If CursorY < .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) + (ControlScrollSize / 2) Then
                        If ControlScrollY < ControlScrollEndY - ControlScrollSize Then
                            ControlScrollY = (CursorY - (.Y + ControlScrollStartY + (ControlScrollEndY - ControlScrollSize)) - (ControlScrollSize / 2)) * -1
                            If ControlScrollY >= ControlScrollEndY - ControlScrollSize Then ControlScrollY = ControlScrollEndY - ControlScrollSize
                            
                            ControlViewCount = ControlViewCount - 1
                            
                            If ControlViewCount < 0 Then
                                ControlViewCount = 0
                            End If
                            
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            
                        End If
                    End If
                    
                    '//Downward
                    If CursorY > .Y + ControlScrollStartY + ((ControlScrollEndY - ControlScrollSize) - ControlScrollY) + ControlScrollSize - (ControlScrollSize / 2) Then
                        If ControlScrollY > 0 Then
                            ControlScrollY = (CursorY - (.Y + ControlScrollStartY + (ControlScrollEndY - ControlScrollSize)) - ControlScrollSize + (ControlScrollSize / 2)) * -1
                            If ControlScrollY <= 0 Then ControlScrollY = 0
                            
                            '
                            ControlViewCount = ControlViewCount + 1
                            
                            If ControlViewCount >= ControlEnum.Control_Count Then
                                ControlViewCount = ControlEnum.Control_Count - 1
                            End If
                            
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            
                        End If
                    End If
                    
                End If
        
        End Select
        
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

Public Sub OptionMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, z As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.Option_Close
                                    If setDidChange Then
                                        OpenChoiceBox TextUIChoiceSave, CB_SAVESETTING
                                    Else
                                        GuiState GUI_OPTION, False
                                    End If
                                Case ButtonEnum.Option_Video, ButtonEnum.Option_Sound, ButtonEnum.Option_Game, ButtonEnum.Option_Control
                                    setWindow = i
                                Case ButtonEnum.Option_sMusicUp
                                    If BGVolume < MAX_VOLUME Then
                                        BGVolume = BGVolume + 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sMusicDown
                                    If BGVolume > 0 Then
                                        BGVolume = BGVolume - 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sSoundUp
                                    If SEVolume < MAX_VOLUME Then
                                        SEVolume = SEVolume + 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sSoundDown
                                    If SEVolume > 0 Then
                                        SEVolume = SEVolume - 1
                                        setDidChange = True
                                    End If
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Window
        Select Case setWindow
            Case ButtonEnum.Option_Video

            Case ButtonEnum.Option_Sound
                
            Case ButtonEnum.Option_Game
                '//Gui Path
                If CursorX >= .X + 165 And CursorX <= .X + 165 + 180 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 18 Then
                    GuiPathEdit = True
                End If
                
                '//Language
                For i = 1 To MAX_LANGUAGE
                    If CursorX >= .X + PaddingLeft + 80 + ((i - 1) * 55) And CursorX <= .X + PaddingLeft + 80 + ((i - 1) * 55) + 45 And CursorY >= .Y + 190 And CursorY <= .Y + 190 + 25 Then
                        tmpCurLanguage = (i - 1)
                        setDidChange = True
                    End If
                Next
            Case ButtonEnum.Option_Control
            
        End Select
        
        '//Control Scroll
        ControlScrollHold = False
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

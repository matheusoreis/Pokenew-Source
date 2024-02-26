Attribute VB_Name = "Choice_Window"
' ***************
' ** ChoiceBox **
' ***************
Public Sub DrawChoiceBox()
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
        
        '//Buttons
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                Select Case i
                    Case ChoiceBox_Yes
                        RenderText Font_Default, TextUIChoiceYes, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIChoiceYes) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White
                    Case ChoiceBox_No
                        RenderText Font_Default, TextUIChoiceNo, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIChoiceNo) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White
                End Select
            
            End If
        Next
        
        '//Render Text
        RenderArrayText Font_Default, ChoiceBoxText, .X + 10, .Y + 36, 250, White
    End With
End Sub

Public Sub ChoiceBoxMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_CHOICEBOX
        
        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
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

Public Sub ChoiceBoxMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
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

Public Sub ChoiceBoxMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim z As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                        Case ButtonEnum.ChoiceBox_Yes
                            Select Case ChoiceBoxType
                            Case CB_EXIT
                                If GameState = GameStateEnum.InMenu Then
                                    '//Exit
                                    CloseChoiceBox
                                    UnloadMain
                                ElseIf GameState = GameStateEnum.InGame Then
                                    GettingMap = True
                                    CloseChoiceBox
                                    InitFade 0, FadeIn, 5
                                End If
                            Case CB_CHARDEL
                                Menu_State MENU_STATE_DELCHAR
                                CloseChoiceBox
                            Case CB_RETURNMENU
                                If GameState = GameStateEnum.InMenu Then
                                    CloseChoiceBox
                                    ResetMenu
                                ElseIf GameState = GameStateEnum.InGame Then
                                    GettingMap = True
                                    CloseChoiceBox
                                    InitFade 0, FadeIn, 6
                                End If
                            Case CB_SAVESETTING
                                '//Save setting
                                SaveSettingConfiguration
                                '//Save Key Input
                                For z = 1 To ControlEnum.Control_Count - 1
                                    ControlKey(z).cAsciiKey = TmpKey(z)
                                Next
                                SaveControlKey

                                CloseChoiceBox
                                If GUI(GuiEnum.GUI_OPTION).Visible Then
                                    GuiState GUI_OPTION, False
                                End If
                            Case CB_EVOLVE
                                '//Do Evolve
                                SendEvolvePoke EvolveSelect
                                CloseChoiceBox
                            Case CB_REQUEST
                                '//Accept Request
                                SendRequestState 1
                                CloseChoiceBox
                            Case CB_RELEASE
                                '//release pokemon
                                Dim hasSelected As Boolean
                                For z = 1 To MAX_STORAGE
                                    If IsPokemonSelected(z) Then
                                        hasSelected = True
                                        SendReleasePokemon PokemonCurSlot, z
                                        ClearPokemonSelected z
                                    End If
                                Next z

                                If Not hasSelected Then
                                    SendReleasePokemon ReleaseStorageSlot, ReleaseStorageData
                                    ReleaseStorageSlot = 0
                                    ReleaseStorageData = 0
                                End If
                                CloseChoiceBox
                            Case CB_BUYSLOT
                                '//Buy Slot
                                SendBuyStorageSlot BuySlotType, BuySlotData
                                BuySlotType = 0
                                BuySlotData = 0
                                CloseChoiceBox
                            Case CB_BUYINV
                                '//Buy Inv Slot
                                SendBuyInvSlot BuySlotData
                                BuySlotData = 0
                                CloseChoiceBox
                            Case CB_FLY
                                SendFlyToBadge FlyBadgeSlot
                                FlyBadgeSlot = 0
                                CloseChoiceBox
                            End Select
                        Case ButtonEnum.ChoiceBox_No
                            Select Case ChoiceBoxType
                            Case CB_SAVESETTING
                                CloseChoiceBox
                                If GUI(GuiEnum.GUI_OPTION).Visible Then
                                    GuiState GUI_OPTION, False
                                End If
                            Case CB_REQUEST
                                '//Decline request
                                SendRequestState 2
                                CloseChoiceBox
                            Case CB_RELEASE
                                ReleaseStorageSlot = 0
                                ReleaseStorageData = 0
                                CloseChoiceBox
                            Case CB_BUYSLOT
                                BuySlotType = 0
                                BuySlotData = 0
                                CloseChoiceBox
                            Case CB_BUYINV
                                BuySlotData = 0
                                CloseChoiceBox
                            Case CB_FLY
                                FlyBadgeSlot = 0
                                CloseChoiceBox
                            Case Else
                                CloseChoiceBox
                            End Select
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub CloseChoiceBox()
    GuiState GUI_CHOICEBOX, False
    ChoiceBoxText = vbNullString
    ChoiceBoxType = 0
    EvolveSelect = 0
End Sub

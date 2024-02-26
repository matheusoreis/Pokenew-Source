Attribute VB_Name = "modLoop"
Option Explicit

Public Declare Function GetTickCount Lib "Kernel32" () As Long          '//This is used to get the frame rate.
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long) '//halts thread of execution

'//Handle all looping/time procedure of the application
Public Sub AppLoop()
    Dim i As Long, X As Long
    Dim Tick As Long
    Dim Tmr25 As Long, Tmr100 As Long, Tmr500 As Long, Tmr250 As Long, Tmr1000 As Long, Tmr3000 As Long
    Dim WalkTmr As Long, ChatTmr As Long
    Dim TickFPS As Long, FPS As Long

    Do While AppRunning
        Tick = GetTickCount     '//Set the inital tick

        ' 0.03 milli/second
        If WalkTmr < Tick Then
            If GameState = GameStateEnum.InGame Then
                If Player_HighIndex > 0 Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If Player(i).Map = Player(MyIndex).Map Then
                                Call ProcessPlayerMovement(i)
                                If PlayerPokemon(i).Num > 0 Then
                                    Call ProcessPlayerPokemonMovement(i)
                                End If
                                If PlayerPokemon(i).Init > 0 Then
                                    Call ProcessPokeball(i)
                                End If
                            End If
                        End If
                    Next i
                End If
                If Npc_HighIndex > 0 Then
                    For i = 1 To Npc_HighIndex
                        If MapNpc(i).Num > 0 Then
                            Call ProcessNpcMovement(i)
                        End If
                        If MapNpcPokemon(i).Num > 0 Then
                            Call ProcessNpcPokemonMovement(i)
                        End If
                        If MapNpcPokemon(i).Init > 0 Then
                            Call ProcessNpcPokeball(i)
                        End If
                    Next
                End If
                If Pokemon_HighIndex > 0 Then
                    For i = 1 To Pokemon_HighIndex
                        If MapPokemon(i).Num > 0 Then
                            Call ProcessMapPokemonMovement(i)
                        End If
                        If CatchBall(i).InUsed Then
                            '//Process Ball
                            Call ProcessCatching(i)
                        End If
                    Next
                End If
            End If
            ProcessMyLogic

            If CursorLoadAnim > 7 Then
                CursorLoadAnim = 0
            Else
                CursorLoadAnim = CursorLoadAnim + 1
            End If

            WalkTmr = Tick + 30
        End If

        ' 0.05 milli/second
        If ChatTmr < Tick Then
            If GameState = GameStateEnum.InGame Then
                If ChatScrollTimer + 150 < Tick Then
                    If ChatScrollUp Then
                        ScrollChatBox 0
                        ChatScrollTimer = GetTickCount
                    End If
                    If ChatScrollDown Then
                        ScrollChatBox 1
                        ChatScrollTimer = GetTickCount
                    End If
                End If

                ' Pokedex
                If PokedexScrollTimer + 150 < Tick Then
                    If PokedexScrollUp Then
                        If PokedexViewCount > 0 Then
                            PokedexViewCount = PokedexViewCount - 1
                            PokedexScrollY = (PokedexViewCount * 132) / MaxPokedexViewLine
                            PokedexScrollY = (132 - PokedexScrollY)
                            PokedexScrollTimer = GetTickCount
                        End If
                    End If
                    If PokedexScrollDown Then
                        If PokedexViewCount < MaxPokedexViewLine Then
                            PokedexViewCount = PokedexViewCount + 1
                            PokedexScrollY = (PokedexViewCount * 132) / MaxPokedexViewLine
                            PokedexScrollY = (132 - PokedexScrollY)
                            PokedexScrollTimer = GetTickCount
                        End If
                    End If
                End If

                ' Ranking
                If RankingScrollTimer + 150 < Tick Then
                    If RankingScrollUp Then
                        If RankingViewCount > 0 Then
                            RankingViewCount = RankingViewCount - 1
                            RankingScrollY = (RankingViewCount * RankingScrollLength) / RankingMaxViewLine
                            RankingScrollY = (RankingScrollLength - RankingScrollY)
                            RankingScrollTimer = GetTickCount
                        End If
                    End If
                    If RankingScrollDown Then
                        If RankingViewCount < RankingMaxViewLine Then
                            RankingViewCount = RankingViewCount + 1
                            RankingScrollY = (RankingViewCount * RankingScrollLength) / RankingMaxViewLine
                            RankingScrollY = (RankingScrollLength - RankingScrollY)
                            RankingScrollTimer = GetTickCount
                        End If
                    End If
                End If

                ' Controls
                If ControlScrollTimer + 150 < Tick Then
                    If ControlScrollUp Then
                        If ControlViewCount > 0 Then
                            ControlViewCount = ControlViewCount - 1
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            ControlScrollTimer = GetTickCount
                        End If
                    End If
                    If ControlScrollDown Then
                        If ControlViewCount + (ControlScrollViewLine) < ControlEnum.Control_Count - 1 Then
                            ControlViewCount = ControlViewCount + 1
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            ControlScrollTimer = GetTickCount
                        End If
                    End If
                End If

                ' Virtual Shop
                If VirtualShopScrollTimer + 150 < Tick Then
                    If VirtualShopScrollUp Then
                        If VirtualShopScrollCount > 0 Then
                            VirtualShopScrollCount = VirtualShopScrollCount - 1
                            VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) \ (VirtualShopMaxViewLine \ VirtualShopViewLines)
                            VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)
                            VirtualShopScrollTimer = GetTickCount
                        End If
                    End If
                    If VirtualShopScrollDown Then
                        If VirtualShopScrollCount < (VirtualShopMaxViewLine \ VirtualShopViewLines) Then
                            VirtualShopScrollCount = VirtualShopScrollCount + 1
                            VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) \ (VirtualShopMaxViewLine \ VirtualShopViewLines)
                            VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)
                            VirtualShopScrollTimer = GetTickCount
                        End If
                    End If
                End If
            ElseIf GameState = GameStateEnum.InMenu Then
                '// Pode ser usado InMenu

                ' -->Controls
                If ControlScrollTimer + 150 < Tick Then
                    If ControlScrollUp Then
                        If ControlViewCount > 0 Then
                            ControlViewCount = ControlViewCount - 1
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            ControlScrollTimer = GetTickCount
                        End If
                    End If
                    If ControlScrollDown Then
                        If ControlViewCount + (ControlScrollViewLine) < ControlEnum.Control_Count - 1 Then
                            ControlViewCount = ControlViewCount + 1
                            ControlScrollY = (ControlViewCount * ControlScrollLength) / ControlMaxViewLine
                            ControlScrollY = (ControlScrollLength - ControlScrollY)
                            ControlScrollTimer = GetTickCount
                        End If
                    End If
                End If
            End If

            ChatTmr = Tick + 50
        End If

        If Tmr100 < Tick Then
            If CreditVisible Then
                If CreditState = 0 Then
                    CreditOffset = CreditOffset + 16
                    If CreditOffset >= (Screen_Height - 40) Then
                        CreditOffset = (Screen_Height - 40)

                        If CreditTextCount > 0 Then
                            For i = 0 To CreditTextCount
                                If Credit(i).y > -32 Then
                                    Credit(i).y = Credit(i).y - 1
                                End If
                            Next
                            '//Check if the last text is gone then reset
                            If Credit(CreditTextCount).y <= -32 Then
                                For i = 0 To CreditTextCount
                                    Credit(i).y = Credit(i).StartY
                                Next
                            End If
                        End If
                    End If
                Else
                    CreditOffset = CreditOffset - 16
                    If CreditOffset <= 0 Then
                        CreditState = 0
                        CreditOffset = 0
                        CreditVisible = False
                    End If
                End If
            End If

            Tmr100 = Tick + 10
        End If

        If Tmr25 < Tick Then
            '//Fade
            FadeLogic

            '//Make sure that background is visible
            If GameState = GameStateEnum.InMenu Then
                If MenuState = MenuStateEnum.StateNormal Or MenuState = MenuStateEnum.StateTitleScreen Then
                    BackgroundXOffset = BackgroundXOffset - 1
                    If BackgroundXOffset <= 0 Then
                        BackgroundXOffset = 640    '//Size of the background texture (Need to change if size changed)
                    End If
                End If
            ElseIf GameState = GameStateEnum.InGame Then
                Call CheckKeys
                If GetForegroundWindow() = frmMain.hwnd Then
                    Call CheckInputKeys
                End If

                '//Action
                If CanMoveNow Then
                    Call CheckMovement
                    Call CheckAttack
                End If

                For i = 1 To 255
                    CheckAnimInstance i
                Next

                If ConvoNum > 0 Then
                    If Len(ConvoText) > ConvoDrawTextLen Then
                        ConvoDrawTextLen = ConvoDrawTextLen + 1
                        If Len(ConvoText) <= ConvoDrawTextLen Then
                            ConvoDrawTextLen = Len(ConvoText)
                        End If
                        ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                    End If
                End If
            End If

            Tmr25 = Tick + 25
        End If

        If Tmr1000 < Tick Then
            If GameState = GameStateEnum.InGame Then

                GameSecond = GameSecond + GameSecond_Velocity
                If GameSecond >= 60 Then
                    GameSecond = 0
                    GameMinute = GameMinute + 1
                    If GameMinute >= 60 Then
                        GameMinute = 0
                        GameHour = GameHour + 1
                        If GameHour >= 24 Then
                            GameHour = 0
                        End If
                    End If
                End If

                If GameHour >= 0 And GameHour <= 5 Then     '// Dawn: 1am - 5am
                    DayAndNightARGB = D3DColorARGB(200, 0, 0, 0)
                    ShowLights = True
                    LightAlpha = 255
                ElseIf GameHour >= 6 And GameHour <= 12 Then    '// Morning: 6am - 12pm
                    DayAndNightARGB = D3DColorARGB(20, 0, 0, 0)
                    ShowLights = True
                    LightAlpha = 20
                ElseIf GameHour >= 13 And GameHour <= 17 Then    '// Afternoon: 1pm - 5pm
                    DayAndNightARGB = D3DColorARGB(0, 0, 0, 0)
                    ShowLights = False
                    LightAlpha = 0
                ElseIf GameHour >= 18 And GameHour <= 21 Then    '// Dusk: 5pm - 8pm
                    DayAndNightARGB = D3DColorARGB(50, 0, 0, 0)
                    ShowLights = True
                    LightAlpha = 100
                Else    '// Night: 9pm - 12am
                    DayAndNightARGB = D3DColorARGB(150, 0, 0, 0)
                    ShowLights = True
                    LightAlpha = 255
                End If

                ' Jornada do jogador, tempo jogado
                Player(MyIndex).TimePlay = Player(MyIndex).TimePlay + 1

                ' Processamento do cooldown dos items da bolsa
                For i = 1 To MAX_PLAYER_INV
                    If PlayerInv(i).Num > 0 Then
                        If PlayerInv(i).ItemCooldown > 0 Then    ' 1 seg
                            PlayerInv(i).ItemCooldown = PlayerInv(i).ItemCooldown - 1
                            
                            For X = 1 To MAX_HOTBAR
                                If PlayerInv(i).Num = Player(MyIndex).Hotbar(X).Num Then
                                    Player(MyIndex).Hotbar(X).TmrCooldown = PlayerInv(i).ItemCooldown
                                End If
                            Next X
                            
                        End If
                    End If
                Next i

                ' Evento exp window
                If ExpMultiply > 0 Then
                    If ExpSecs > 0 Then
                        ExpSecs = ExpSecs - 1
                    End If
                End If
            End If

            Tmr1000 = Tick + 1000
        End If

        If Tmr250 < Tick Then
            If GenderAnim = 0 Then
                GenderAnim = 2
            Else
                GenderAnim = 0
            End If

            MapFrameAnim = MapFrameAnim + 1
            If MapFrameAnim > MAX_MAP_FRAME Then
                MapFrameAnim = 0
            End If

            ShinySummaryStep = ShinySummaryStep + 1
            If ShinySummaryStep > 2 Then
                ShinySummaryStep = 0
            End If

            Tmr250 = Tick + 200
        End If

        ' 0.5 milli/second
        If Tick > Tmr500 Then
            '//Check for disconnection

            If TextLine = "|" Then
                TextLine = ""
            Else
                TextLine = "|"
            End If

            If MapAnim = YES Then
                MapAnim = NO
            Else
                MapAnim = YES
            End If

            Tmr500 = Tick + 500
        End If

        ' 3 Seconds
        If Tick > Tmr3000 Then
            Select Case GameState
            Case GameStateEnum.InMenu
                If IsLoggedIn Then
                    Connected = IsConnected

                    '//Update Ping every second
                    If Connected Then
                        If GameSetting.ShowPing Then
                            CheckPing
                        End If
                    Else
                        AddAlert "You got disconnected from the server", White
                        ResetMenu
                    End If
                End If
            Case GameStateEnum.InGame
                If IsLoggedIn Then
                    If IsPlaying(MyIndex) Then
                        Connected = IsConnected

                        '//Update Ping every second
                        If Connected Then
                            CheckPing
                        Else
                            AddAlert "You got disconnected from the server", White
                            ResetMenu
                        End If
                    End If
                End If
            End Select

            Tmr3000 = Tick + 3000
        End If

        '//Make sure that it doesn't need to update frame if the screen is minimized or hidden
        If Not frmMain.WindowState = vbMinimized Then
            Render_Screen           '//Update Frame
        End If
        DoEvents                '//Allow windows time to think; otherwise you'll get into a really tight (and bad) loop...

        '//Prevent from overusing memory
        Do While GetTickCount < Tick + 10
            DoEvents
            Sleep 1
        Loop

        '//Count FPS
        If TickFPS < Tick Then
            TickFPS = Tick + 1000
            GameFps = FPS
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop
End Sub

Private Sub ProcessPlayerMovement(ByVal Index As Long)
    Dim MovementSpeed As Long

    '//Check if player is walking, and if so process moving them over
    With Player(Index)
        If .Action = ACTION_SLIDE Then
            MovementSpeed = 14
        Else
            If .Status = StatusEnum.Paralize Then
                MovementSpeed = 2
            Else
                If .TempSprite = TEMP_SPRITE_GROUP_BIKE Then
                    MovementSpeed = 6
                ElseIf .TempSprite = TEMP_SPRITE_GROUP_MOUNT Then
                    If ShiftKey = True Then
                        If CheckPassivaMount Then
                            MovementSpeed = 10
                        Else
                            MovementSpeed = 7
                        End If
                    Else
                        MovementSpeed = 7
                    End If
                Else
                    MovementSpeed = 5    '//TEMP
                End If
            End If
        End If

        Select Case .Dir
        Case DIR_UP
            .yOffset = .yOffset - MovementSpeed
            If .yOffset < 0 Then .yOffset = 0
        Case DIR_DOWN
            .yOffset = .yOffset + MovementSpeed
            If .yOffset > 0 Then .yOffset = 0
        Case DIR_LEFT
            .xOffset = .xOffset - MovementSpeed
            If .xOffset < 0 Then .xOffset = 0
        Case DIR_RIGHT
            .xOffset = .xOffset + MovementSpeed
            If .xOffset > 0 Then .xOffset = 0
        End Select

        ' Check if completed walking over to the next tile
        If .Moving = YES Then
            If .Dir = DIR_RIGHT Or .Dir = DIR_DOWN Then
                If (.xOffset >= 0) And (.yOffset >= 0) Then
                    .Moving = NO

                    If .TempSprite = TEMP_SPRITE_GROUP_MOUNT Then
                        If .Step = 0 Then
                            .Step = 1
                        ElseIf .Step = 1 Then
                            .Step = 2
                        ElseIf .Step = 2 Then
                            .Step = 3
                        Else
                            .Step = 0
                        End If

                        .IdleTimer = GetTickCount
                        .IdleAnim = 0
                        .IdleFrameTmr = GetTickCount
                    Else
                        If .Step = 0 Then
                            .Step = 2
                        Else
                            .Step = 0
                        End If
                    End If
                End If
            Else
                If (.xOffset <= 0) And (.yOffset <= 0) Then
                    .Moving = NO

                    If .TempSprite = TEMP_SPRITE_GROUP_MOUNT Then
                        If .Step = 0 Then
                            .Step = 1
                        ElseIf .Step = 1 Then
                            .Step = 2
                        ElseIf .Step = 2 Then
                            .Step = 3
                        Else
                            .Step = 0
                        End If

                        .IdleTimer = GetTickCount
                        .IdleAnim = 0
                        .IdleFrameTmr = GetTickCount
                    Else
                        If .Step = 0 Then
                            .Step = 2
                        Else
                            .Step = 0
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    '//Check if npc is walking, and if so process moving them over
    With MapNpc(MapNpcNum)
        MovementSpeed = 4 '//TEMP
        
        Select Case .Dir
            Case DIR_UP
                .yOffset = .yOffset - MovementSpeed
                If .yOffset < 0 Then .yOffset = 0
            Case DIR_DOWN
                .yOffset = .yOffset + MovementSpeed
                If .yOffset > 0 Then .yOffset = 0
            Case DIR_LEFT
                .xOffset = .xOffset - MovementSpeed
                If .xOffset < 0 Then .xOffset = 0
            Case DIR_RIGHT
                .xOffset = .xOffset + MovementSpeed
                If .xOffset > 0 Then .xOffset = 0
        End Select
    
        ' Check if completed walking over to the next tile
        If .Moving = YES Then
            If .Dir = DIR_RIGHT Or .Dir = DIR_DOWN Then
                If (.xOffset >= 0) And (.yOffset >= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                End If
            Else
                If (.xOffset <= 0) And (.yOffset <= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessNpcPokemonMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    '//Check if npc is walking, and if so process moving them over
    With MapNpcPokemon(MapNpcNum)
        MovementSpeed = 4 '//TEMP
        
        Select Case .Dir
            Case DIR_UP
                .yOffset = .yOffset - MovementSpeed
                If .yOffset < 0 Then .yOffset = 0
            Case DIR_DOWN
                .yOffset = .yOffset + MovementSpeed
                If .yOffset > 0 Then .yOffset = 0
            Case DIR_LEFT
                .xOffset = .xOffset - MovementSpeed
                If .xOffset < 0 Then .xOffset = 0
            Case DIR_RIGHT
                .xOffset = .xOffset + MovementSpeed
                If .xOffset > 0 Then .xOffset = 0
        End Select
    
        ' Check if completed walking over to the next tile
        If .Moving = YES Then
            If .Dir = DIR_RIGHT Or .Dir = DIR_DOWN Then
                If (.xOffset >= 0) And (.yOffset >= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                End If
            Else
                If (.xOffset <= 0) And (.yOffset <= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessMapPokemonMovement(ByVal PokemonNum As Long)
Dim MovementSpeed As Long

    '//Check if npc is walking, and if so process moving them over
    With MapPokemon(PokemonNum)
        If .Status = StatusEnum.Paralize Then
            MovementSpeed = 2 'CalculateSpeed(.Stat(StatEnum.Spd))
        Else
            MovementSpeed = 4
        End If
        
        Select Case .Dir
            Case DIR_UP
                .yOffset = .yOffset - MovementSpeed
                If .yOffset < 0 Then .yOffset = 0
            Case DIR_DOWN
                .yOffset = .yOffset + MovementSpeed
                If .yOffset > 0 Then .yOffset = 0
            Case DIR_LEFT
                .xOffset = .xOffset - MovementSpeed
                If .xOffset < 0 Then .xOffset = 0
            Case DIR_RIGHT
                .xOffset = .xOffset + MovementSpeed
                If .xOffset > 0 Then .xOffset = 0
        End Select
    
        ' Check if completed walking over to the next tile
        If .Moving = YES Then
            If .Dir = DIR_RIGHT Or .Dir = DIR_DOWN Then
                If (.xOffset >= 0) And (.yOffset >= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                    .IdleTimer = GetTickCount
                    .IdleAnim = 0
                    .IdleFrameTmr = GetTickCount
                End If
            Else
                If (.xOffset <= 0) And (.yOffset <= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                    .IdleTimer = GetTickCount
                    .IdleAnim = 0
                    .IdleFrameTmr = GetTickCount
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessPlayerPokemonMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    '//Check if npc is walking, and if so process moving them over
    With PlayerPokemon(Index)
        If .Status = StatusEnum.Paralize Then
            MovementSpeed = 2
        Else
            MovementSpeed = CalculateSpeed(GetStatBuff(.Stat(StatEnum.Spd), .StatBuff(StatEnum.Spd)))
        End If
        
        If MovementSpeed <= 2 Then MovementSpeed = 2
        If MovementSpeed >= 14 Then MovementSpeed = 14
        
        Select Case .Dir
            Case DIR_UP
                .yOffset = .yOffset - MovementSpeed
                If .yOffset < 0 Then .yOffset = 0
            Case DIR_DOWN
                .yOffset = .yOffset + MovementSpeed
                If .yOffset > 0 Then .yOffset = 0
            Case DIR_LEFT
                .xOffset = .xOffset - MovementSpeed
                If .xOffset < 0 Then .xOffset = 0
            Case DIR_RIGHT
                .xOffset = .xOffset + MovementSpeed
                If .xOffset > 0 Then .xOffset = 0
        End Select
    
        ' Check if completed walking over to the next tile
        If .Moving = YES Then
            If .Dir = DIR_RIGHT Or .Dir = DIR_DOWN Then
                If (.xOffset >= 0) And (.yOffset >= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                    .IdleTimer = GetTickCount
                    .IdleAnim = 0
                    .IdleFrameTmr = GetTickCount
                End If
            Else
                If (.xOffset <= 0) And (.yOffset <= 0) Then
                    .Moving = NO
                    If .Step = 0 Then
                        .Step = 2
                    Else
                        .Step = 0
                    End If
                    .IdleTimer = GetTickCount
                    .IdleAnim = 0
                    .IdleFrameTmr = GetTickCount
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessCatching(ByVal PokeSlot As Long)
    If CatchBall(PokeSlot).FrameTimer <= GetTickCount Then
        Select Case CatchBall(PokeSlot).State
            Case 0 '//Init
                If CatchBall(PokeSlot).FrameState = 0 Then
                    CatchBall(PokeSlot).Frame = 2
                    CatchBall(PokeSlot).FrameState = 1
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 150
                ElseIf CatchBall(PokeSlot).FrameState = 1 Then
                    CatchBall(PokeSlot).Frame = 0
                    CatchBall(PokeSlot).FrameState = 2 '//Done
                    CatchBall(PokeSlot).FrameTimer = GetTickCount
                End If
            Case 1 '//Shake
                If CatchBall(PokeSlot).FrameState = 0 Then
                    CatchBall(PokeSlot).Frame = 3
                    CatchBall(PokeSlot).FrameState = 1
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 50
                ElseIf CatchBall(PokeSlot).FrameState = 1 Then
                    CatchBall(PokeSlot).Frame = 0
                    CatchBall(PokeSlot).FrameState = 2
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 50
                ElseIf CatchBall(PokeSlot).FrameState = 2 Then
                    CatchBall(PokeSlot).Frame = 4
                    CatchBall(PokeSlot).FrameState = 3
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 50
                ElseIf CatchBall(PokeSlot).FrameState = 3 Then
                    CatchBall(PokeSlot).Frame = 0
                    CatchBall(PokeSlot).FrameState = 4 '//Done
                    CatchBall(PokeSlot).FrameTimer = GetTickCount
                End If
            Case 2 '//Success
                If CatchBall(PokeSlot).FrameState = 0 Then
                    CatchBall(PokeSlot).Frame = 5
                    CatchBall(PokeSlot).FrameState = 1
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 1 Then
                    CatchBall(PokeSlot).Frame = 6
                    CatchBall(PokeSlot).FrameState = 2
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 2 Then
                    CatchBall(PokeSlot).Frame = 7
                    CatchBall(PokeSlot).FrameState = 3
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 3 Then
                    CatchBall(PokeSlot).Frame = 8
                    CatchBall(PokeSlot).FrameState = 4
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 4 Then
                    CatchBall(PokeSlot).Frame = 9
                    CatchBall(PokeSlot).FrameState = 5
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 5 Then
                    CatchBall(PokeSlot).Frame = 10
                    CatchBall(PokeSlot).FrameState = 6
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                ElseIf CatchBall(PokeSlot).FrameState = 6 Then
                    CatchBall(PokeSlot).Frame = 11
                    CatchBall(PokeSlot).FrameState = 7
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 100
                Else
                    CatchBall(PokeSlot).InUsed = False '//Close
                End If
            Case 3 '//Fail
                If CatchBall(PokeSlot).FrameState = 0 Then
                    CatchBall(PokeSlot).Frame = 2
                    CatchBall(PokeSlot).FrameState = 1
                    CatchBall(PokeSlot).FrameTimer = GetTickCount + 150
                Else
                    CatchBall(PokeSlot).InUsed = False '//Close
                End If
        End Select
    End If
End Sub

Private Sub ProcessPokeball(ByVal Index As Long)
    With PlayerPokemon(Index)
        If .Init = YES Then
            If .FrameTimer <= GetTickCount Then
                If .State = 0 Then '//Opening
                    If .FrameState = 0 Then
                        .Frame = 1
                        .FrameState = 1
                        .FrameTimer = GetTickCount + 100
                    ElseIf .FrameState = 1 Then
                        .Frame = 2
                        .FrameState = 2
                        .FrameTimer = GetTickCount + 100
                    Else
                        .Init = NO
                    End If
                Else    '//Closing
                    If .FrameState = 0 Then
                        .Frame = 1
                        .FrameState = 1
                        .FrameTimer = GetTickCount + 100
                    ElseIf .FrameState = 1 Then
                        .Frame = 0
                        .FrameState = 2
                        .FrameTimer = GetTickCount + 100
                    Else
                        .Init = NO
                        ClearPlayerPokemon Index
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessNpcPokeball(ByVal Index As Long)
    With MapNpcPokemon(Index)
        If .Init = YES Then
            If .FrameTimer <= GetTickCount Then
                If .State = 0 Then '//Opening
                    If .FrameState = 0 Then
                        .Frame = 1
                        .FrameState = 1
                        .FrameTimer = GetTickCount + 100
                    ElseIf .FrameState = 1 Then
                        .Frame = 2
                        .FrameState = 2
                        .FrameTimer = GetTickCount + 100
                    Else
                        .Init = NO
                    End If
                Else    '//Closing
                    If .FrameState = 0 Then
                        .Frame = 1
                        .FrameState = 1
                        .FrameTimer = GetTickCount + 100
                    ElseIf .FrameState = 1 Then
                        .Frame = 0
                        .FrameState = 2
                        .FrameTimer = GetTickCount + 100
                    Else
                        .Init = NO
                        ClearNpcPokemon Index
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ProcessMyLogic()
    If GettingMap Then Exit Sub

    If MyIndex > 0 Then
        With Player(MyIndex)
            If .Action = ACTION_SLIDE Then
                If .ActionTmr <= GetTickCount Then
                    If Map.Tile(.X, .y).Attribute = MapAttribute.Slide Then
                        .Action = ACTION_SLIDE
                        .ActionTmr = GetTickCount + 50
                        ForcePlayerMove .Dir
                    Else
                        .Action = 0
                        .ActionTmr = 0
                    End If
                End If
            End If
        End With
    End If
End Sub

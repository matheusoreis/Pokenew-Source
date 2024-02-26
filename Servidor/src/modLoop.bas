Attribute VB_Name = "modLoop"
Option Explicit

Public Declare Function GetTickCount Lib "Kernel32" () As Long          '//This is used to get the frame rate.
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long) '//halts thread of execution

'//Handle all looping/time procedure of the application
Public Sub AppLoop()
Dim i As Long
Dim Tick As Long
Dim Tmr25 As Long, Tmr150 As Long, Tmr500 As Long, Tmr1000 As Long
Dim TickCPS As Long, CPS As Long
Dim LastUpdateSavePlayers As Long

    Do While AppRunning
        Tick = GetTickCount     '//Set the inital tick
        
        If Tick > Tmr25 Then
            '//Update CPS
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            
            '//Player Logic
            UpdatePlayerLogic
            UpdatePokemonLogic
            UpdateMapLogic
            
            Tmr25 = GetTickCount + 25
        End If
        
        If Tick > Tmr500 Then
            '//Check for disconnections every half second
            If Player_HighIndex > 0 Then
                For i = 1 To Player_HighIndex
                    '//Check for socket status
                    If frmServer.Socket(i).State > sckConnected Then
                        Call CloseSocket(i)
                    End If
                Next
            End If
            
            Tmr500 = GetTickCount + 500
        End If
        
        If Tick > Tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            '//Loop Events 1 sec.
            Call EventsLoop
            '//Loop Player Time Played.
            Call PlayerPlayTime
            '//Loop Game Time Base
            Call GameTimeLoop
            
            Tmr1000 = GetTickCount + 1000
        End If
        
        DoEvents                '//Allow windows time to think; otherwise you'll get into a really tight (and bad) loop...
        
        '//Prevent from overusing memory
        If Not CPSUnlock Then
            Do While GetTickCount < Tick + 10
                DoEvents
                Sleep 1
            Loop
        End If
        
        '//Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub GameTimeLoop()
'frmServer.lblGameTime = "Game Time: " & TimeSerial(GameHour, GameMinute, GameSecs)
    GameSecs = GameSecs + GameSecs_Velocity
    If GameSecs >= 60 Then
        GameSecs = 0
        GameMinute = GameMinute + 1
        If GameMinute >= 60 Then
            GameMinute = 0
            GameHour = GameHour + 1
            If GameHour >= 24 Then
                GameHour = 0
                SendClientTimeToAll
            End If
            SendClientTimeToAll
        End If
    End If
End Sub

Private Sub PlayerPlayTime()
    Dim i As Integer
    ' Contabiliza o tempo do jogador jogado.
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                Player(i, TempPlayer(i).UseChar).TimePlay = Player(i, TempPlayer(i).UseChar).TimePlay + 1
            End If
        End If
    Next i

End Sub

Private Sub EventsLoop()
    Dim i As Integer

    With EventExp
        ' Evento exp por tempo
        If .ExpEvent Then
            If .ExpSecs > 0 Then
                .ExpSecs = .ExpSecs - 1
            Else
                .ExpEvent = False
                frmServer.cmdExp.Caption = "Ativar"
                frmServer.scrlExp.Enabled = True
                frmServer.txtExpHour.Enabled = True

                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then

                        If TempPlayer(i).UseChar > 0 Then
                            If Player(i, TempPlayer(i).UseChar).Access <= 0 Then
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Event Exp Desactivated.", BrightRed
                                Case LANG_EN: AddAlert i, "Event Exp Desactivated.", BrightRed
                                Case LANG_ES: AddAlert i, "Event Exp Desactivated.", BrightRed
                                End Select
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End With
End Sub


Private Sub UpdatePlayerLogic()
    Dim i As Long
    Dim Tick As Long
    Dim Value As Long
    Dim RandomNumber As Long

    For i = 1 To Player_HighIndex
        Tick = GetTickCount

        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                '//Check Player Pokemon Logic
                If PlayerPokemon(i).Num > 0 Then
                    '//Moveset logic
                    If PlayerPokemon(i).QueueMove > 0 Then
                        '//Check Cast Time
                        If PlayerPokemon(i).MoveCastTime <= GetTickCount Then
                            If PlayerPokemon(i).MoveAttackCount < PokemonMove(PlayerPokemon(i).QueueMove).AmountOfAttack Then
                                '//Check For Duration
                                If PlayerPokemon(i).MoveDuration >= Tick Then
                                    If PlayerPokemon(i).MoveInterval <= Tick Then
                                        ProcessPlayerMove i, PlayerPokemon(i).QueueMove
                                        PlayerPokemon(i).MoveInterval = GetTickCount + (PokemonMove(PlayerPokemon(i).QueueMove).Interval)
                                        PlayerPokemon(i).MoveAttackCount = PlayerPokemon(i).MoveAttackCount + 1
                                    End If
                                Else
                                    '//InitCooldown
                                    If PlayerPokemon(i).slot > 0 And PlayerPokemon(i).QueueMoveSlot > 0 Then
                                        PlayerPokemons(i).Data(PlayerPokemon(i).slot).Moveset(PlayerPokemon(i).QueueMoveSlot).CD = GetTickCount
                                    End If
                                    '//Clear Queue Move
                                    PlayerPokemon(i).QueueMove = 0
                                    PlayerPokemon(i).QueueMoveSlot = 0
                                    PlayerPokemon(i).MoveDuration = 0
                                    PlayerPokemon(i).MoveCastTime = 0
                                    PlayerPokemon(i).MoveInterval = 0
                                End If
                            Else
                                '//InitCooldown
                                If PlayerPokemon(i).slot > 0 And PlayerPokemon(i).QueueMoveSlot > 0 Then
                                    PlayerPokemons(i).Data(PlayerPokemon(i).slot).Moveset(PlayerPokemon(i).QueueMoveSlot).CD = GetTickCount
                                End If
                                '//Clear Queue Move
                                PlayerPokemon(i).QueueMove = 0
                                PlayerPokemon(i).QueueMoveSlot = 0
                                PlayerPokemon(i).MoveDuration = 0
                                PlayerPokemon(i).MoveCastTime = 0
                                PlayerPokemon(i).MoveInterval = 0
                            End If
                        End If
                    End If
                End If

                '//Check if trying to catch a pokemon
                If TempPlayer(i).TmpCatchPokeNum > 0 Then
                    If TempPlayer(i).TmpCatchTimer <= GetTickCount Then
                        If TempPlayer(i).TmpCatchTries < 3 Then
                            If TempPlayer(i).TmpCatchValue > 0 Then
                                Value = 1048560 / Sqr(Sqr(16711680 / TempPlayer(i).TmpCatchValue))
                                RandomNumber = Random(0, 65535)
                                If RandomNumber > Value Then
                                    '//it Broke
                                    MapPokemon(TempPlayer(i).TmpCatchPokeNum).InCatch = NO
                                    MapPokemon(TempPlayer(i).TmpCatchPokeNum).targetType = TARGET_TYPE_PLAYER
                                    MapPokemon(TempPlayer(i).TmpCatchPokeNum).TargetIndex = i
                                    SendMapPokemonCatchState MapPokemon(TempPlayer(i).TmpCatchPokeNum).Map, TempPlayer(i).TmpCatchPokeNum, MapPokemon(TempPlayer(i).TmpCatchPokeNum).x, MapPokemon(TempPlayer(i).TmpCatchPokeNum).Y, 3, TempPlayer(i).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                    TempPlayer(i).TmpCatchPokeNum = 0
                                    TempPlayer(i).TmpCatchTimer = 0
                                    TempPlayer(i).TmpCatchTries = 0
                                    TempPlayer(i).TmpCatchValue = 0
                                    TempPlayer(i).TmpCatchUseBall = 0
                                    Select Case TempPlayer(i).CurLanguage
                                    Case LANG_PT: AddAlert i, "Your Pokeball broke", White
                                    Case LANG_EN: AddAlert i, "Your Pokeball broke", White
                                    Case LANG_ES: AddAlert i, "Your Pokeball broke", White
                                    End Select
                                Else
                                    '//Continue
                                    TempPlayer(i).TmpCatchTries = TempPlayer(i).TmpCatchTries + 1
                                    TempPlayer(i).TmpCatchTimer = GetTickCount + 500
                                    '//Do Animation
                                    SendMapPokemonCatchState MapPokemon(TempPlayer(i).TmpCatchPokeNum).Map, TempPlayer(i).TmpCatchPokeNum, MapPokemon(TempPlayer(i).TmpCatchPokeNum).x, MapPokemon(TempPlayer(i).TmpCatchPokeNum).Y, 1, TempPlayer(i).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                    Select Case TempPlayer(i).CurLanguage
                                    Case LANG_PT: AddAlert i, "The pokeball shaked...", White
                                    Case LANG_EN: AddAlert i, "The pokeball shaked...", White
                                    Case LANG_ES: AddAlert i, "The pokeball shaked...", White
                                    End Select
                                End If
                            Else
                                '//it Broke
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).InCatch = NO
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).targetType = TARGET_TYPE_PLAYER
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).TargetIndex = i
                                SendMapPokemonCatchState MapPokemon(TempPlayer(i).TmpCatchPokeNum).Map, TempPlayer(i).TmpCatchPokeNum, MapPokemon(TempPlayer(i).TmpCatchPokeNum).x, MapPokemon(TempPlayer(i).TmpCatchPokeNum).Y, 3, TempPlayer(i).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                TempPlayer(i).TmpCatchPokeNum = 0
                                TempPlayer(i).TmpCatchTimer = 0
                                TempPlayer(i).TmpCatchTries = 0
                                TempPlayer(i).TmpCatchValue = 0
                                TempPlayer(i).TmpCatchUseBall = 0
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Your Pokeball broke", White
                                Case LANG_EN: AddAlert i, "Your Pokeball broke", White
                                Case LANG_ES: AddAlert i, "Your Pokeball broke", White
                                End Select
                            End If
                        Else
                            '//Success
                            If CountFreePokemonSlot(i) < 5 Then
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Warning: You only have few slot left for pokemon", White
                                Case LANG_EN: AddAlert i, "Warning: You only have few slot left for pokemon", White
                                Case LANG_ES: AddAlert i, "Warning: You only have few slot left for pokemon", White
                                End Select
                            End If

                            '//Give Player Pokemon
                            If CatchMapPokemonData(i, TempPlayer(i).TmpCatchPokeNum, TempPlayer(i).TmpCatchUseBall) Then
                                '//Success
                                '//Clear map pokemon
                                SendMapPokemonCatchState MapPokemon(TempPlayer(i).TmpCatchPokeNum).Map, TempPlayer(i).TmpCatchPokeNum, MapPokemon(TempPlayer(i).TmpCatchPokeNum).x, MapPokemon(TempPlayer(i).TmpCatchPokeNum).Y, 2, TempPlayer(i).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Congratiolations! You have captured the pokemon...", White
                                Case LANG_EN: AddAlert i, "Congratiolations! You have captured the pokemon...", White
                                Case LANG_ES: AddAlert i, "Congratiolations! You have captured the pokemon...", White
                                End Select
                                ClearMapPokemon TempPlayer(i).TmpCatchPokeNum

                                TempPlayer(i).TmpCatchPokeNum = 0
                                TempPlayer(i).TmpCatchTimer = 0
                                TempPlayer(i).TmpCatchTries = 0
                                TempPlayer(i).TmpCatchValue = 0
                                TempPlayer(i).TmpCatchUseBall = 0
                            Else
                                '//Broke
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).InCatch = NO
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).targetType = TARGET_TYPE_PLAYER
                                MapPokemon(TempPlayer(i).TmpCatchPokeNum).TargetIndex = i
                                SendMapPokemonCatchState MapPokemon(TempPlayer(i).TmpCatchPokeNum).Map, TempPlayer(i).TmpCatchPokeNum, MapPokemon(TempPlayer(i).TmpCatchPokeNum).x, MapPokemon(TempPlayer(i).TmpCatchPokeNum).Y, 3, TempPlayer(i).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                                TempPlayer(i).TmpCatchPokeNum = 0
                                TempPlayer(i).TmpCatchTimer = 0
                                TempPlayer(i).TmpCatchTries = 0
                                TempPlayer(i).TmpCatchValue = 0
                                TempPlayer(i).TmpCatchUseBall = 0
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Your Pokeball broke", White
                                Case LANG_EN: AddAlert i, "Your Pokeball broke", White
                                Case LANG_ES: AddAlert i, "Your Pokeball broke", White
                                End Select
                            End If
                        End If
                    End If
                End If

                '//Duel
                If TempPlayer(i).InDuel > 0 Or TempPlayer(i).InNpcDuel > 0 Then
                    '//Starting
                    If TempPlayer(i).DuelTime > 0 Then
                        If TempPlayer(i).DuelTimeTmr <= GetTickCount Then
                            TempPlayer(i).DuelTime = TempPlayer(i).DuelTime - 1
                            With Player(i, TempPlayer(i).UseChar)
                                If TempPlayer(i).DuelTime > 0 Then
                                    SendActionMsg .Map, TempPlayer(i).DuelTime, .x * 32, .Y * 32, White
                                End If

                                If TempPlayer(i).DuelTime <= 0 Then
                                    '//Init Battle
                                    SendActionMsg .Map, "Start!", .x * 32, .Y * 32, White
                                    TempPlayer(i).DuelTimeTmr = GetTickCount + 25000
                                    TempPlayer(i).WarningTimer = GetTickCount + 5000
                                    If PlayerPokemon(i).Num <= 0 Then
                                        Select Case TempPlayer(i).CurLanguage
                                        Case LANG_PT: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        Case LANG_EN: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        Case LANG_ES: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        End Select
                                    End If
                                Else
                                    TempPlayer(i).DuelTimeTmr = GetTickCount + 1000
                                End If
                            End With
                        End If
                    Else    '//Current
                        If PlayerPokemon(i).Num <= 0 Then
                            If TempPlayer(i).DuelTimeTmr <= GetTickCount Then
                                '//PvP
                                If TempPlayer(i).InDuel > 0 Then
                                    If IsPlaying(TempPlayer(i).InDuel) Then
                                        If TempPlayer(TempPlayer(i).InDuel).UseChar > 0 Then
                                            If TempPlayer(TempPlayer(i).InDuel).InDuel = i Then
                                                '//Check result
                                                If PlayerPokemon(TempPlayer(i).InDuel).Num > 0 Then
                                                    '//Lose
                                                    SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Lose!", Player(i, TempPlayer(i).UseChar).x * 32, Player(i, TempPlayer(i).UseChar).Y * 32, White
                                                    SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Win!", Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).x * 32, Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Y * 32, White
                                                    Player(i, TempPlayer(i).UseChar).Lose = Player(i, TempPlayer(i).UseChar).Lose + 1
                                                    Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Win = Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Win + 1
                                                    SendPlayerPvP (i)
                                                    SendPlayerPvP (TempPlayer(i).InDuel)
                                                Else
                                                    '//Draw
                                                    SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Tie!", Player(i, TempPlayer(i).UseChar).x * 32, Player(i, TempPlayer(i).UseChar).Y * 32, White
                                                    SendActionMsg Player(i, TempPlayer(i).UseChar).Map, "Tie!", Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).x * 32, Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Y * 32, White
                                                    Player(i, TempPlayer(i).UseChar).Tie = Player(i, TempPlayer(i).UseChar).Tie + 1
                                                    Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Tie = Player(TempPlayer(i).InDuel, TempPlayer(TempPlayer(i).InDuel).UseChar).Tie + 1
                                                    SendPlayerPvP (i)
                                                    SendPlayerPvP (TempPlayer(i).InDuel)
                                                End If
                                                TempPlayer(TempPlayer(i).InDuel).InDuel = 0
                                                TempPlayer(TempPlayer(i).InDuel).DuelTime = 0
                                                TempPlayer(TempPlayer(i).InDuel).DuelTimeTmr = 0
                                                TempPlayer(TempPlayer(i).InDuel).WarningTimer = 0
                                                TempPlayer(TempPlayer(i).InDuel).PlayerRequest = 0
                                                TempPlayer(TempPlayer(i).InDuel).RequestType = 0
                                                SendRequest TempPlayer(i).InDuel
                                            End If
                                        End If
                                    End If
                                    TempPlayer(i).InDuel = 0
                                    TempPlayer(i).DuelTime = 0
                                    TempPlayer(i).DuelTimeTmr = 0
                                    TempPlayer(i).WarningTimer = 0
                                    TempPlayer(i).PlayerRequest = 0
                                    TempPlayer(i).RequestType = 0
                                    SendRequest i
                                End If
                                If TempPlayer(i).InNpcDuel > 0 Then
                                    '//Adicionado a apenas um método.
                                    PlayerLoseToNpc i, TempPlayer(i).InNpcDuel
                                End If
                            Else
                                If TempPlayer(i).WarningTimer <= GetTickCount Then
                                    If PlayerPokemon(i).Num <= 0 Then
                                        Select Case TempPlayer(i).CurLanguage
                                        Case LANG_PT: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        Case LANG_EN: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        Case LANG_ES: AddAlert i, "You have " & Round((TempPlayer(i).DuelTimeTmr - GetTickCount) / 1000, 0) & "sec/s to release your pokemon, otherwise you will lose the duel", White
                                        End Select
                                    End If
                                    TempPlayer(i).WarningTimer = GetTickCount + 5000
                                End If
                            End If
                        Else
                            TempPlayer(i).DuelTimeTmr = GetTickCount + 25000
                            TempPlayer(i).WarningTimer = GetTickCount
                        End If
                    End If
                End If

                '//Action
                'If PlayerPokemon(i).Num <= 0 Then
                '    With Player(i, TempPlayer(i).UseChar)
                '        If .Action = ACTION_SLIDE Then
                '            If GetTickCount > .ActionTmr Then
                '                If Map(.Map).Tile(.x, .y).Attribute = MapAttribute.Slide Then
                '                    .Action = ACTION_SLIDE
                '                    .ActionTmr = GetTickCount + 100
                '                    ForcePlayerMove i, .Dir
                '                Else
                '                    .Action = 0
                '                    SendPlayerAction i
                '                    .ActionTmr = 0
                '                End If
                '            End If
                '        End If
                '    End With
                'End If
            End If
        End If
    Next
    DoEvents
End Sub

Public Sub UpdateSavePlayers()
Dim i As Long

    AddLog "Server saving players..."
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                SavePlayerDatas i
            End If
        End If
        DoEvents
    Next
    AddLog "Server saving complete..."
End Sub

Private Sub UpdateMapLogic()
Dim MapNum As Long
Dim MapNpcNum As Long, NpcNum As Long
Dim randNum As Long
Dim DidWalk As Boolean
Dim Target As Long, TargetX As Long, TargetY As Long
Dim NearTarget As Boolean
Dim x As Byte, DidSpawn As Boolean
Dim DuelIndex As Long
Dim Exiting As Boolean
Dim QueueMove As Long

    For MapNum = 1 To MAX_MAP
        Exiting = False
        If PlayerOnMap(MapNum) = YES Then
            '//Map Npc
            For MapNpcNum = 1 To MAX_MAP_NPC
                NpcNum = MapNpc(MapNum, MapNpcNum).Num
                
                '//Verifica se o npc pode spawnar ou despawnar
                Call CheckSpawnNpc(MapNum, MapNpcNum)
                
                If NpcNum > 0 Then
                    ' ******************
                    ' ** Npc Movement **
                    ' ******************
                    If Map(MapNum).Npc(MapNpcNum) > 0 And NpcNum > 0 Then
                        If MapNpc(MapNum, MapNpcNum).InBattle <= 0 Then
                            '//Check behaviour if it can move
                            If Npc(NpcNum).Behaviour = BEHAVIOUR_MOVE Then
                                If MapNpc(MapNum, MapNpcNum).MoveTmr <= GetTickCount Then
                                    '//Randomize number to prevent continues movement
                                    randNum = Int(Rnd * 15)
                                    
                                    If randNum = 1 Then
                                        '//Randomize number for direction
                                        randNum = Int(Rnd * 4)
                                        
                                        '//Process Move
                                        NpcMove MapNum, MapNpcNum, randNum
                                        MapNpc(MapNum, MapNpcNum).MoveTmr = GetTickCount + 1000
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '//Target
                    If MapNpc(MapNum, MapNpcNum).InBattle > 0 Then
                        If MapNpcPokemon(MapNum, MapNpcNum).Num > 0 Then
                        
                        Else
                            If MapNpc(MapNum, MapNpcNum).FaintWaitTimer <= GetTickCount Then
                                '//Try To Spawn Another Pokemon
                                For x = 1 To MAX_PLAYER_POKEMON
                                    If MapNpc(MapNum, MapNpcNum).PokemonAlive(x) = YES Then
                                        If Npc(MapNpc(MapNum, MapNpcNum).Num).PokemonNum(x) > 0 Then
                                            MapNpc(MapNum, MapNpcNum).CurPokemon = x
                                            SpawnNpcPokemon MapNum, MapNpcNum, x
                                            DidSpawn = True
                                        End If
                                    End If
                                Next
                                If Not DidSpawn Then
                                    '//Player Win
                                    DuelIndex = MapNpc(MapNum, MapNpcNum).InBattle
                                    If IsPlaying(DuelIndex) Then
                                        If TempPlayer(DuelIndex).UseChar > 0 Then
                                            '//Adicionado a apenas um método.
                                            PlayerWinToNpc DuelIndex, MapNpcNum
                                        End If
                                    End If
                                    MapNpc(MapNum, MapNpcNum).InBattle = 0
                                End If
                            End If
                        End If
                        
                        '// Check if player is online
                        Target = MapNpc(MapNum, MapNpcNum).InBattle
                        If IsPlaying(Target) Then
                            If TempPlayer(Target).UseChar > 0 Then
                                If TempPlayer(Target).DuelTime <= 0 Then
                                    If PlayerPokemon(Target).Num > 0 Then
                                        TargetX = PlayerPokemon(Target).x
                                        TargetY = PlayerPokemon(Target).Y
                                    Else
                                        Target = 0
                                        TargetX = 0
                                        TargetY = 0
                                    End If
                                Else
                                    Target = 0
                                    TargetX = 0
                                    TargetY = 0
                                End If
                            Else
                                Target = 0
                                TargetX = 0
                                TargetY = 0
                                MapNpc(MapNum, MapNpcNum).InBattle = 0
                                NpcPokemonCallBack MapNum, MapNpcNum
                                '//ToDo: despawn mappokemon
                            End If
                        Else
                            Target = 0
                            TargetX = 0
                            TargetY = 0
                            MapNpc(MapNum, MapNpcNum).InBattle = 0
                            NpcPokemonCallBack MapNum, MapNpcNum
                            '//ToDo: despawn mappokemon
                        End If
                    End If
                    
                    ' ***************
                    ' ** Attacking **
                    ' ***************
                    If Target > 0 Then
                        If MapNpcPokemon(MapNum, MapNpcNum).AtkTmr <= GetTickCount Then
                            '//Check Direction
                            If MapNpcPokemon(MapNum, MapNpcNum).QueueMove <= 0 Then
                                'Select Case MapPokemon(MapPokeNum).Dir
                                '    Case DIR_LEFT: If MapPokemon(MapPokeNum).x - 1 = TargetX And MapPokemon(MapPokeNum).y = TargetY Then NearTarget = True
                                '    Case DIR_RIGHT: If MapPokemon(MapPokeNum).x + 1 = TargetX And MapPokemon(MapPokeNum).y = TargetY Then NearTarget = True
                                '    Case DIR_UP: If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).y - 1 = TargetY Then NearTarget = True
                                '    Case DIR_DOWN: If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).y + 1 = TargetY Then NearTarget = True
                                'End Select
                                NearTarget = False
                                If IsOnAoERange(4, MapNpcPokemon(MapNum, MapNpcNum).x, MapNpcPokemon(MapNum, MapNpcNum).Y, TargetX, TargetY) Then NearTarget = True
                                
                                If NearTarget Then
                                    '//Check Random use of move
                                    randNum = Random(1, 4)
                                    If randNum >= 1 And randNum <= 4 Then
                                        '//Check if moveset is available
                                        If MapNpcPokemon(MapNum, MapNpcNum).Moveset(randNum).Num > 0 Then
                                            If MapNpcPokemon(MapNum, MapNpcNum).Moveset(randNum).CD <= GetTickCount Then
                                                NpcPokemonCastMove MapNum, MapNpcNum, MapNpcPokemon(MapNum, MapNpcNum).Moveset(randNum).Num, randNum, False
                                                MapNpcPokemon(MapNum, MapNpcNum).AtkTmr = GetTickCount + 1000
                                                'SendNpcAttack MapNum, MapPokeNum
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        QueueMove = MapNpcPokemon(MapNum, MapNpcNum).QueueMove
                        If QueueMove > 0 Then
                            '//Check Cast Time
                            If MapNpcPokemon(MapNum, MapNpcNum).MoveCastTime <= GetTickCount Then
                                '//Check For Duration
                                If MapNpcPokemon(MapNum, MapNpcNum).MoveAttackCount < PokemonMove(QueueMove).AmountOfAttack Then
                                    If MapNpcPokemon(MapNum, MapNpcNum).MoveDuration >= GetTickCount Then
                                        If MapNpcPokemon(MapNum, MapNpcNum).MoveInterval <= GetTickCount Then
                                            ProcessNpcPokemonMove MapNum, MapNpcNum, QueueMove
                                            If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then
                                                Exiting = True
                                            End If
                                            MapNpcPokemon(MapNum, MapNpcNum).MoveInterval = GetTickCount + PokemonMove(QueueMove).Interval
                                            MapNpcPokemon(MapNum, MapNpcNum).MoveAttackCount = MapNpcPokemon(MapNum, MapNpcNum).MoveAttackCount + 1
                                        End If
                                    Else
                                        If MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot > 0 Then
                                            MapNpcPokemon(MapNum, MapNpcNum).Moveset(MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot).CD = GetTickCount
                                        End If
                                        '//Clear Queue Move
                                        MapNpcPokemon(MapNum, MapNpcNum).QueueMove = 0
                                        MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot = 0
                                        MapNpcPokemon(MapNum, MapNpcNum).MoveDuration = 0
                                        MapNpcPokemon(MapNum, MapNpcNum).MoveCastTime = 0
                                        MapNpcPokemon(MapNum, MapNpcNum).MoveInterval = 0
                                    End If
                                Else
                                    If MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot > 0 Then
                                        MapNpcPokemon(MapNum, MapNpcNum).Moveset(MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot).CD = GetTickCount
                                    End If
                                    '//Clear Queue Move
                                    MapNpcPokemon(MapNum, MapNpcNum).QueueMove = 0
                                    MapNpcPokemon(MapNum, MapNpcNum).QueueMoveSlot = 0
                                    MapNpcPokemon(MapNum, MapNpcNum).MoveDuration = 0
                                    MapNpcPokemon(MapNum, MapNpcNum).MoveCastTime = 0
                                    MapNpcPokemon(MapNum, MapNpcNum).MoveInterval = 0
                                End If
                            End If
                        End If
                    End If
                    
                    If Not Exiting Then
                        '//Pokemon
                        With MapNpcPokemon(MapNum, MapNpcNum)
                            If .Num > 0 Then
                                ' **************
                                ' ** Movement **
                                ' **************
                                If Target > 0 Then
                                    If .MoveTmr <= GetTickCount Then
                                        randNum = Int(Rnd * 3)
                                        DidWalk = False
                
                                        '//CheckMovement
                                        Select Case randNum
                                            Case 0
                                                If .Y > TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_UP) Then DidWalk = True
                                                End If
                                                If .Y < TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_DOWN) Then DidWalk = True
                                                End If
                                                If .x > TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_LEFT) Then DidWalk = True
                                                End If
                                                If .x < TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_RIGHT) Then DidWalk = True
                                                End If
                                            Case 1
                                                If .x < TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_RIGHT) Then DidWalk = True
                                                End If
                                                If .x > TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_LEFT) Then DidWalk = True
                                                End If
                                                If .Y < TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_DOWN) Then DidWalk = True
                                                End If
                                                If .Y > TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_UP) Then DidWalk = True
                                                End If
                                            Case 2
                                                If .Y < TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_DOWN) Then DidWalk = True
                                                End If
                                                If .Y > TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_UP) Then DidWalk = True
                                                End If
                                                If .x < TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_RIGHT) Then DidWalk = True
                                                End If
                                                If .x > TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_LEFT) Then DidWalk = True
                                                End If
                                            Case 3
                                                If .x > TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_LEFT) Then DidWalk = True
                                                End If
                                                If .x < TargetX And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_RIGHT) Then DidWalk = True
                                                End If
                                                If .Y > TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_UP) Then DidWalk = True
                                                End If
                                                If .Y < TargetY And Not DidWalk Then
                                                    If NpcPokemonMove(MapNum, MapNpcNum, DIR_DOWN) Then DidWalk = True
                                                End If
                                        End Select
                                                    
                                        '//Check Direction
                                        If Not DidWalk Then
                                            If .x - 1 = TargetX And .Y = TargetY Then
                                                If Not .Dir = DIR_LEFT Then Call NpcPokemonDir(MapNum, MapNpcNum, DIR_LEFT)
                                                DidWalk = True
                                            End If
                                            If .x + 1 = TargetX And .Y = TargetY Then
                                                If Not .Dir = DIR_RIGHT Then Call NpcPokemonDir(MapNum, MapNpcNum, DIR_RIGHT)
                                                DidWalk = True
                                            End If
                                            If .x = TargetX And .Y - 1 = TargetY Then
                                                If Not .Dir = DIR_UP Then Call NpcPokemonDir(MapNum, MapNpcNum, DIR_UP)
                                                DidWalk = True
                                            End If
                                            If .x = TargetX And .Y + 1 = TargetY Then
                                                If Not .Dir = DIR_DOWN Then Call NpcPokemonDir(MapNum, MapNpcNum, DIR_DOWN)
                                                DidWalk = True
                                            End If
                                                        
                                            If Not DidWalk Then
                                                '//Randomize number to prevent continues movement
                                                randNum = Int(Rnd * 15)
                                                            
                                                If randNum = 1 Then
                                                    '//Randomize number for direction
                                                    randNum = Int(Rnd * 5)
        
                                                    '//Process Move
                                                    If NpcPokemonMove(MapNum, MapNpcNum, randNum) Then
                                                        '//Do Nothing
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Private Sub UpdatePokemonLogic()
    Dim MapPokeNum As Long
    Dim MapNum As Long
    Dim TargetX As Long, TargetY As Long, targetType As Byte, Target As Long
    Dim DidWalk As Boolean
    Dim randNum As Long
    Dim NearTarget As Boolean
    Dim Tick As Long
    Dim i As Long
    Dim onSightRange As Long, onSightDistanceX As Long, onSightDistanceY As Long
    Dim Exiting As Boolean
    Dim QueueMove As Long

    '//Check error
    If Pokemon_HighIndex <= 0 Then Exit Sub

    For MapPokeNum = 1 To Pokemon_HighIndex
        Exiting = False
        Tick = GetTickCount

        ' *************
        ' ** Respawn **
        ' *************
        '//Check If dead
        If MapPokemon(MapPokeNum).Num <= 0 Then
            '//Check if does exist
            If MapPokemon(MapPokeNum).PokemonIndex > 0 Then
                If Spawn(MapPokeNum).Fishing = NO Then    ' Não spawna pokemon de pesca automaticamente
                    If MapPokemon(MapPokeNum).Respawn <= GetTickCount Then
                        '//Spawn pokemon
                        SpawnMapPokemon MapPokeNum
                    End If
                End If
            End If
        End If

        '//If Alive, Do event
        If MapPokemon(MapPokeNum).Num > 0 And MapPokemon(MapPokeNum).InCatch = NO Then
            MapNum = MapPokemon(MapPokeNum).Map

            If PlayerOnMap(MapNum) = YES Then
                ' ********************
                ' ** Getting Target **
                ' ********************
                targetType = 0
                Target = 0
                TargetX = 0
                TargetY = 0

                If MapPokemon(MapPokeNum).TargetIndex > 0 Then
                    targetType = MapPokemon(MapPokeNum).targetType
                    Target = MapPokemon(MapPokeNum).TargetIndex
                    TargetX = MapPokemon(MapPokeNum).x
                    TargetY = MapPokemon(MapPokeNum).Y

                    '//Check Target
                    Select Case targetType
                    Case TARGET_TYPE_PLAYER
                        '//Check
                        If IsPlaying(Target) Then
                            If TempPlayer(Target).UseChar > 0 Then
                                If Player(Target, TempPlayer(Target).UseChar).Map = MapNum Then
                                    '//Check if it have pokemon
                                    If PlayerPokemon(Target).Num > 0 Then
                                        '//Switch Target
                                        MapPokemon(MapPokeNum).targetType = TARGET_TYPE_PLAYERPOKEMON
                                    Else
                                        '//Follow
                                        TargetX = Player(Target, TempPlayer(Target).UseChar).x
                                        TargetY = Player(Target, TempPlayer(Target).UseChar).Y
                                        onSightRange = Pokemon(MapPokemon(MapPokeNum).Num).Range + 4
                                        onSightDistanceX = MapPokemon(MapPokeNum).x - TargetX
                                        onSightDistanceY = MapPokemon(MapPokeNum).Y - TargetY
                                        If onSightDistanceX <= onSightRange And onSightDistanceY <= onSightRange Then

                                        Else
                                            '//Fish system
                                            If Spawn(MapPokeNum).Fishing = YES Then
                                                ClearMapPokemon MapPokeNum
                                            End If
                                            ' Lost Target
                                            MapPokemon(MapPokeNum).targetType = 0
                                            MapPokemon(MapPokeNum).TargetIndex = 0
                                            targetType = 0
                                            Target = 0
                                            TargetX = 0
                                            TargetY = 0
                                        End If
                                    End If
                                Else
                                    '//Fish system
                                    If Spawn(MapPokeNum).Fishing = YES Then
                                        ClearMapPokemon MapPokeNum
                                    End If
                                    ' Lost Target
                                    MapPokemon(MapPokeNum).targetType = 0
                                    MapPokemon(MapPokeNum).TargetIndex = 0
                                    targetType = 0
                                    Target = 0
                                    TargetX = 0
                                    TargetY = 0
                                End If
                            Else
                                '//Fish system
                                If Spawn(MapPokeNum).Fishing = YES Then
                                    ClearMapPokemon MapPokeNum
                                End If
                                ' Lost Target
                                MapPokemon(MapPokeNum).targetType = 0
                                MapPokemon(MapPokeNum).TargetIndex = 0
                                targetType = 0
                                Target = 0
                                TargetX = 0
                                TargetY = 0
                            End If
                        Else
                            '//Fish system
                            If Spawn(MapPokeNum).Fishing = YES Then
                                ClearMapPokemon MapPokeNum
                            End If
                            ' Lost Target
                            MapPokemon(MapPokeNum).targetType = 0
                            MapPokemon(MapPokeNum).TargetIndex = 0
                            targetType = 0
                            Target = 0
                            TargetX = 0
                            TargetY = 0
                        End If
                    Case TARGET_TYPE_PLAYERPOKEMON
                        '//Check if it have pokemon
                        If PlayerPokemon(Target).Num > 0 Then
                            '//Follow
                            TargetX = PlayerPokemon(Target).x
                            TargetY = PlayerPokemon(Target).Y
                            onSightRange = Pokemon(MapPokemon(MapPokeNum).Num).Range + 4
                            onSightDistanceX = MapPokemon(MapPokeNum).x - TargetX
                            onSightDistanceY = MapPokemon(MapPokeNum).Y - TargetY
                            If onSightDistanceX <= onSightRange And onSightDistanceY <= onSightRange Then

                            Else
                                '//Fish system
                                If Spawn(MapPokeNum).Fishing = YES Then
                                    ClearMapPokemon MapPokeNum
                                End If
                                ' Lost Target
                                MapPokemon(MapPokeNum).targetType = 0
                                MapPokemon(MapPokeNum).TargetIndex = 0
                                targetType = 0
                                Target = 0
                                TargetX = 0
                                TargetY = 0
                            End If
                        Else
                            '//Switch Target
                            MapPokemon(MapPokeNum).targetType = TARGET_TYPE_PLAYER
                        End If
                    End Select
                ElseIf MapPokemon(MapPokeNum).TargetIndex = 0 Then
                    '//Checking Target
                    If Pokemon(MapPokemon(MapPokeNum).Num).Behaviour = 1 Or Pokemon(MapPokemon(MapPokeNum).Num).Behaviour = 3 Then    '//2 = Attack On Sight / 4 = Flee On Sight
                        For i = 1 To Player_HighIndex
                            If IsPlaying(i) Then
                                If TempPlayer(i).UseChar > 0 Then
                                    If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                                        If TempPlayer(i).MapSwitchTmr = NO Then
                                            If Player(i, TempPlayer(i).UseChar).Access <= ACCESS_MAPPER Then
                                                onSightRange = Pokemon(MapPokemon(MapPokeNum).Num).Range
                                                'If PlayerPokemon(i).Num > 0 Then
                                                'onSightDistanceX = MapPokemon(MapPokeNum).x - PlayerPokemon(i).x
                                                'onSightDistanceY = MapPokemon(MapPokeNum).y - PlayerPokemon(i).y
                                                'Else
                                                onSightDistanceX = MapPokemon(MapPokeNum).x - Player(i, TempPlayer(i).UseChar).x
                                                onSightDistanceY = MapPokemon(MapPokeNum).Y - Player(i, TempPlayer(i).UseChar).Y
                                                'End If
                                                '//Make sure we get a positive value
                                                If onSightDistanceX < 0 Then onSightDistanceX = onSightDistanceX * -1
                                                If onSightDistanceY < 0 Then onSightDistanceY = onSightDistanceY * -1

                                                If onSightDistanceX <= onSightRange And onSightDistanceY <= onSightRange Then
                                                    MapPokemon(MapPokeNum).targetType = TARGET_TYPE_PLAYER
                                                    MapPokemon(MapPokeNum).TargetIndex = i
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If

                ' ***************
                ' ** Attacking **
                ' ***************
                If Target > 0 Then
                    If MapPokemon(MapPokeNum).AtkTmr <= GetTickCount Then
                        '//Check Direction
                        If MapPokemon(MapPokeNum).QueueMove <= 0 Then
                            'Select Case MapPokemon(MapPokeNum).Dir
                            '    Case DIR_LEFT: If MapPokemon(MapPokeNum).x - 1 = TargetX And MapPokemon(MapPokeNum).y = TargetY Then NearTarget = True
                            '    Case DIR_RIGHT: If MapPokemon(MapPokeNum).x + 1 = TargetX And MapPokemon(MapPokeNum).y = TargetY Then NearTarget = True
                            '    Case DIR_UP: If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).y - 1 = TargetY Then NearTarget = True
                            '    Case DIR_DOWN: If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).y + 1 = TargetY Then NearTarget = True
                            'End Select
                            NearTarget = False
                            If IsOnAoERange(4, MapPokemon(MapPokeNum).x, MapPokemon(MapPokeNum).Y, TargetX, TargetY) Then NearTarget = True

                            If NearTarget Then
                                '//Check Random use of move
                                randNum = Random(1, 4)
                                If randNum >= 1 And randNum <= 4 Then
                                    '//Check if moveset is available
                                    If MapPokemon(MapPokeNum).Moveset(randNum).Num > 0 Then
                                        If MapPokemon(MapPokeNum).Moveset(randNum).CD <= Tick Then
                                            NpcCastMove MapPokeNum, MapPokemon(MapPokeNum).Moveset(randNum).Num, randNum
                                            MapPokemon(MapPokeNum).AtkTmr = GetTickCount + 1000
                                            SendNpcAttack MapNum, MapPokeNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    QueueMove = MapPokemon(MapPokeNum).QueueMove
                    If QueueMove > 0 Then
                        '//Check Cast Time
                        If MapPokemon(MapPokeNum).MoveCastTime <= Tick Then
                            '//Check For Duration
                            If MapPokemon(MapPokeNum).MoveAttackCount < PokemonMove(QueueMove).AmountOfAttack Then
                                If MapPokemon(MapPokeNum).MoveDuration >= Tick Then
                                    If MapPokemon(MapPokeNum).MoveInterval <= Tick Then
                                        ProcessNpcMove MapPokeNum, QueueMove
                                        If MapPokemon(MapPokeNum).Num <= 0 Then
                                            Exiting = True
                                        End If
                                        MapPokemon(MapPokeNum).MoveInterval = GetTickCount + PokemonMove(QueueMove).Interval
                                        MapPokemon(MapPokeNum).MoveAttackCount = MapPokemon(MapPokeNum).MoveAttackCount + 1
                                    End If
                                Else
                                    If MapPokemon(MapPokeNum).QueueMoveSlot > 0 Then
                                        MapPokemon(MapPokeNum).Moveset(MapPokemon(MapPokeNum).QueueMoveSlot).CD = Tick
                                    End If
                                    '//Clear Queue Move
                                    MapPokemon(MapPokeNum).QueueMove = 0
                                    MapPokemon(MapPokeNum).QueueMoveSlot = 0
                                    MapPokemon(MapPokeNum).MoveDuration = 0
                                    MapPokemon(MapPokeNum).MoveCastTime = 0
                                    MapPokemon(MapPokeNum).MoveInterval = 0
                                End If
                            Else
                                If MapPokemon(MapPokeNum).QueueMoveSlot > 0 Then
                                    MapPokemon(MapPokeNum).Moveset(MapPokemon(MapPokeNum).QueueMoveSlot).CD = Tick
                                End If
                                '//Clear Queue Move
                                MapPokemon(MapPokeNum).QueueMove = 0
                                MapPokemon(MapPokeNum).QueueMoveSlot = 0
                                MapPokemon(MapPokeNum).MoveDuration = 0
                                MapPokemon(MapPokeNum).MoveCastTime = 0
                                MapPokemon(MapPokeNum).MoveInterval = 0
                            End If
                        End If
                    End If
                End If

                If Not Exiting Then
                    ' **************
                    ' ** Movement **
                    ' **************
                    If MapPokemon(MapPokeNum).MoveTmr <= GetTickCount Then

                        If IsWithinSpawnTime(MapPokeNum, GameHour) = False Then
                            '//Despawn
                            ClearMapPokemon MapPokeNum
                            Exit For
                        End If
                        '//Follow Target
                        If Target > 0 Then
                            randNum = Int(Rnd * 3)
                            DidWalk = False

                            '//CheckMovement
                            Select Case randNum
                            Case 0
                                If MapPokemon(MapPokeNum).Y > TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_UP) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y < TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_DOWN) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x > TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_LEFT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x < TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_RIGHT) Then DidWalk = True
                                End If
                            Case 1
                                If MapPokemon(MapPokeNum).x < TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_RIGHT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x > TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_LEFT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y < TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_DOWN) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y > TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_UP) Then DidWalk = True
                                End If
                            Case 2
                                If MapPokemon(MapPokeNum).Y < TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_DOWN) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y > TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_UP) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x < TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_RIGHT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x > TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_LEFT) Then DidWalk = True
                                End If
                            Case 3
                                If MapPokemon(MapPokeNum).x > TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_LEFT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x < TargetX And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_RIGHT) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y > TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_UP) Then DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).Y < TargetY And Not DidWalk Then
                                    If PokemonProcessMove(MapPokeNum, DIR_DOWN) Then DidWalk = True
                                End If
                            End Select

                            '//Check Direction
                            If Not DidWalk Then
                                If MapPokemon(MapPokeNum).x - 1 = TargetX And MapPokemon(MapPokeNum).Y = TargetY Then
                                    If Not MapPokemon(MapPokeNum).Dir = DIR_LEFT Then Call PokemonDir(MapPokeNum, DIR_LEFT)
                                    DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x + 1 = TargetX And MapPokemon(MapPokeNum).Y = TargetY Then
                                    If Not MapPokemon(MapPokeNum).Dir = DIR_RIGHT Then Call PokemonDir(MapPokeNum, DIR_RIGHT)
                                    DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).Y - 1 = TargetY Then
                                    If Not MapPokemon(MapPokeNum).Dir = DIR_UP Then Call PokemonDir(MapPokeNum, DIR_UP)
                                    DidWalk = True
                                End If
                                If MapPokemon(MapPokeNum).x = TargetX And MapPokemon(MapPokeNum).Y + 1 = TargetY Then
                                    If Not MapPokemon(MapPokeNum).Dir = DIR_DOWN Then Call PokemonDir(MapPokeNum, DIR_DOWN)
                                    DidWalk = True
                                End If

                                If Not DidWalk Then
                                    '//Randomize number to prevent continues movement
                                    randNum = Int(Rnd * 15)

                                    If randNum = 1 Then
                                        '//Randomize number for direction
                                        randNum = Int(Rnd * 5)

                                        '//Process Move
                                        If PokemonProcessMove(MapPokeNum, randNum) Then
                                            '//Do Nothing
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            '//Randomize number to prevent continues movement
                            randNum = Int(Rnd * 15)

                            If randNum = 1 Then
                                '//Randomize number for direction
                                randNum = Int(Rnd * 5)

                                '//Process Move
                                If PokemonProcessMove(MapPokeNum, randNum) Then
                                    '//Do Nothing
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    DoEvents
End Sub

Private Sub HandleShutdown()
    If Secs Mod 10 = 0 Or Secs <= 5 Then
        Call SendGlobalMsg("Server Shutdown in " & Secs & " seconds.", White)
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        DestroyServer
        Exit Sub
    End If
End Sub

Attribute VB_Name = "modPokemon"
Option Explicit

Public Sub SpawnMapPokemon(ByVal MapPokeNum As Long, Optional ByVal ForceSpawn As Boolean = False, Optional ByVal ForceShiny As Byte = NO, Optional ByVal FishIndex As Long = 0)
    Dim MapNum As Long, x As Long, Y As Long
    Dim RndNum As Long
    Dim x2 As Long, y2 As Long
    Dim gotData As Boolean

    '//Check Error
    If MapPokemon(MapPokeNum).Num > 0 Then Exit Sub
    If MapPokeNum <= 0 Or MapPokeNum > MAX_GAME_POKEMON Then Exit Sub

    '//Check all
    If Spawn(MapPokeNum).PokeNum > 0 Then
        MapNum = Random(1, MAX_MAP)
        '//Check Position
        If Spawn(MapPokeNum).randomMap = NO Then
            If Spawn(MapPokeNum).MapNum > 0 Then
                MapNum = Spawn(MapPokeNum).MapNum
            End If
        End If

        If MapNum <= 1 Then MapNum = 1
        If MapNum >= MAX_MAP Then MapNum = MAX_MAP

        If FishIndex = 0 Then
            If Not Map(MapNum).Moral = MAP_MORAL_SAFARI Then    '//Can Spawn at Saffari
                If Map(MapNum).Moral = MAP_MORAL_ARENA Or Map(MapNum).Moral = MAP_MORAL_SAFE Then    '//Don't spawn
                    Exit Sub
                End If
            End If
        End If
        If Len(Trim$(Map(MapNum).Name)) <= 0 Then
            Exit Sub
        End If

        '//check rarity
        If Not ForceSpawn Then
            RndNum = Random(0, Spawn(MapPokeNum).Rarity)
            If Not RndNum = 0 Then
                MapPokemon(MapPokeNum).Respawn = GetTickCount + Spawn(MapPokeNum).Respawn
                Exit Sub
            End If
        End If

        '//check HeldItem equipped
        If Spawn(MapPokeNum).HeldItem > 0 Then
            MapPokemon(MapPokeNum).HeldItem = Spawn(MapPokeNum).HeldItem
        End If

        '//Check Position
        gotData = False
        If FishIndex = 0 Then    ' Adaptado pra usar com a pesca, que cai em cima do jogador o peixe
            If Spawn(MapPokeNum).randomXY = NO Then
                x = Spawn(MapPokeNum).MapX
                Y = Spawn(MapPokeNum).MapY
            Else
                '//randomize value for 100 times
                If Not gotData Then
                    For RndNum = 1 To 100
                        x = Random(0, Map(MapNum).MaxX)
                        Y = Random(0, Map(MapNum).MaxY)

                        If NpcTileOpen(MapNum, x, Y) Then
                            gotData = True
                            Exit For
                        End If
                    Next
                End If

                '//spawn on the free tile
                If Not gotData Then
                    For x2 = 0 To Map(MapNum).MaxX
                        For y2 = 0 To Map(MapNum).MaxY
                            If NpcTileOpen(MapNum, x2, y2) Then
                                x = x2
                                Y = y2
                                gotData = True
                                Exit For
                            End If
                        Next
                    Next
                End If

                If Not gotData Then
                    x = Spawn(MapPokeNum).MapX
                    Y = Spawn(MapPokeNum).MapY
                End If
            End If
        Else    '//Sistema de fish
            If IsPlaying(FishIndex) Then
                If TempPlayer(FishIndex).UseChar > 0 Then
                    x = GetPlayerX(FishIndex)
                    Y = GetPlayerY(FishIndex)
                    MapNum = GetPlayerMap(FishIndex)
                End If
            End If
        End If
        If x <= 0 Then x = 0
        If x >= Map(MapNum).MaxX Then x = Map(MapNum).MaxX
        If Y <= 0 Then Y = 0
        If Y >= Map(MapNum).MaxY Then Y = Map(MapNum).MaxY


        'Debug.Print Pokemon(Spawn(MapPokeNum).PokeNum).Name
        If IsWithinSpawnTime(MapPokeNum, GameHour) = True Then
            ' Spawn Pokemon
            SpawnPokemon MapPokeNum, Spawn(MapPokeNum).PokeNum, MapNum, x, Y, Random(0, 3), ForceSpawn, ForceShiny
        End If
    End If
End Sub

Function IsWithinSpawnTime(MapPokeNum As Long, hour As Byte) As Boolean
    Dim isNightSpawn As Boolean

    If Spawn(MapPokeNum).SpawnTimeMin > Spawn(MapPokeNum).SpawnTimeMax Then
        ' Spawn noturno
        If hour >= Spawn(MapPokeNum).SpawnTimeMin Or hour <= Spawn(MapPokeNum).SpawnTimeMax Then
            IsWithinSpawnTime = True
        End If
    Else
        ' Spawn regular
        If hour >= Spawn(MapPokeNum).SpawnTimeMin And hour <= Spawn(MapPokeNum).SpawnTimeMax Then
            IsWithinSpawnTime = True
        End If
    End If
End Function

Public Sub SpawnAllMapPokemon()
Dim i As Long

    '//Check all
    For i = 1 To MAX_GAME_POKEMON
        ClearMapPokemon i, True
        '//Add Data
        MapPokemon(i).PokemonIndex = Spawn(i).PokeNum
        MapPokemon(i).Respawn = GetTickCount
        '//Spawn
        If Spawn(i).Fishing = NO Then
            SpawnMapPokemon i
        End If
    Next
End Sub

Public Sub ClearMapPokemon(ByVal MapPokeNum As Long, Optional ByVal RemoveData As Boolean = False)
Dim i As Long
Dim MapNum As Long

    '//Add Cache
    MapNum = MapPokemon(MapPokeNum).Map
    If Not RemoveData Then
        If MapPokemon(MapPokeNum).PokemonIndex > 0 Then
            i = MapPokemon(MapPokeNum).PokemonIndex
        End If
    End If
    
    '//Clear
    Call ZeroMemory(ByVal VarPtr(MapPokemon(MapPokeNum)), LenB(MapPokemon(MapPokeNum)))
    SendPokemonData MapPokeNum, , MapNum
    
    '//Load Cache
    If Not RemoveData Then
        If i > 0 Then
            MapPokemon(MapPokeNum).PokemonIndex = i
            MapPokemon(MapPokeNum).Respawn = GetTickCount + Spawn(MapPokeNum).Respawn
        End If
    End If
End Sub

Public Function FindPokemonSlot(ByVal MapNum As Long) As Long
Dim i As Long

    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num <= 0 And MapPokemon(i).PokemonIndex <= 0 Then
            FindPokemonSlot = i
            Exit Function
        End If
    Next
End Function

Public Function SpawnPokemon(ByVal slot As Long, ByVal PokemonNum As Long, ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long, ByVal Dir As Byte, Optional ByVal ForceSpawn As Boolean = False, Optional ByVal ForceShiny As Byte = NO) As Boolean
    Dim bs As Byte, m As Long, s As Byte
    Dim ShinyChanceVal As Long, ShinyLuckVal As Long
    Dim MoveSlot As Long

    SpawnPokemon = False

    '//Check for error
    If PokemonNum <= 0 Or PokemonNum > MAX_POKEMON Then Exit Function
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Function
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Function
    If slot <= 0 Or slot > MAX_GAME_POKEMON Then Exit Function
    If x <= 0 Then x = 0
    If x >= Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If Y <= 0 Then Y = 0
    If Y >= Map(MapNum).MaxY Then Y = Map(MapNum).MaxY

    '//Update HighIndex
    If slot > Pokemon_HighIndex Then
        Pokemon_HighIndex = slot
        '//Send to all
        SendPokemonHighIndex
    End If

    With MapPokemon(slot)
        '//General
        .Num = PokemonNum

        '//Location
        .Map = MapNum
        .x = x
        .Y = Y
        .Dir = Dir

        '//Nature none implementada a partir do spawn editor
        If Spawn(slot).Nature = -1 Then
            .Nature = Random(0, PokemonNature.PokemonNature_Count - 1)
            If .Nature <= 0 Then .Nature = 0
            If .Nature >= (PokemonNature.PokemonNature_Count - 1) Then .Nature = PokemonNature.PokemonNature_Count - 1
        ElseIf Spawn(slot).Nature >= 0 Then
            .Nature = Spawn(slot).Nature
        End If

        '//Shiny Randomizer
        ShinyChanceVal = Random(0, Options.ShinyRarity)
        ShinyLuckVal = Random(0, Options.ShinyRarity)

        If ShinyChanceVal = ShinyLuckVal Then
            .IsShiny = YES
        Else
            .IsShiny = NO
        End If
        If ForceSpawn Then
            .IsShiny = ForceShiny
        End If

        .Gender = Random(GENDER_MALE, GENDER_FEMALE)
        If Not .Gender = GENDER_MALE And Not .Gender = GENDER_FEMALE Then
            .Gender = GENDER_MALE
        End If

        '//Status
        .Status = 0

        '//Level
        .Level = Random(Spawn(slot).MinLevel, Spawn(slot).MaxLevel)
        If .Level <= 1 Then .Level = 1
        If .Level >= MAX_LEVEL Then .Level = MAX_LEVEL

        '//Stats
        For bs = 1 To StatEnum.Stat_Count - 1
            .Stat(bs).EV = 0
            .Stat(bs).IV = Random(1, 31)
            If .Stat(bs).IV > 31 Then .Stat(bs).IV = 31
            If .Stat(bs).IV < 1 Then .Stat(bs).IV = 1
            .Stat(bs).Value = CalculatePokemonStat(bs, .Num, .Level, .Stat(bs).EV, .Stat(bs).IV, .Nature)
        Next

        '//Vital
        .MaxHp = .Stat(StatEnum.HP).Value * Spawn(slot).pokeBuff
        .CurHp = .MaxHp

        '//Moveset
        If PokemonNum > 0 Then
            For m = MAX_POKEMON_MOVESET To 1 Step -1
                '//Got Move
                If Pokemon(PokemonNum).Moveset(m).MoveNum > 0 Then
                    '//Check level
                    If .Level >= Pokemon(PokemonNum).Moveset(m).MoveLevel Then
                        For s = 1 To MAX_MOVESET
                            If .Moveset(s).Num <= 0 Then
                                MoveSlot = s
                                Exit For
                            End If
                        Next

                        '//Add Move
                        If MoveSlot > 0 Then
                            .Moveset(MoveSlot).Num = Pokemon(PokemonNum).Moveset(m).MoveNum
                            .Moveset(MoveSlot).TotalPP = PokemonMove(Pokemon(PokemonNum).Moveset(m).MoveNum).PP
                            .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).TotalPP
                        End If
                    End If
                End If
            Next
        End If

        '//Clear Target
        .TargetIndex = 0
        .targetType = 0

        '//Update Data to map
        SendPokemonData slot, , MapNum
    End With
End Function

Public Function PokemonProcessMove(ByVal MapPokemonNum As Long, ByVal Dir As Byte, Optional ByVal CheckDir As Boolean = False) As Boolean
Dim DidMove As Boolean
Dim expEarn As Long
Dim RndNum As Byte
Dim RandomNum As Long
Dim MoveSpeed As Long
Dim i As Byte, pCount As Byte

    '//Exit out when error
    If MapPokemonNum <= 0 Or MapPokemonNum > MAX_GAME_POKEMON Then Exit Function
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Function

    DidMove = False

    With MapPokemon(MapPokemonNum)
        If .Num <= 0 Then Exit Function
        'If .QueueMove > 0 Then Exit Function
        If .Status = StatusEnum.Sleep Then
            If .Status = StatusEnum.Paralize Then
                .MoveTmr = GetTickCount + 1000
            Else
                MoveSpeed = CalculateSpeed(GetNpcPokemonStat(MapPokemonNum, Spd))
                If MoveSpeed <= 250 Then MoveSpeed = 250
                .MoveTmr = GetTickCount + MoveSpeed
            End If
            Exit Function
        End If
        If .Status = StatusEnum.Frozen Then
            If .Status = StatusEnum.Paralize Then
                .MoveTmr = GetTickCount + 1000
            Else
                MoveSpeed = CalculateSpeed(GetNpcPokemonStat(MapPokemonNum, Spd))
                If MoveSpeed <= 250 Then MoveSpeed = 250
                .MoveTmr = GetTickCount + MoveSpeed
            End If
            Exit Function
        End If

        If .IsConfuse = YES Then
            Dir = Random(0, 3)
            If Dir < 0 Then Dir = 0
            If Dir > DIR_RIGHT Then Dir = DIR_RIGHT
            RndNum = Random(1, 10)
            If RndNum = 1 Then
                .IsConfuse = 0
            End If
        End If

        Select Case Dir
            Case DIR_UP
                .Dir = DIR_UP
                
                '//Check to make sure not outside of boundries
                If .Y > 0 Then
                    If Not CheckDirection(.Map, DIR_UP, .x, .Y, True) Then
                        .Y = .Y - 1
                        DidMove = True
                    End If
                End If
            Case DIR_DOWN
                .Dir = DIR_DOWN
                
                '//Check to make sure not outside of boundries
                If .Y < Map(.Map).MaxY Then
                    If Not CheckDirection(.Map, DIR_DOWN, .x, .Y, True) Then
                        .Y = .Y + 1
                        DidMove = True
                    End If
                End If
            Case DIR_LEFT
                .Dir = DIR_LEFT
                
                '//Check to make sure not outside of boundries
                If .x > 0 Then
                    If Not CheckDirection(.Map, DIR_LEFT, .x, .Y, True) Then
                        .x = .x - 1
                        DidMove = True
                    End If
                End If
            Case DIR_RIGHT
                .Dir = DIR_RIGHT
                
                '//Check to make sure not outside of boundries
                If .x < Map(.Map).MaxX Then
                    If Not CheckDirection(.Map, DIR_RIGHT, .x, .Y, True) Then
                        .x = .x + 1
                        DidMove = True
                    End If
                End If
        End Select
        
        If DidMove Then
            '//Poison
            If .Status = StatusEnum.Poison Then
                If .StatusMove >= 4 Then
                    If .StatusDamage > 0 Then
                        If .StatusDamage >= .CurHp Then
                            .CurHp = 0
                            SendActionMsg .Map, "-" & .StatusDamage, .x * 32, .Y * 32, Magenta
                            
                            If Spawn(MapPokemonNum).NoExp = NO Then
                                If .LastAttacker > 0 Then
                                    If IsPlaying(.LastAttacker) Then
                                        If TempPlayer(.LastAttacker).UseChar > 0 Then
                                            If PlayerPokemon(.LastAttacker).Num > 0 Then
                                                If PlayerPokemon(.LastAttacker).slot > 0 Then
                                                    expEarn = ((Pokemon(.Num).BaseExp * .Level) * 1 * 1) / 7
                                                    If TempPlayer(.LastAttacker).InParty > 0 Then
                                                        '//Share to party
                                                        pCount = PartyCount(.LastAttacker)
                                                        For i = 1 To MAX_PARTY
                                                            If TempPlayer(.LastAttacker).PartyIndex(i) > 0 Then
                                                                If IsPlaying(TempPlayer(.LastAttacker).PartyIndex(i)) Then
                                                                    If TempPlayer(TempPlayer(.LastAttacker).PartyIndex(i)).UseChar > 0 Then
                                                                        If Player(TempPlayer(.LastAttacker).PartyIndex(i), TempPlayer(TempPlayer(.LastAttacker).PartyIndex(i)).UseChar).Map = .Map Then
                                                                            GivePlayerPokemonExp TempPlayer(.LastAttacker).PartyIndex(i), PlayerPokemon(TempPlayer(.LastAttacker).PartyIndex(i)).slot, (expEarn / pCount)
                                                                            GivePlayerExp TempPlayer(.LastAttacker).PartyIndex(i), ((expEarn / 4) / pCount)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next
                                                    Else
                                                        GivePlayerPokemonExp .LastAttacker, PlayerPokemon(.LastAttacker).slot, expEarn
                                                        GivePlayerExp .LastAttacker, (expEarn / 4)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            
                            ClearMapPokemon MapPokemonNum
                            Exit Function
                        Else
                            .CurHp = .CurHp - .StatusDamage
                            SendActionMsg .Map, "-" & .StatusDamage, .x * 32, .Y * 32, Magenta
                            '//Update
                            SendPokemonVital MapPokemonNum
                        End If
                        '//Reset
                        .StatusMove = 0
                    Else
                        .StatusDamage = (.MaxHp / 16)
                    End If
                Else
                    .StatusMove = .StatusMove + 1
                End If
            End If
        
            If .Status = StatusEnum.Paralize Then
                .MoveTmr = GetTickCount + 1000
            Else
                MoveSpeed = CalculateSpeed(GetNpcPokemonStat(MapPokemonNum, Spd))
                If MoveSpeed <= 250 Then MoveSpeed = 250
                .MoveTmr = GetTickCount + MoveSpeed
            End If
            SendPokemonMove MapPokemonNum
            PokemonProcessMove = True
        Else
            '//Update Dir
            SendPokemonDir MapPokemonNum
            PokemonProcessMove = False
        End If
    End With
End Function

Public Sub PokemonDir(ByVal MapPokemonNum As Long, ByVal Dir As Byte)
    '//Exit out when error
    If MapPokemonNum <= 0 Or MapPokemonNum > MAX_GAME_POKEMON Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    
    If MapPokemon(MapPokemonNum).Num > 0 Then
        MapPokemon(MapPokemonNum).Dir = Dir
        '//Update Dir
        SendPokemonDir MapPokemonNum
    End If
End Sub

'//Exp
Public Function GetPokemonNextExp(ByVal Level As Long, ByVal ExpType As Byte) As Long
    Select Case ExpType
        Case GrowthRateEnum.Erratic
            If Level < 50 Then
                GetPokemonNextExp = ((Level ^ 3) * (100 - Level)) / 50
            ElseIf Level >= 50 And Level < 68 Then
                GetPokemonNextExp = ((Level ^ 3) * (150 - Level)) / 100
            ElseIf Level >= 68 And Level < 98 Then
                GetPokemonNextExp = ((Level ^ 3) * ((1191 - (10 * Level)) / 3)) / 50
            Else
                GetPokemonNextExp = ((Level ^ 3) * (160 - Level)) / 100
            End If
        Case GrowthRateEnum.Fast
            GetPokemonNextExp = 0.8 * (Level ^ 3)
        Case GrowthRateEnum.MediumFast
            GetPokemonNextExp = (Level ^ 3)
        Case GrowthRateEnum.MediumSlow
            GetPokemonNextExp = 1.2 * (Level ^ 3) - 15 * (Level ^ 2) + 100 * (Level) - 140
        Case GrowthRateEnum.Slow
            GetPokemonNextExp = 1.25 * (Level ^ 3)
        Case GrowthRateEnum.Fluctuating
            If Level < 15 Then
                GetPokemonNextExp = (Level ^ 3) * ((((Level + 1) / 3) + 24) / 50)
            ElseIf Level >= 15 And Level < 36 Then
                GetPokemonNextExp = (Level ^ 3) * ((Level + 4) / 50)
            Else
                GetPokemonNextExp = (Level ^ 3) * (((Level / 2) + 32) / 50)
            End If
        Case Else
            If Level < 15 Then
                GetPokemonNextExp = (Level ^ 3) * ((((Level + 1) / 3) + 24) / 50)
            ElseIf Level >= 15 And Level < 36 Then
                GetPokemonNextExp = (Level ^ 3) * ((Level + 4) / 50)
            Else
                GetPokemonNextExp = (Level ^ 3) * (((Level / 2) + 32) / 50)
            End If
    End Select
End Function

Public Sub DefeatMapPokemon(ByVal MapPokeNum As Long)
Dim expEarn As Long, i As Byte, pCount As Byte

    With MapPokemon(MapPokeNum)
        If Spawn(MapPokeNum).NoExp = NO Then
            If .LastAttacker > 0 Then
                If IsPlaying(.LastAttacker) Then
                    If TempPlayer(.LastAttacker).UseChar > 0 Then
                        If PlayerPokemon(.LastAttacker).Num > 0 Then
                            If PlayerPokemon(.LastAttacker).slot > 0 Then
                                expEarn = ((Pokemon(.Num).BaseExp * .Level) * 1 * 1) / 7
                                If TempPlayer(.LastAttacker).InParty > 0 Then
                                    '//Share to party
                                    pCount = PartyCount(.LastAttacker)
                                    For i = 1 To MAX_PARTY
                                        If TempPlayer(.LastAttacker).PartyIndex(i) > 0 Then
                                            If IsPlaying(TempPlayer(.LastAttacker).PartyIndex(i)) Then
                                                If TempPlayer(TempPlayer(.LastAttacker).PartyIndex(i)).UseChar > 0 Then
                                                    If Player(TempPlayer(.LastAttacker).PartyIndex(i), TempPlayer(TempPlayer(.LastAttacker).PartyIndex(i)).UseChar).Map = .Map Then
                                                        GivePlayerPokemonExp TempPlayer(.LastAttacker).PartyIndex(i), PlayerPokemon(TempPlayer(.LastAttacker).PartyIndex(i)).slot, (expEarn / pCount)
                                                        GivePlayerExp TempPlayer(.LastAttacker).PartyIndex(i), ((expEarn / 4) / pCount)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                Else
                                    GivePlayerPokemonExp .LastAttacker, PlayerPokemon(.LastAttacker).slot, expEarn
                                    GivePlayerExp .LastAttacker, (expEarn / 4)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub


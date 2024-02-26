Attribute VB_Name = "modLogic"
Option Explicit

'//This function change a single digit number to two digit (often used on time)
Public Function KeepTwoDigit(ByVal Val As Long) As String
    If Val > 9 Then
        KeepTwoDigit = Val
    Else
        KeepTwoDigit = "0" & Val
    End If
End Function

Public Sub TextAdd(TargetObj As TextBox, ByVal sString As String)
Dim SelText As String

    CountText = CountText + 1
    If CountText > 500 Then
        TargetObj.Text = ""
        CountText = 0
    End If

    sString = KeepTwoDigit(hour(Now)) & ":" & KeepTwoDigit(Minute(Now)) & " : " & sString

    SelText = TargetObj.Text
    If Len(SelText) > 0 Then
        TargetObj.Text = SelText & vbNewLine & sString
    Else
        TargetObj.Text = sString
    End If
    TargetObj.SelStart = Len(TargetObj.Text)
End Sub

Public Function isNameLegal(ByVal KeyAscii As Integer, Optional ByVal DisableSpaceBar As Boolean = False) As Boolean
    If DisableSpaceBar Then
        If KeyAscii = 32 Then
            isNameLegal = False
            Exit Function
        End If
    End If
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii = 32) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        isNameLegal = True
    End If
End Function

Public Function isStringLegal(ByVal KeyAscii As Integer, Optional ByVal DisableSpaceBar As Boolean = False) As Boolean
    If DisableSpaceBar Then
        If KeyAscii = 32 Then
            isStringLegal = False
            Exit Function
        End If
    End If
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii = 32) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 33 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or (KeyAscii >= 123 And KeyAscii <= 126) Then
        isStringLegal = True
    End If
End Function

Public Function CheckNameInput(ByVal Name As String, Optional ByVal HaveStringLimit As Boolean = False, Optional ByVal MaxLimit As Long = 0, Optional ByVal isString As Boolean = False) As Boolean
Dim i As Long, n As Long

    If Not HaveStringLimit Then
        ' Check if name is within the letter limit
        If Len(Name) <= 2 Or Len(Name) >= MaxLimit Then
            CheckNameInput = False
            Exit Function
        End If
    End If
    
    ' Check Legal Asc key
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))
        
        If isString Then
            If Not isStringLegal(n, True) Then
                CheckNameInput = False
                Exit Function
            End If
        Else
            If Not isNameLegal(n, True) Then
                CheckNameInput = False
                Exit Function
            End If
        End If
    Next
    
    CheckNameInput = True
End Function

Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function isValidMapPoint(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
    isValidMapPoint = False
    If x < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If x > Map(MapNum).MaxX Then Exit Function
    If Y > Map(MapNum).MaxY Then Exit Function
    isValidMapPoint = True
End Function

Public Function CheckDirection(ByVal MapNum As Long, ByVal Direction As Byte, ByVal x As Long, ByVal Y As Long, Optional ByVal NpcChecking As Boolean = False) As Boolean
Dim wX As Long, wY As Long
Dim i As Long
Dim xIndex As Long

    CheckDirection = False
 
    Select Case Direction
        Case DIR_UP
            wX = x
            wY = Y - 1
        Case DIR_DOWN
            wX = x
            wY = Y + 1
        Case DIR_LEFT
            wX = x - 1
            wY = Y
        Case DIR_RIGHT
            wX = x + 1
            wY = Y
    End Select

    If wX < 0 Or wX > Map(MapNum).MaxX Or wY < 0 Or wY > Map(MapNum).MaxY Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.Blocked Then
        CheckDirection = True
        Exit Function
    End If
    If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.ConvoTile Then
        CheckDirection = True
        Exit Function
    End If
    If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.BothStorage Or Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.InvStorage Or Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.PokemonStorage Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.FishSpot Then
        CheckDirection = True
        Exit Function
    End If
    
    If NpcChecking Then
        For i = 1 To MAX_MAP_NPC
            '//Check Npc
            If MapNpc(MapNum, i).Num > 0 Then
                If MapNpc(MapNum, i).x = wX And MapNpc(MapNum, i).Y = wY Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
            If MapNpcPokemon(MapNum, i).Num > 0 Then
                If MapNpcPokemon(MapNum, i).x = wX And MapNpcPokemon(MapNum, i).Y = wY Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        Next
    
        '//Check Player
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                        If Player(i, TempPlayer(i).UseChar).x = wX And Player(i, TempPlayer(i).UseChar).Y = wY Then
                            CheckDirection = True
                            Exit Function
                        End If
                        '//Player Pokemon
                        If PlayerPokemon(i).Num > 0 Then
                            If PlayerPokemon(i).x = wX And PlayerPokemon(i).Y = wY Then
                                CheckDirection = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Next
        
        '//Check Npc
        For i = 1 To Pokemon_HighIndex
            If MapPokemon(i).Num > 0 Then
                If MapPokemon(i).Map = MapNum Then
                    If MapPokemon(i).x = wX And MapPokemon(i).Y = wY Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        Next
        
        If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.NpcAvoid Then
            CheckDirection = True
            Exit Function
        End If
        
        If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.Warp Then
            CheckDirection = True
            Exit Function
        End If
        
        If Map(MapNum).Tile(wX, wY).Attribute = MapAttribute.WarpCheckpoint Then
            CheckDirection = True
            Exit Function
        End If
    End If
End Function

' ********************
' ** Npc Properties **
' ********************
Public Sub ClearMapNpc(ByVal MapNum As Long, ByVal MapNpcNum As Byte)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum, MapNpcNum)), LenB(MapNpc(MapNum, MapNpcNum)))
End Sub

Public Sub ClearMapNpcs()
Dim x As Long, Y As Long
    
    For x = 1 To MAX_MAP
        For Y = 1 To MAX_MAP_NPC
            Call ClearMapNpc(x, Y)
        Next
    Next
End Sub

Public Function NpcTileOpen(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
    NpcTileOpen = True
    
    '//Check if npc can step on the tile
    If Not Map(MapNum).Tile(x, Y).Attribute = MapAttribute.Walkable Then
        NpcTileOpen = False
        Exit Function
    End If
End Function

Public Sub CheckSpawnNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Dim i As Long, SpawnAllDay As Boolean, WeekDayNum As Byte

    If Map(MapNum).Npc(MapNpcNum) = 0 Then Exit Sub

    SpawnAllDay = True

    For i = 1 To WeekDayEnum.Count_WeekDay - 1
        If Npc(Map(MapNum).Npc(MapNpcNum)).SpawnWeekDay(i) = YES Then
            SpawnAllDay = False
            Exit For
        End If
    Next i

    If Not SpawnAllDay Then
        WeekDayNum = Weekday(Date)
        For i = 1 To WeekDayEnum.Count_WeekDay - 1
            If Npc(Map(MapNum).Npc(MapNpcNum)).SpawnWeekDay(i) = YES Then
                If i = WeekDayNum Then
                    If MapNpc(MapNum, MapNpcNum).Num = 0 Then
                        Call SpawnNpc(MapNum, MapNpcNum)
                    End If
                    Exit Sub
                Else
                    If MapNpc(MapNum, MapNpcNum).Num <> 0 Then
                        Call DespawnNpc(MapNum, MapNpcNum)
                    End If
                    Exit Sub
                End If
            End If
        Next i
    Else
        If MapNpc(MapNum, MapNpcNum).Num = 0 Then
            Call SpawnNpc(MapNum, MapNpcNum)
        End If
    End If
End Sub

Public Sub DespawnNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Dim x As Long, Y As Long
    Dim i As Long
    Dim DidSpawn As Boolean

    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub

    '//Check if Npc Exist
    If Map(MapNum).Npc(MapNpcNum) > 0 Then
        With MapNpc(MapNum, MapNpcNum)
            '//Input data
            .Num = 0
            .x = 0
            .Y = 0

            '//Send data to map
            SendSpawnMapNpc MapNum, MapNpcNum
        End With
    End If
End Sub

Public Sub SpawnNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim x As Long, Y As Long
Dim i As Long
Dim DidSpawn As Boolean

    '//Check for error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    
    '//Check if Npc Exist
    If Map(MapNum).Npc(MapNpcNum) > 0 Then
        With MapNpc(MapNum, MapNpcNum)
            '//Check position
            DidSpawn = False
            
            '//check on tiles if it have a specific location
            If Not DidSpawn Then
                For x = 0 To Map(MapNum).MaxX
                    For Y = 0 To Map(MapNum).MaxY
                        If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.NpcSpawn Then
                            If Map(MapNum).Tile(x, Y).Data1 = MapNpcNum Then
                                .x = x
                                .Y = Y
                                .Dir = Map(MapNum).Tile(x, Y).Data2
                                DidSpawn = True
                                GoTo Continue
                            End If
                        End If
                    Next Y
                Next x
Continue:
            End If
            
            
            '//randomize value for 100 times
            If Not DidSpawn Then
                For i = 1 To 100
                    x = Random(0, Map(MapNum).MaxX)
                    Y = Random(0, Map(MapNum).MaxY)
                    
                    If NpcTileOpen(MapNum, x, Y) Then
                        .x = x
                        .Y = Y
                        .Dir = Random(0, 3)
                        DidSpawn = True
                        Exit For
                    End If
                Next i
            End If
            
            '//spawn on the free tile
            If Not DidSpawn Then
                For x = 0 To Map(MapNum).MaxX
                    For Y = 0 To Map(MapNum).MaxY
                        If NpcTileOpen(MapNum, x, Y) Then
                            .x = x
                            .Y = Y
                            .Dir = Random(0, 3)
                            DidSpawn = True
                        End If
                    Next Y
                Next x
            End If
            
            If DidSpawn Then
                '//Input data
                .Num = Map(MapNum).Npc(MapNpcNum)
                
                '//Send data to map
                SendSpawnMapNpc MapNum, MapNpcNum
            End If
        End With
    End If
End Sub

Public Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPC
        Call SpawnNpc(MapNum, i)
    Next
End Sub

Public Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP
        Call SpawnMapNpcs(i)
    Next
End Sub

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte)
Dim DidMove As Boolean

    '//Exit out when error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    If MapNpc(MapNum, MapNpcNum).Num <= 0 Then Exit Sub
    If Map(MapNum).Npc(MapNpcNum) <= 0 Then Exit Sub

    DidMove = False

    With MapNpc(MapNum, MapNpcNum)
        Select Case Dir
            Case DIR_UP
                .Dir = DIR_UP
                
                '//Check to make sure not outside of boundries
                If .Y > 0 Then
                    If Not CheckDirection(MapNum, DIR_UP, .x, .Y, True) Then
                        .Y = .Y - 1
                        DidMove = True
                    End If
                End If
            Case DIR_DOWN
                .Dir = DIR_DOWN
                
                '//Check to make sure not outside of boundries
                If .Y < Map(MapNum).MaxY Then
                    If Not CheckDirection(MapNum, DIR_DOWN, .x, .Y, True) Then
                        .Y = .Y + 1
                        DidMove = True
                    End If
                End If
            Case DIR_LEFT
                .Dir = DIR_LEFT
                
                '//Check to make sure not outside of boundries
                If .x > 0 Then
                    If Not CheckDirection(MapNum, DIR_LEFT, .x, .Y, True) Then
                        .x = .x - 1
                        DidMove = True
                    End If
                End If
            Case DIR_RIGHT
                .Dir = DIR_RIGHT
                
                '//Check to make sure not outside of boundries
                If .x < Map(MapNum).MaxX Then
                    If Not CheckDirection(MapNum, DIR_RIGHT, .x, .Y, True) Then
                        .x = .x + 1
                        DidMove = True
                    End If
                End If
        End Select
        
        If DidMove Then
            SendNpcMove MapNum, MapNpcNum
        Else
            '//Update Dir
            SendNpcDir MapNum, MapNpcNum
        End If
    End With
End Sub

Public Function CheckOpenTile(ByVal MapNum As Long, ByVal x As Long, ByVal Y As Long) As Boolean
Dim i As Long

    CheckOpenTile = True
    
    If x < 0 Or Y < 0 Or x > Map(MapNum).MaxX Or Y > Map(MapNum).MaxY Then
        CheckOpenTile = False
        Exit Function
    End If
    '//Check if npc can step on the tile
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.Blocked Then
        CheckOpenTile = False
        Exit Function
    End If
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.ConvoTile Then
        CheckOpenTile = False
        Exit Function
    End If
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.BothStorage Or Map(MapNum).Tile(x, Y).Attribute = MapAttribute.InvStorage Or Map(MapNum).Tile(x, Y).Attribute = MapAttribute.PokemonStorage Then
        CheckOpenTile = False
        Exit Function
    End If
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.Warp Then
        CheckOpenTile = False
        Exit Function
    End If
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.WarpCheckpoint Then
        CheckOpenTile = False
        Exit Function
    End If
    
    If Map(MapNum).Tile(x, Y).Attribute = MapAttribute.FishSpot Then
        CheckOpenTile = False
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_NPC
        '//Check Npc
        If MapNpc(MapNum, i).Num > 0 Then
            If MapNpc(MapNum, i).x = x And MapNpc(MapNum, i).Y = Y Then
                CheckOpenTile = False
                Exit Function
            End If
        End If
    Next
    
    '//Check Player
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                    If Player(i, TempPlayer(i).UseChar).x = x And Player(i, TempPlayer(i).UseChar).Y = Y Then
                        CheckOpenTile = False
                        Exit Function
                    End If
                    '//Player Pokemon
                    If PlayerPokemon(i).Num > 0 Then
                        If PlayerPokemon(i).x = x And PlayerPokemon(i).Y = Y Then
                            CheckOpenTile = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
        
    '//Check Npc
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = MapNum Then
                If MapPokemon(i).x = x And MapPokemon(i).Y = Y Then
                    CheckOpenTile = False
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'///////////////////////////
'///// NPC Pokemon /////////
'///////////////////////////
Public Sub SpawnNpcPokemon(ByVal MapNum As Long, ByVal NpcIndex As Long, ByVal NpcPokeSlot As Byte)
Dim i As Byte, startPosX As Long, startPosY As Long
Dim foundPosition As Boolean
Dim x As Long, Y As Long

    '//Check for error
    If NpcIndex <= 0 Or NpcIndex > MAX_MAP_NPC Then Exit Sub
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If NpcPokeSlot <= 0 Or NpcPokeSlot > MAX_GAME_POKEMON Then Exit Sub
    If MapNpc(MapNum, NpcIndex).Num <= 0 Then Exit Sub
    If Npc(MapNpc(MapNum, NpcIndex).Num).PokemonNum(NpcPokeSlot) <= 0 Then Exit Sub
    
    With MapNpcPokemon(MapNum, NpcIndex)
        '//General
        .Num = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonNum(NpcPokeSlot)
        
        foundPosition = False
        For x = MapNpc(MapNum, NpcIndex).x - 1 To MapNpc(MapNum, NpcIndex).x + 1
            For Y = MapNpc(MapNum, NpcIndex).Y - 1 To MapNpc(MapNum, NpcIndex).Y + 1
                If x = MapNpc(MapNum, NpcIndex).x And Y = MapNpc(MapNum, NpcIndex).Y Then
                    
                Else
                    '//Check if OpenTile
                    If CheckOpenTile(MapNum, x, Y) Then
                        startPosX = x
                        startPosY = Y
                        foundPosition = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        If foundPosition Then
            '//Location
            .x = startPosX
            .Y = startPosY
        Else
            '//Location
            .x = MapNpc(MapNum, NpcIndex).x
            .Y = MapNpc(MapNum, NpcIndex).Y
        End If
        .Dir = DIR_DOWN
            
        '//Nature
        '//Random
        If Npc(MapNpc(MapNum, NpcIndex).Num).PokemonNature(NpcPokeSlot) = -1 Then
            .Nature = Random(0, PokemonNature.PokemonNature_Count - 1)
            '//Selected
        ElseIf Npc(MapNpc(MapNum, NpcIndex).Num).PokemonNature(NpcPokeSlot) >= 0 Then
            .Nature = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonNature(NpcPokeSlot)
        End If
        
        .IsShiny = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonIsShiny(NpcPokeSlot)
        
        .HeldItem = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonItem(NpcPokeSlot)
        
        .Gender = GENDER_MALE
        
        '//Status
        .Status = 0

        '//Level
        .Level = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonLevel(NpcPokeSlot)
        If .Level <= 1 Then .Level = 1
        If .Level >= MAX_LEVEL Then .Level = MAX_LEVEL

        '//Stats
        For i = 1 To StatEnum.Stat_Count - 1
            .Stat(i).EV = 0
            
            If Npc(MapNpc(MapNum, NpcIndex).Num).PokemonIvFull(NpcPokeSlot) = NO Then
                .Stat(i).IV = Random(1, 31)
            Else
                .Stat(i).IV = 31
            End If
            If .Stat(i).IV > 31 Then .Stat(i).IV = 31
            If .Stat(i).IV < 1 Then .Stat(i).IV = 1
            .Stat(i).Value = CalculatePokemonStat(i, .Num, .Level, .Stat(i).EV, .Stat(i).IV, .Nature)
        Next
            
        '//Vital
        .MaxHp = .Stat(StatEnum.HP).Value
        .CurHp = .MaxHp
        
        '//Moveset
        For i = 1 To MAX_MOVESET
            If Npc(MapNpc(MapNum, NpcIndex).Num).PokemonMoveset(NpcPokeSlot, i) > 0 Then
                .Moveset(i).Num = Npc(MapNpc(MapNum, NpcIndex).Num).PokemonMoveset(NpcPokeSlot, i)
                .Moveset(i).TotalPP = PokemonMove(.Moveset(i).Num).PP
                .Moveset(i).CurPP = PokemonMove(.Moveset(i).Num).PP
            End If
        Next
        
        '//Update Data to map
        SendNpcPokemonData MapNum, NpcIndex, YES, 0, .x, .Y
        
        .MoveTmr = GetTickCount + 1000
    End With
End Sub

Public Function NpcPokemonMove(ByVal MapNum As Long, ByVal NpcIndex As Long, ByVal Dir As Byte) As Boolean
Dim DidMove As Boolean
Dim expEarn As Long
Dim RndNum As Byte
Dim RandomNum As Long
Dim MoveSpeed As Long

    '//Exit out when error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Function
    If NpcIndex <= 0 Or NpcIndex > MAX_MAP_NPC Then Exit Function
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Function
    If MapNpc(MapNum, NpcIndex).CurPokemon <= 0 Then Exit Function
    
    DidMove = False

    With MapNpcPokemon(MapNum, NpcIndex)
        If .Num <= 0 Then Exit Function
        'If .QueueMove > 0 Then Exit Function
        If .Status = StatusEnum.Sleep Then
            If .Status = StatusEnum.Paralize Then
                .MoveTmr = GetTickCount + 1000
            Else
                MoveSpeed = CalculateSpeed(GetMapNpcPokemonStat(MapNum, NpcIndex, Spd))
                If MoveSpeed <= 250 Then MoveSpeed = 250
                .MoveTmr = GetTickCount + MoveSpeed
            End If
            Exit Function
        End If
        If .Status = StatusEnum.Frozen Then
            If .Status = StatusEnum.Paralize Then
                .MoveTmr = GetTickCount + 1000
            Else
                MoveSpeed = CalculateSpeed(GetMapNpcPokemonStat(MapNum, NpcIndex, Spd))
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
                    If Not CheckDirection(MapNum, DIR_UP, .x, .Y, True) Then
                        .Y = .Y - 1
                        DidMove = True
                    End If
                End If
            Case DIR_DOWN
                .Dir = DIR_DOWN
                
                '//Check to make sure not outside of boundries
                If .Y < Map(MapNum).MaxY Then
                    If Not CheckDirection(MapNum, DIR_DOWN, .x, .Y, True) Then
                        .Y = .Y + 1
                        DidMove = True
                    End If
                End If
            Case DIR_LEFT
                .Dir = DIR_LEFT
                
                '//Check to make sure not outside of boundries
                If .x > 0 Then
                    If Not CheckDirection(MapNum, DIR_LEFT, .x, .Y, True) Then
                        .x = .x - 1
                        DidMove = True
                    End If
                End If
            Case DIR_RIGHT
                .Dir = DIR_RIGHT
                
                '//Check to make sure not outside of boundries
                If .x < Map(MapNum).MaxX Then
                    If Not CheckDirection(MapNum, DIR_RIGHT, .x, .Y, True) Then
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
                            SendActionMsg MapNum, "-" & .StatusDamage, .x * 32, .Y * 32, Magenta
                            
                            MapNpc(MapNum, NpcIndex).PokemonAlive(MapNpc(MapNum, NpcIndex).CurPokemon) = NO
                            NpcPokemonCallBack MapNum, NpcIndex
                            Exit Function
                        Else
                            .CurHp = .CurHp - .StatusDamage
                            SendActionMsg MapNum, "-" & .StatusDamage, .x * 32, .Y * 32, Magenta
                            '//Update
                            'SendPokemonVital MapPokemonNum
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
                MoveSpeed = CalculateSpeed(GetMapNpcPokemonStat(MapNum, NpcIndex, Spd))
                If MoveSpeed <= 250 Then MoveSpeed = 250
                .MoveTmr = GetTickCount + MoveSpeed
            End If
            SendNpcPokemonMove MapNum, NpcIndex
            NpcPokemonMove = True
        Else
            '//Update Dir
            SendNpcPokemonDir MapNum, NpcIndex
            NpcPokemonMove = False
        End If
    End With
End Function

Public Sub NpcPokemonDir(ByVal MapNum As Long, ByVal NpcIndex As Long, ByVal Dir As Byte)
    '//Exit out when error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If NpcIndex <= 0 Or NpcIndex > MAX_MAP_NPC Then Exit Sub
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    If MapNpc(MapNum, NpcIndex).CurPokemon <= 0 Then Exit Sub
    
    If MapNpcPokemon(MapNum, NpcIndex).Num > 0 Then
        MapNpcPokemon(MapNum, NpcIndex).Dir = Dir
        '//Update Dir
        SendNpcPokemonDir MapNum, NpcIndex
    End If
End Sub

Public Sub ClearMapNpcPokemon(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpcPokemon(MapNum, MapNpcNum)), LenB(MapNpcPokemon(MapNum, MapNpcNum)))
End Sub

Public Sub NpcPokemonCallBack(ByVal MapNum As Long, ByVal MapNpcNum As Long, Optional ByVal DidFaint As Boolean = False)
    '//Exit out when error
    If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub
    If MapNpcPokemon(MapNum, MapNpcNum).Num <= 0 Then Exit Sub
    
    SendNpcPokemonData MapNum, MapNpcNum, YES, 1, MapNpcPokemon(MapNum, MapNpcNum).x, MapNpcPokemon(MapNum, MapNpcNum).Y
    Call ClearMapNpcPokemon(MapNum, MapNpcNum)
    If DidFaint Then
        MapNpc(MapNum, MapNpcNum).FaintWaitTimer = GetTickCount + 1500
    End If
End Sub

Public Sub UpdateRank(ByVal Name As String, ByVal Level As Long, ByVal Exp As Long)
Dim i As Long

    ' Check existing
    For i = 1 To MAX_RANK
        If LCase$(Trim$(Rank(i).Name)) = LCase$(Trim$(Name)) Then
            ' remove
            Rank(i).Name = vbNullString
            Rank(i).Level = 0
            Rank(i).Exp = 0
            ClearRank i
            Exit For
        End If
    Next
    
    For i = 1 To MAX_RANK
        If Level > Rank(i).Level Then
            MoveRank i
            Rank(i).Name = Name
            Rank(i).Level = Level
            Rank(i).Exp = Exp
            'SaveRank
            Exit Sub
        ElseIf Level = Rank(i).Level Then
            If Exp > Rank(i).Exp Then
                MoveRank i
                Rank(i).Name = Name
                Rank(i).Level = Level
                Rank(i).Exp = Exp
                'SaveRank
                Exit Sub
            End If
        End If
    Next
    
    SendRankToAll
    'SaveRank
End Sub

Public Sub MoveRank(ByVal StartNum As Long)
Dim i As Long

    For i = 9 To StartNum Step -1
        Rank(i + 1) = Rank(i)
    Next
End Sub

Public Sub ClearRank(ByVal RankNum As Long)
Dim i As Long

    For i = RankNum To 9
        Rank(i) = Rank(i + 1)
    Next
End Sub

Public Function FindNpcDropSlotHaveItem(ByVal Index As Long, ByVal TargetIndex As Long) As Byte
    Dim i As Byte
    
    For i = 1 To MAX_DROP
        If Pokemon(MapPokemon(TargetIndex).Num).DropNum(i) > 0 Then
            FindNpcDropSlotHaveItem = FindNpcDropSlotHaveItem + 1
        End If
    Next i
End Function

Public Function GetAtributeName(Atributte As StatEnum) As String
    Select Case Atributte
    Case StatEnum.HP: GetAtributeName = "HP"
    Case StatEnum.Atk: GetAtributeName = "Atk"
    Case StatEnum.Def: GetAtributeName = "Def"
    Case StatEnum.SpAtk: GetAtributeName = "SpAtk"
    Case StatEnum.SpDef: GetAtributeName = "SpDef"
    Case StatEnum.Spd: GetAtributeName = "Spd"
    End Select
End Function

Public Sub BuyInvSlot(ByVal Index As Long, ByVal InvNum As Byte)
    If InvNum <= 0 Or InvNum > MAX_PLAYER_INV Then
        Exit Sub
    End If
    If TempPlayer(Index).UseChar <= 0 Then
        Exit Sub
    End If

    With PlayerInv(Index).Data(InvNum)
        If .Locked = YES Then
            If IsNextInvSlotToBuy(Index, InvNum) = True Then
                If Player(Index, TempPlayer(Index).UseChar).Cash >= INV_SLOTS_PRICE Then

                    .Locked = NO
                    Player(Index, TempPlayer(Index).UseChar).Cash = Player(Index, TempPlayer(Index).UseChar).Cash - INV_SLOTS_PRICE
                    Call SendPlayerInvSlot(Index, InvNum)
                    Call SendPlayerCash(Index)

                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Você desbloqueou o slot " & InvNum & " da sua bag", White
                    Case LANG_EN: AddAlert Index, "Você desbloqueou o slot " & InvNum & " da sua bag", White
                    Case LANG_ES: AddAlert Index, "Você desbloqueou o slot " & InvNum & " da sua bag", White
                    End Select
                Else
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Você não possui " & INV_SLOTS_PRICE & " de Cash", White
                    Case LANG_EN: AddAlert Index, "Você não possui " & INV_SLOTS_PRICE & " de Cash", White
                    Case LANG_ES: AddAlert Index, "Você não possui " & INV_SLOTS_PRICE & " de Cash", White
                    End Select
                End If
            Else
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "O Slot " & InvNum & " da sua bag já está liberado", White
                Case LANG_EN: AddAlert Index, "O Slot " & InvNum & " da sua bag já está liberado", White
                Case LANG_ES: AddAlert Index, "O Slot " & InvNum & " da sua bag já está liberado", White
                End Select
            End If
        End If
    End With
End Sub

Private Function IsNextInvSlotToBuy(ByVal Index As Long, ByVal i As Byte) As Boolean
    Dim l As Byte

    IsNextInvSlotToBuy = False
    For l = 1 To MAX_PLAYER_INV
        If PlayerInv(Index).Data(l).Locked = YES Then
            If l = i Then
                IsNextInvSlotToBuy = True
                Exit Function
            Else
                Exit Function
            End If
        End If
    Next l
End Function

Public Sub PlayerWinToNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long)
    Dim NpcNum As Long, MapNum As Long

    MapNum = Player(Attacker, TempPlayer(Attacker).UseChar).Map

    NpcNum = MapNpc(MapNum, MapNpcNum).Num

    SendActionMsg MapNum, "Win!", Player(Attacker, TempPlayer(Attacker).UseChar).x * 32, Player(Attacker, TempPlayer(Attacker).UseChar).Y * 32, White
    Select Case TempPlayer(Attacker).CurLanguage
    Case LANG_PT: AddAlert Attacker, "You win on a duel!", White
    Case LANG_EN: AddAlert Attacker, "You win on a duel!", White
    Case LANG_ES: AddAlert Attacker, "You win on a duel!", White
    End Select

    TempPlayer(Attacker).InNpcDuel = 0
    TempPlayer(Attacker).DuelTime = 0
    TempPlayer(Attacker).DuelTimeTmr = 0
    TempPlayer(Attacker).WarningTimer = 0
    SendPlayerNpcDuel Attacker
    '//Send Reward
    If Npc(NpcNum).Reward > 0 Then
        Player(Attacker, TempPlayer(Attacker).UseChar).Money = Player(Attacker, TempPlayer(Attacker).UseChar).Money + Npc(NpcNum).Reward
        If Player(Attacker, TempPlayer(Attacker).UseChar).Money >= MAX_MONEY Then
            Player(Attacker, TempPlayer(Attacker).UseChar).Money = MAX_MONEY
        End If
        SendPlayerData Attacker
        AddAlert Attacker, "You got $" & Npc(NpcNum).Reward, White
        If Npc(NpcNum).RewardExp > 0 Then
            GivePlayerExp Attacker, Npc(NpcNum).RewardExp
        End If
    End If
    If Npc(NpcNum).WinEvent > 0 Then
        TempPlayer(Attacker).CurConvoNum = Npc(NpcNum).WinEvent
        TempPlayer(Attacker).CurConvoData = 0    '//Always start at 0
        TempPlayer(Attacker).CurConvoNpc = NpcNum
        TempPlayer(Attacker).CurConvoMapNpc = MapNpcNum
        ProcessConversation Attacker, TempPlayer(Attacker).CurConvoNum, TempPlayer(Attacker).CurConvoData, TempPlayer(Attacker).CurConvoNpc
    End If

    '//Não pode rebatalhar se ganhar.
    Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).Win = NO
    
    If Npc(NpcNum).Rebatle = REBATLE_NONE Or Npc(NpcNum).Rebatle = REBATLE_LOSE Then    '//Não pode rebatalhar neste dia, se ganhar
            Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = Day(Date)
            Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = Month(Date)
    ElseIf Npc(NpcNum).Rebatle = REBATLE_NEVER Then
        Player(Attacker, TempPlayer(Attacker).UseChar).NpcBattledDay(NpcNum).Win = YES
    End If
End Sub

Public Sub PlayerLoseToNpc(ByVal Victim As Long, ByVal MapNpcNum As Long)
    Dim NpcNum As Long, MapNum As Long
    
    MapNum = Player(Victim, TempPlayer(Victim).UseChar).Map

    NpcNum = MapNpc(MapNum, MapNpcNum).Num
    
    If MapNpcNum > 0 Then
        MapNpc(MapNum, MapNpcNum).InBattle = 0
        NpcPokemonCallBack MapNum, MapNpcNum
        SendActionMsg MapNum, "Lose!", Player(Victim, TempPlayer(Victim).UseChar).x * 32, Player(Victim, TempPlayer(Victim).UseChar).Y * 32, White
        TempPlayer(Victim).InNpcDuel = 0
        TempPlayer(Victim).DuelTime = 0
        TempPlayer(Victim).DuelTimeTmr = 0
        TempPlayer(Victim).WarningTimer = 0
        SendPlayerNpcDuel Victim

        If Npc(NpcNum).Rebatle = REBATLE_NONE Then    '//Não pode rebatalhar se perder
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = Day(Date)
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = Month(Date)
        ElseIf Npc(NpcNum).Rebatle = REBATLE_LOSE Or Npc(NpcNum).Rebatle = REBATLE_NEVER Then    '//Pode rebatalhar se perder
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledDay(NpcNum).NpcBattledAt = 0
            Player(Victim, TempPlayer(Victim).UseChar).NpcBattledMonth(NpcNum).NpcBattledAt = 0
        End If
    End If
End Sub

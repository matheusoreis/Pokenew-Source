Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SCharacters) = GetAddress(AddressOf HandleCharacters)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SMap) = GetAddress(AddressOf HandleMap)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SLeftGame) = GetAddress(AddressOf HandleLeftGame)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SSpawnMapNpc) = GetAddress(AddressOf HandleSpawnMapNpc)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPokemonData) = GetAddress(AddressOf HandlePokemonData)
    HandleDataSub(SPokemonHighIndex) = GetAddress(AddressOf HandlePokemonHighIndex)
    HandleDataSub(SPokemonMove) = GetAddress(AddressOf HandlePokemonMove)
    HandleDataSub(SPokemonDir) = GetAddress(AddressOf HandlePokemonDir)
    HandleDataSub(SPokemonVital) = GetAddress(AddressOf HandlePokemonVital)
    HandleDataSub(SChatbubble) = GetAddress(AddressOf HandleChatbubble)
    HandleDataSub(SPlayerPokemonData) = GetAddress(AddressOf HandlePlayerPokemonData)
    HandleDataSub(SPlayerPokemonMove) = GetAddress(AddressOf HandlePlayerPokemonMove)
    HandleDataSub(SPlayerPokemonXY) = GetAddress(AddressOf HandlePlayerPokemonXY)
    HandleDataSub(SPlayerPokemonDir) = GetAddress(AddressOf HandlePlayerPokemonDir)
    HandleDataSub(SPlayerPokemonVital) = GetAddress(AddressOf HandlePlayerPokemonVital)
    HandleDataSub(SPlayerPokemonPP) = GetAddress(AddressOf HandlePlayerPokemonPP)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvSlot) = GetAddress(AddressOf HandlePlayerInvSlot)
    HandleDataSub(SPlayerPokemons) = GetAddress(AddressOf HandlePlayerPokemons)
    HandleDataSub(SPlayerPokemonSlot) = GetAddress(AddressOf HandlePlayerPokemonSlot)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SPlayAnimation) = GetAddress(AddressOf HandlePlayAnimation)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SNewMove) = GetAddress(AddressOf HandleNewMove)
    HandleDataSub(SGetData) = GetAddress(AddressOf HandleGetData)
    HandleDataSub(SMapPokemonCatchState) = GetAddress(AddressOf HandleMapPokemonCatchState)
    HandleDataSub(SPlayerVital) = GetAddress(AddressOf HandlePlayerVital)
    HandleDataSub(SPlayerInvStorage) = GetAddress(AddressOf HandlePlayerInvStorage)
    HandleDataSub(SPlayerInvStorageSlot) = GetAddress(AddressOf HandlePlayerInvStorageSlot)
    HandleDataSub(SPlayerPokemonStorage) = GetAddress(AddressOf HandlePlayerPokemonStorage)
    HandleDataSub(SPlayerPokemonStorageSlot) = GetAddress(AddressOf HandlePlayerPokemonStorageSlot)
    HandleDataSub(SStorage) = GetAddress(AddressOf HandleStorage)
    HandleDataSub(SInitConvo) = GetAddress(AddressOf HandleInitConvo)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SRequest) = GetAddress(AddressOf HandleRequest)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SOpenTrade) = GetAddress(AddressOf HandleOpenTrade)
    HandleDataSub(SUpdateTradeItem) = GetAddress(AddressOf HandleUpdateTradeItem)
    HandleDataSub(STradeUpdateMoney) = GetAddress(AddressOf HandleTradeUpdateMoney)
    HandleDataSub(SSetTradeState) = GetAddress(AddressOf HandleSetTradeState)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(SPlayerPokedex) = GetAddress(AddressOf HandlePlayerPokedex)
    HandleDataSub(SPlayerPokedexSlot) = GetAddress(AddressOf HandlePlayerPokedexSlot)
    HandleDataSub(SPokemonStatus) = GetAddress(AddressOf HandlePokemonStatus)
    HandleDataSub(SMapNpcPokemonStatus) = GetAddress(AddressOf HandleMapNpcPokemonStatus)
    HandleDataSub(SPlayerPokemonStatus) = GetAddress(AddressOf HandlePlayerPokemonStatus)
    HandleDataSub(SClearPlayer) = GetAddress(AddressOf HandleClearPlayer)
    HandleDataSub(SPlayerPokemonsStat) = GetAddress(AddressOf HandlePlayerPokemonsStat)
    HandleDataSub(SPlayerPokemonStatBuff) = GetAddress(AddressOf HandlePlayerPokemonStatBuff)
    HandleDataSub(SPlayerStatus) = GetAddress(AddressOf HandlePlayerStatus)
    HandleDataSub(SWeather) = GetAddress(AddressOf HandleWeather)
    HandleDataSub(SNpcPokemonData) = GetAddress(AddressOf HandleNpcPokemonData)
    HandleDataSub(SNpcPokemonMove) = GetAddress(AddressOf HandleNpcPokemonMove)
    HandleDataSub(SNpcPokemonDir) = GetAddress(AddressOf HandleNpcPokemonDir)
    HandleDataSub(SNpcPokemonVital) = GetAddress(AddressOf HandleNpcPokemonVital)
    HandleDataSub(SPlayerNpcDuel) = GetAddress(AddressOf HandlePlayerNpcDuel)
    HandleDataSub(SRelearnMove) = GetAddress(AddressOf HandleReleaseMove)
    HandleDataSub(SPlayerAction) = GetAddress(AddressOf HandlePlayerAction)
    HandleDataSub(SPlayerExp) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SParty) = GetAddress(AddressOf HandleParty)
    '//Editors
    HandleDataSub(SInitMap) = GetAddress(AddressOf HandleInitMap)
    HandleDataSub(SInitNpc) = GetAddress(AddressOf HandleInitNpc)
    HandleDataSub(SNpcs) = GetAddress(AddressOf HandleNpcs)
    HandleDataSub(SInitPokemon) = GetAddress(AddressOf HandleInitPokemon)
    HandleDataSub(SPokemons) = GetAddress(AddressOf HandlePokemons)
    HandleDataSub(SInitItem) = GetAddress(AddressOf HandleInitItem)
    HandleDataSub(SItems) = GetAddress(AddressOf HandleItems)
    HandleDataSub(SInitPokemonMove) = GetAddress(AddressOf HandleInitPokemonMove)
    HandleDataSub(SPokemonMoves) = GetAddress(AddressOf HandlePokemonMoves)
    HandleDataSub(SInitAnimation) = GetAddress(AddressOf HandleInitAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SInitSpawn) = GetAddress(AddressOf HandleInitSpawn)
    HandleDataSub(SSpawn) = GetAddress(AddressOf HandleSpawn)
    HandleDataSub(SInitConversation) = GetAddress(AddressOf HandleInitConversation)
    HandleDataSub(SConversation) = GetAddress(AddressOf HandleConversation)
    HandleDataSub(SInitShop) = GetAddress(AddressOf HandleInitShop)
    HandleDataSub(SShop) = GetAddress(AddressOf HandleShop)
    HandleDataSub(SInitQuest) = GetAddress(AddressOf HandleInitQuest)
    HandleDataSub(SQuest) = GetAddress(AddressOf HandleQuest)
    HandleDataSub(SRank) = GetAddress(AddressOf HandleRank)
    HandleDataSub(SDataLimit) = GetAddress(AddressOf HandleDataLimit)
    HandleDataSub(SPlayerPvP) = GetAddress(AddressOf HandlePlayerPvP)
    HandleDataSub(SPlayerCash) = GetAddress(AddressOf HandlePlayerCash)
    HandleDataSub(SRequestCash) = GetAddress(AddressOf HandleRequestCash)
    HandleDataSub(SEventInfo) = GetAddress(AddressOf HandleEventInfo)
    HandleDataSub(SRequestServerInfo) = GetAddress(AddressOf HandleRequestServerInfo)
    HandleDataSub(SClientTime) = GetAddress(AddressOf HandleClientTime)
    HandleDataSub(SSendVirtualShop) = GetAddress(AddressOf HandleVirtualShop)
    HandleDataSub(SFishMode) = GetAddress(AddressOf HandleFishMode)
    HandleDataSub(SMapReport) = GetAddress(AddressOf HandleMapReport)
End Sub

Public Sub HandleData(ByRef data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        UnloadMain
        Exit Sub
    End If
    If MsgType >= SMSG_Count Then
        UnloadMain
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), MyIndex, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Public Sub IncomingData(ByVal dataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    frmMain.Socket.GetData buffer, vbUnicode, dataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    On Error GoTo errorHandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
    
    Exit Sub
errorHandler:
    Ping = 5000
    PingEnd = 0
    PingStart = 0
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player_HighIndex = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleAlertMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim willDisconnect As Byte
Dim NotHideLoad As Byte
Dim alertText As String, alertColor As Long
Dim showAlert As Boolean

    showAlert = True
    '//Update client's index
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    alertText = Trim$(buffer.ReadString)
    alertColor = buffer.ReadLong
    willDisconnect = buffer.ReadByte
    NotHideLoad = buffer.ReadByte
    Set buffer = Nothing
    
    If NotHideLoad = NO Then
        '//Hide Loading Screen
        SetStatus False
    Else
        If IsLoading Then
            SetStatus True, alertText
            showAlert = False
        End If
    End If
    
    If showAlert Then
        AddAlert alertText, alertColor
    End If
    
    If willDisconnect = YES Then
        DestroyTCP
        '//Check if currently logged in
        If IsLoggedIn Then
            ResetMenu
        End If
    End If
End Sub

Private Sub HandleLoginOk(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Data1 As Byte, InputIndex As Long

    '//Hide Loading Screen
    SetStatus False
    
    '//Update client's index
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InputIndex = buffer.ReadLong
    Data1 = buffer.ReadByte
    Set buffer = Nothing
    
    If Data1 = 0 Then
        MyIndex = InputIndex
        '//Check save password
        If GameSetting.SavePass = YES Then
            GameSetting.Username = Trim$(User)
            GameSetting.Password = Trim$(Pass)
            SaveSetting
        End If
    End If
    
    IsLoggedIn = True
    WaitTimer = 0
    
    '//Hide possible open'ed gui
    GuiState GUI_LOGIN, False
    GuiState GUI_REGISTER, False
    GuiState GUI_CHARACTERCREATE, False
    
    CurChar = 1
    '//Open CharacterSelection window
    GuiState GUI_CHARACTERSELECT, True
End Sub

Private Sub HandleCharacters(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_PLAYERCHAR
        pCharName(i) = Trim$(buffer.ReadString)
        pCharSprite(i) = buffer.ReadLong
        If pCharSprite(i) > 0 And Len(pCharName(i)) > 0 Then
            pCharInUsed(i) = True
        Else
            pCharInUsed(i) = False
        End If
    Next
    Set buffer = Nothing
End Sub

Private Sub HandleInGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    '//Hide possible open'ed gui
    GuiState GUI_LOGIN, False
    GuiState GUI_REGISTER, False
    GuiState GUI_CHARACTERCREATE, False
    GuiState GUI_CHARACTERSELECT, False

    '//Open Game Gui
    GuiState GUI_CHATBOX, True
    ChatTab = "/map"

    '//Hide Loading Screen
    SetStatus False

    WaitTimer = 0

    InitFade 0, FadeIn, 4

    For i = 1 To MAX_POKEMON
        If Len(Trim$(Pokemon(i).Name)) > 0 Then
            PokedexHighIndex = i
            MaxPokedexViewLine = (PokedexHighIndex / 8)
            MaxPokedexViewLine = MaxPokedexViewLine - 3
        End If
    Next

    For i = 1 To MAX_RANK
        If Len(Trim$(Rank(i).Name)) > 0 Then
            RankingHighIndex = i
            RankingMaxViewLine = RankingHighIndex
            RankingMaxViewLine = RankingMaxViewLine - RankingScrollViewLine
        End If
    Next
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, X As Byte
Dim isPvP As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    
    If i <= 0 Or i > MAX_PLAYER Then Exit Sub
    
    With Player(i)
        .Name = Trim$(buffer.ReadString)
        .Sprite = buffer.ReadLong
        .Access = buffer.ReadByte
        .Map = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        .CurHP = buffer.ReadLong
        .Money = buffer.ReadLong
        .TempSprite = buffer.ReadLong
        .TempSpriteID = buffer.ReadLong
        .TempSpritePassiva = buffer.ReadLong
        For X = 1 To MAX_BADGE
            .Badge(X) = buffer.ReadByte
        Next
        .Level = buffer.ReadLong
        .CurExp = buffer.ReadLong
        For X = 1 To MAX_HOTBAR
            .Hotbar(X).Num = buffer.ReadLong
        Next
        .StealthMode = buffer.ReadByte
        'isPvP = buffer.ReadByte
        .win = buffer.ReadLong
        .Lose = buffer.ReadLong
        .Tie = buffer.ReadLong
        
        .Cash = buffer.ReadLong
        
        .Started = CDate(buffer.ReadString)
        .TimePlay = buffer.ReadLong
        
        .FishMode = buffer.ReadByte
        .FishRod = buffer.ReadByte
        
        '//Prevent from moving
        .Moving = NO
        .xOffset = 0
        .yOffset = 0
    End With
    
    Set buffer = Nothing
End Sub

Private Sub HandleMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long
Dim X As Long, Y As Long
Dim i As Long, a As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNum = buffer.ReadLong
    
    With Map
        '//General
        .Revision = buffer.ReadLong
        .Name = Trim$(buffer.ReadString)
        .Moral = buffer.ReadByte
        
        '//Size
        .MaxX = buffer.ReadLong
        .MaxY = buffer.ReadLong
        If .MaxX < MAX_MAPX Then .MaxX = MAX_MAPX
        If .MaxY < MAX_MAPY Then .MaxY = MAX_MAPY
        
        '//Redim size
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)
    End With
    
    '//Tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                '//Layer
                For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                    For a = MapLayerType.Normal To MapLayerType.Animated
                        .Layer(i, a).Tile = buffer.ReadLong
                        .Layer(i, a).TileX = buffer.ReadLong
                        .Layer(i, a).TileY = buffer.ReadLong
                        '//Map Anim
                        .Layer(i, a).MapAnim = buffer.ReadLong
                    Next
                Next
                '//Tile Data
                .Attribute = buffer.ReadByte
                .Data1 = buffer.ReadLong
                .Data2 = buffer.ReadLong
                .Data3 = buffer.ReadLong
                .Data4 = buffer.ReadLong
            End With
        Next
    Next
    
    With Map
        '//Map Link
        .LinkUp = buffer.ReadLong
        .LinkDown = buffer.ReadLong
        .LinkLeft = buffer.ReadLong
        .LinkRight = buffer.ReadLong
        
        '//Map Data
        .Music = Trim$(buffer.ReadString)
        
        '//Npc
        For i = 1 To MAX_MAP_NPC
            .Npc(i) = buffer.ReadLong
        Next
        
        '//Moral
        .KillPlayer = buffer.ReadByte
        .IsCave = buffer.ReadByte
        .CaveLight = buffer.ReadByte
        .SpriteType = buffer.ReadByte
        .StartWeather = buffer.ReadByte
        .NoCure = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    SaveMap MapNum
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long, Rev As Long
Dim NeedMap As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNum = buffer.ReadLong
    '//Check for MapVersion
    Rev = buffer.ReadLong
    Set buffer = Nothing
    
    GettingMap = True
    
    '//Clear Temporary Map Data
    'ClearMap
    ClearMapNpcs
    ClearMapPokemons
    ClearChatbubble
    ClearMapNpcPokemons
    
    '//Check map version on cache
    If CheckRev(MapNum, Rev) Then
        '//Version matched, no need for update
        NeedMap = NO
        '//Load the map cache
        LoadMap MapNum
    Else
        '//Version did not match, need for update
        NeedMap = YES
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteByte NeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleMapDone(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    GettingMap = False
    CanMoveNow = True
    
    '//get the npc high index
    For i = MAX_MAP_NPC To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If
    Next
    
    '//Map Music
    If Trim$(Map.Music) <> "None." Then
        If CurMusic <> Map.Music Then
            PlayMusic Trim$(Map.Music), False, True
        End If
    Else
        StopMusic True
    End If
    
    '//Clear Pokeball
    For i = 1 To MAX_GAME_POKEMON
        CatchBall(i).InUsed = False
    Next

    For i = 1 To 255
        ClearAnimInstance (i)
        ClearActionMsg (i)
    Next
    Action_HighIndex = 1
    
    '//Clear
    ClearSelMenu
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With Player(thePlayer)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES
        
        '//Set offset
        Select Case .Dir
            Case DIR_UP
                .yOffset = TILE_Y
            Case DIR_DOWN
                .yOffset = TILE_Y * -1
            Case DIR_LEFT
                .xOffset = TILE_X
            Case DIR_RIGHT
                .xOffset = TILE_X * -1
        End Select
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With Player(thePlayer)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong

        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With Player(thePlayer)
        .Dir = buffer.ReadByte

        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleLeftGame(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ClearPlayer (buffer.ReadLong)
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String, Colour As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Msg = buffer.ReadString
    Colour = buffer.ReadLong
    Set buffer = Nothing
    
    If InStr(1, Msg, ColourChar) > 0 Then
        Debug.Print "Encontrou"
    End If
    
    AddText KeepTwoDigit(Hour(time)) & ":" & KeepTwoDigit(Minute(time)) & " " & Trim$(Msg), Colour
End Sub

Private Sub HandleSpawnMapNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNpcNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    With MapNpc(MapNpcNum)
        '//General
        .Num = buffer.ReadLong
        
        '//Location
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    If MapNpcNum > Npc_HighIndex Then
        Npc_HighIndex = MapNpcNum
        '//make sure we're not overflowing
        If Npc_HighIndex > MAX_MAP_NPC Then Npc_HighIndex = MAX_MAP_NPC
    End If
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_MAP_NPC
        With MapNpc(i)
            '//General
            .Num = buffer.ReadLong
            
            '//Location
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadByte
        End With
    Next
    Set buffer = Nothing
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNpcNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    With MapNpc(MapNpcNum)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES
        
        '//Set offset
        Select Case .Dir
            Case DIR_UP
                .yOffset = TILE_Y
            Case DIR_DOWN
                .yOffset = TILE_Y * -1
            Case DIR_LEFT
                .xOffset = TILE_X
            Case DIR_RIGHT
                .xOffset = TILE_X * -1
        End Select
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNpcNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapNpcNum = buffer.ReadLong
    With MapNpc(MapNpcNum)
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokemonIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    PokemonIndex = buffer.ReadLong
    With MapPokemon(PokemonIndex)
        '//General
        .Num = buffer.ReadLong

        '//Location
        .Map = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Vital
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
        
        '//Shiny
        .IsShiny = buffer.ReadByte
        
        '//Happiness
        .Happiness = buffer.ReadByte
        
        '//Gender
        .Gender = buffer.ReadByte
        
        '//Status
        .Status = buffer.ReadByte
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonHighIndex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Pokemon_HighIndex = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    With MapPokemon(MapPokeNum)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES
        
        '//Set offset
        Select Case .Dir
            Case DIR_UP
                .yOffset = TILE_Y
            Case DIR_DOWN
                .yOffset = TILE_Y * -1
            Case DIR_LEFT
                .xOffset = TILE_X
            Case DIR_RIGHT
                .xOffset = TILE_X * -1
        End Select
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    With MapPokemon(MapPokeNum)
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    With MapPokemon(MapPokeNum)
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleChatbubble(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    AddChatBubble buffer.ReadLong, buffer.ReadByte, Trim$(buffer.ReadString), buffer.ReadLong, buffer.ReadLong, buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim thePlayer As Long
    Dim initState As Byte
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .Init = buffer.ReadByte
        .State = buffer.ReadByte

        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte

        .Slot = buffer.ReadByte

        '//Stat
        For i = 1 To StatEnum.Stat_Count - 1
            .Stat(i) = buffer.ReadLong
            .StatIV(i) = buffer.ReadLong
            .StatEV(i) = buffer.ReadLong
        Next

        '//Vital
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong

        '//Shiny
        .IsShiny = buffer.ReadByte

        '//Happiness
        .Happiness = buffer.ReadByte

        '//Gender
        .Gender = buffer.ReadByte

        '//Status
        .Status = buffer.ReadByte

        '//Held Item
        .HeldItem = buffer.ReadLong

        '//Ball Used
        .BallUsed = buffer.ReadByte

        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES

        '//Ball Location
        .BallX = buffer.ReadLong
        .BallY = buffer.ReadLong

        '//Init
        If .Init = YES Then
            If .State = 0 Then
                .Frame = 0
            ElseIf .State = 1 Then
                .Frame = 2
            End If
            .FrameState = 0
            .FrameTimer = GetTickCount + 100
        End If
        Set buffer = Nothing

        If thePlayer = MyIndex Then
            '//Reset Set Move
            SetAttackMove = 0

            '//Cries Sound Apenas pro MyIndex
            If .Num > 0 Then
                If .Init = YES Then
                    '//Play
                    If Trim$(Pokemon(.Num).Sound) <> "None." Then
                        PlayMusic Trim$(Pokemon(.Num).Sound), True, False
                    End If
                End If
            End If
        End If

    End With
End Sub

Private Sub HandlePlayerPokemonMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES
        
        '//Set offset
        Select Case .Dir
            Case DIR_UP
                .yOffset = TILE_Y
            Case DIR_DOWN
                .yOffset = TILE_Y * -1
            Case DIR_LEFT
                .xOffset = TILE_X
            Case DIR_RIGHT
                .xOffset = TILE_X * -1
        End Select
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonXY(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong

        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .Dir = buffer.ReadByte

        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .Slot = buffer.ReadLong
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
    End With
    Set buffer = Nothing
    
    If thePlayer = MyIndex Then
        If PlayerPokemon(MyIndex).Slot > 0 Then
            PlayerPokemons(PlayerPokemon(MyIndex).Slot).CurHP = PlayerPokemon(MyIndex).CurHP
            PlayerPokemons(PlayerPokemon(MyIndex).Slot).MaxHP = PlayerPokemon(MyIndex).MaxHP
        End If
    End If
End Sub

Private Sub HandlePlayerPokemonPP(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MoveSlot = buffer.ReadByte
    With PlayerPokemon(MyIndex)
        .Slot = buffer.ReadLong
        If .Slot > 0 Then
            PlayerPokemons(.Slot).Moveset(MoveSlot).CurPP = buffer.ReadLong
            PlayerPokemons(.Slot).Moveset(MoveSlot).TotalPP = buffer.ReadLong
        End If
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerInv(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    Dim NoNext As Boolean, CD As Long, Y As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_PLAYER_INV
        With PlayerInv(i)
            .Num = buffer.ReadLong
            .value = buffer.ReadLong
            .Status.Locked = buffer.ReadByte
            .Status.Opacity = 150
            CD = buffer.ReadLong
            
            '//Inv
            .ItemCooldown = (CD \ 1000)    ' CONVERTE PRA SEGUNDOS NO CLIENT
            '//Hotbar
            For Y = 1 To MAX_HOTBAR
                If Player(MyIndex).Hotbar(Y).Num = .Num Then
                    Player(MyIndex).Hotbar(Y).TmrCooldown = .ItemCooldown
                End If
            Next Y

            If Not NoNext Then
                If .Status.Locked = YES Then
                    .Status.Opacity = 255
                    NoNext = True
                End If
            End If
        End With
    Next
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerInvSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Byte
    Dim VarLocked As Byte, CD As Long, Y As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadByte
    With PlayerInv(Slot)
        .Num = buffer.ReadLong
        .value = buffer.ReadLong
        VarLocked = buffer.ReadByte
        CD = buffer.ReadLong
        
        '//Inv
            .ItemCooldown = (CD \ 1000)    ' CONVERTE PRA SEGUNDOS NO CLIENT
            '//Hotbar
            For Y = 1 To MAX_HOTBAR
                If Player(MyIndex).Hotbar(Y).Num = .Num Then
                    Player(MyIndex).Hotbar(Y).TmrCooldown = .ItemCooldown
                End If
            Next Y

        If VarLocked <> .Status.Locked Then
            .Status.Locked = VarLocked
            If .Status.Locked = NO Then
                If Slot >= MAX_PLAYER_INV Then Exit Sub
                Slot = Slot + 1

                PlayerInv(Slot).Status.Locked = YES
                PlayerInv(Slot).Status.Opacity = 255
            End If
        End If
    End With

    Set buffer = Nothing
End Sub

Private Sub HandlePlayerInvStorage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Byte, Y As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For X = 1 To MAX_STORAGE_SLOT
        PlayerInvStorage(X).Unlocked = buffer.ReadByte
        For Y = 1 To MAX_STORAGE
            With PlayerInvStorage(X)
                .data(Y).Num = buffer.ReadLong
                .data(Y).value = buffer.ReadLong
            End With
        Next
    Next
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerInvStorageSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte, sData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadByte
    sData = buffer.ReadByte
    With PlayerInvStorage(Slot)
        .data(sData).Num = buffer.ReadLong
        .data(sData).value = buffer.ReadLong
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonStorage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Byte, Y As Byte, z As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For X = 1 To MAX_STORAGE_SLOT
        PlayerPokemonStorage(X).Unlocked = buffer.ReadByte
        For Y = 1 To MAX_STORAGE
            With PlayerPokemonStorage(X)
                .data(Y).Num = buffer.ReadLong
                
                '//Stats
                .data(Y).Level = buffer.ReadByte
                For z = 1 To StatEnum.Stat_Count - 1
                    .data(Y).Stat(z) = buffer.ReadLong
                    .data(Y).StatIV(z) = buffer.ReadLong
                    .data(Y).StatEV(z) = buffer.ReadLong
                Next
                
                '//Vital
                .data(Y).CurHP = buffer.ReadLong
                .data(Y).MaxHP = buffer.ReadLong
                
                '//Nature
                .data(Y).Nature = buffer.ReadByte
                
                '//Shiny
                .data(Y).IsShiny = buffer.ReadByte
                
                '//Happiness
                .data(Y).Happiness = buffer.ReadByte
                
                '//Gender
                .data(Y).Gender = buffer.ReadByte
                
                '//Status
                .data(Y).Status = buffer.ReadByte
                
                '//Exp
                .data(Y).CurExp = buffer.ReadLong
                .data(Y).NextExp = buffer.ReadLong
                
                '//Moveset
                For z = 1 To MAX_MOVESET
                    .data(Y).Moveset(z).Num = buffer.ReadLong
                    .data(Y).Moveset(z).CurPP = buffer.ReadLong
                    .data(Y).Moveset(z).TotalPP = buffer.ReadLong
                Next
                
                '//Ball Used
                .data(Y).BallUsed = buffer.ReadByte
                
                .data(Y).HeldItem = buffer.ReadLong
            End With
        Next
    Next
    Set buffer = Nothing
    DoEvents
End Sub

Private Sub HandlePlayerPokemonStorageSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte, sData As Byte
Dim X As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadByte
    sData = buffer.ReadByte
    With PlayerPokemonStorage(Slot)
        .data(sData).Num = buffer.ReadLong
                
        '//Stats
        .data(sData).Level = buffer.ReadByte
        For X = 1 To StatEnum.Stat_Count - 1
            .data(sData).Stat(X) = buffer.ReadLong
            .data(sData).StatIV(X) = buffer.ReadLong
            .data(sData).StatEV(X) = buffer.ReadLong
        Next
                
        '//Vital
        .data(sData).CurHP = buffer.ReadLong
        .data(sData).MaxHP = buffer.ReadLong
                
        '//Nature
        .data(sData).Nature = buffer.ReadByte
        
        '//Shiny
        .data(sData).IsShiny = buffer.ReadByte
        
        '//Happiness
        .data(sData).Happiness = buffer.ReadByte
        
        '//Gender
        .data(sData).Gender = buffer.ReadByte
        
        '//Status
        .data(sData).Status = buffer.ReadByte
                
        '//Exp
        .data(sData).CurExp = buffer.ReadLong
        .data(sData).NextExp = buffer.ReadLong
                
        '//Moveset
        For X = 1 To MAX_MOVESET
            .data(sData).Moveset(X).Num = buffer.ReadLong
            .data(sData).Moveset(X).CurPP = buffer.ReadLong
            .data(sData).Moveset(X).TotalPP = buffer.ReadLong
        Next
                
        '//Ball Used
        .data(sData).BallUsed = buffer.ReadByte
        
        .data(sData).HeldItem = buffer.ReadLong
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleStorage(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    StorageType = buffer.ReadByte
    Set buffer = Nothing
    
    InvCurSlot = 1
    PokemonCurSlot = 1
    
    If StorageType = 1 Then '//Inv
        '//Open Inventory Storage
        If Not GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
            GuiState GUI_INVSTORAGE, True
        End If
        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
            GuiState GUI_POKEMONSTORAGE, False
        End If
    ElseIf StorageType = 2 Then '//Pokemon
        '//Open Pokemon Storage
        If Not GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
            GuiState GUI_POKEMONSTORAGE, True
        End If
        If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
            GuiState GUI_INVSTORAGE, False
        End If
    Else
        If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
            GuiState GUI_INVSTORAGE, False
        End If
        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
            GuiState GUI_POKEMONSTORAGE, False
        End If
    End If
End Sub

Private Sub HandleInitConvo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ConvoNum = buffer.ReadLong
    ConvoData = buffer.ReadByte
    ConvoNpcNum = buffer.ReadLong
    ConvoText = Trim$(buffer.ReadString)
    ConvoNoReply = buffer.ReadByte
    For i = 1 To 3
        ConvoReply(i) = Trim$(buffer.ReadString)
    Next
    Set buffer = Nothing
    
    If ConvoNum > 0 Then
        ConvoRenderText = vbNullString
        ConvoDrawTextLen = 0
        ConvoShowButton = False
        GUI(GuiEnum.GUI_CONVO).Visible = True
    Else
        GUI(GuiEnum.GUI_CONVO).Visible = False
    End If
End Sub

Private Sub HandlePlayerPokemons(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Byte, X As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_PLAYER_POKEMON
        With PlayerPokemons(i)
            .Num = buffer.ReadLong
            
            .Level = buffer.ReadByte
        
            For X = 1 To StatEnum.Stat_Count - 1
                .Stat(X) = buffer.ReadLong
                .StatIV(X) = buffer.ReadLong
                .StatEV(X) = buffer.ReadLong
            Next
            
            '//Vital
            .CurHP = buffer.ReadLong
            .MaxHP = buffer.ReadLong
            
            '//Nature
            .Nature = buffer.ReadByte
            
            '//Shiny
            .IsShiny = buffer.ReadByte
            
            '//Happiness
            .Happiness = buffer.ReadByte
            
            '//Gender
            .Gender = buffer.ReadByte
            
            '//Status
            .Status = buffer.ReadByte
            
            '//Exp
            .CurExp = buffer.ReadLong
            .NextExp = buffer.ReadLong
            
            '//Moveset
            For X = 1 To MAX_MOVESET
                .Moveset(X).Num = buffer.ReadLong
                .Moveset(X).CurPP = buffer.ReadByte
                .Moveset(X).TotalPP = buffer.ReadByte
            Next
            
            '//Ball Used
            .BallUsed = buffer.ReadByte
            
            .HeldItem = buffer.ReadLong
        End With
    Next
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte, X As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadByte
    With PlayerPokemons(Slot)
        .Num = buffer.ReadLong
        
        .Level = buffer.ReadByte
        
        For X = 1 To StatEnum.Stat_Count - 1
            .Stat(X) = buffer.ReadLong
            .StatIV(X) = buffer.ReadLong
            .StatEV(X) = buffer.ReadLong
        Next
        
        '//Vital
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
        
        '//Nature
        .Nature = buffer.ReadByte
        
        '//Shiny
        .IsShiny = buffer.ReadByte
        
        '//Happiness
        .Happiness = buffer.ReadByte
        
        '//Gender
        .Gender = buffer.ReadByte
        
        '//Status
        .Status = buffer.ReadByte
        
        '//Exp
        .CurExp = buffer.ReadLong
        .NextExp = buffer.ReadLong
        
        '//Moveset
        For X = 1 To MAX_MOVESET
            .Moveset(X).Num = buffer.ReadLong
            .Moveset(X).CurPP = buffer.ReadByte
            .Moveset(X).TotalPP = buffer.ReadByte
        Next
        
        '//Ball Used
        .BallUsed = buffer.ReadByte
        
        '//held item
        .HeldItem = buffer.ReadLong
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    CreateActionMsg Trim$(buffer.ReadString), buffer.ReadLong, buffer.ReadLong, buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    If IsPlaying(thePlayer) Then
        If Player(thePlayer).Map = Player(MyIndex).Map Then
            With PlayerPokemon(thePlayer)
                If .Num > 0 Then
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                    
                    .IdleTimer = GetTickCount
                    .IdleAnim = 0
                    .IdleFrameTmr = GetTickCount
                End If
            End With
        End If
    End If
    Set buffer = Nothing
End Sub

Private Sub HandlePlayAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= 255 Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim mappoke As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    mappoke = buffer.ReadLong
    If MapPokemon(mappoke).Num > 0 Then
        If MapPokemon(mappoke).Map = Player(MyIndex).Map Then
            With MapPokemon(mappoke)
                .Attacking = 1
                .AttackTimer = GetTickCount
                    
                .IdleTimer = GetTickCount
                .IdleAnim = 0
                .IdleFrameTmr = GetTickCount
            End With
        End If
    End If
    Set buffer = Nothing
End Sub

Private Sub HandleNewMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MoveLearnPokeSlot = buffer.ReadByte
    MoveLearnNum = buffer.ReadLong
    MoveLearnIndex = buffer.ReadByte
    Set buffer = Nothing
    
    '//Show Move Replace Gui
    GuiState GUI_MOVEREPLACE, True
End Sub

Private Sub HandleGetData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InvUseDataType = buffer.ReadByte
    InvUseSlot = buffer.ReadByte
    Set buffer = Nothing
    
    DragInvSlot = 0
End Sub

Private Sub HandleMapPokemonCatchState(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    If MapPokeNum > 0 Then
        CatchBall(MapPokeNum).X = buffer.ReadLong
        CatchBall(MapPokeNum).Y = buffer.ReadLong
        CatchBall(MapPokeNum).State = buffer.ReadByte
        CatchBall(MapPokeNum).Pic = buffer.ReadByte
        CatchBall(MapPokeNum).InUsed = True
        CatchBall(MapPokeNum).Frame = 0
        CatchBall(MapPokeNum).FrameState = 0
        CatchBall(MapPokeNum).FrameTimer = GetTickCount + 50
        If CatchBall(MapPokeNum).State = 0 Then
            '//Init
            CatchBall(MapPokeNum).Frame = 1
        End If
    End If
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    Player(thePlayer).CurHP = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim sH As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ShopNum = buffer.ReadLong
    Set buffer = Nothing
    If ShopNum > 0 Then
        ShopAddY = 1
        For sH = MAX_SHOP_ITEM To 1 Step -1
            If Shop(ShopNum).ShopItem(sH).Num > 0 Then
                ShopCountItem = sH
                Exit For
            End If
        Next
        If Not GUI(GuiEnum.GUI_SHOP).Visible Then
            GuiState GUI_SHOP, True
        End If
    Else
        ShopNum = 0
        ShopAddY = 0
        ShopCountItem = 0
        If GUI(GuiEnum.GUI_SHOP).Visible Then
            GuiState GUI_SHOP, False
        End If
    End If
End Sub

Private Sub HandleRequest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    PlayerRequest = buffer.ReadLong
    RequestType = buffer.ReadByte
    Set buffer = Nothing
    
    If RequestType > 0 Then
        If PlayerRequest > 0 Then
            If PlayerRequest <> MyIndex Then
                If IsPlaying(PlayerRequest) Then
                    If Player(PlayerRequest).Map = Player(MyIndex).Map Then
                        If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                            If RequestType = 1 Then '//Duel
                                OpenChoiceBox Trim$(Player(PlayerRequest).Name) & TextUIChoiceDuel, CB_REQUEST
                            ElseIf RequestType = 2 Then '//Trade
                                OpenChoiceBox Trim$(Player(PlayerRequest).Name) & TextUIChoiceTrade, CB_REQUEST
                            ElseIf RequestType = 3 Then '//Party
                                OpenChoiceBox Trim$(Player(PlayerRequest).Name) & TextUIChoiceParty, CB_REQUEST
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        '//Close
        If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
            GuiState GUI_CHOICEBOX, False
        End If
    End If
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim SoundName As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    SoundName = Trim$(buffer.ReadString)
    Set buffer = Nothing
    
    '//Play Menu Music
    If SoundName <> "None." Then
        PlayMusic SoundName, True, False
    End If
End Sub

Private Sub HandleOpenTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim SoundName As String

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    TradeIndex = buffer.ReadLong
    Set buffer = Nothing
    
    If RequestType > 0 Then
        If PlayerRequest > 0 Then
            If PlayerRequest <> MyIndex Then
                If IsPlaying(PlayerRequest) Then
                    If Player(PlayerRequest).Map = Player(MyIndex).Map Then
                        If Not GUI(GuiEnum.GUI_TRADE).Visible Then
                            GuiState GUI_TRADE, True
                            TradeInputMoney = "0"
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub HandleUpdateTradeItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim myTrade As Byte
Dim TradeSlot As Byte
Dim TmpTrade As TradeRec
Dim X As Byte

    If TradeIndex <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    myTrade = buffer.ReadByte
    TradeSlot = buffer.ReadByte
    With TmpTrade.data(TradeSlot)
        .TradeType = buffer.ReadByte
        
        .Num = buffer.ReadLong
        .value = buffer.ReadLong
        
        .Level = buffer.ReadByte
        
        For X = 1 To StatEnum.Stat_Count - 1
            .Stat(X) = buffer.ReadLong
            .StatIV(X) = buffer.ReadLong
            .StatEV(X) = buffer.ReadLong
        Next
        
        '//Vital
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
        
        '//Nature
        .Nature = buffer.ReadByte
        
        '//Shiny
        .IsShiny = buffer.ReadByte
        
        '//Happiness
        .Happiness = buffer.ReadByte
        
        '//Gender
        .Gender = buffer.ReadByte
        
        '//Status
        .Status = buffer.ReadByte
        
        '//Exp
        .CurExp = buffer.ReadLong
        .NextExp = buffer.ReadLong
        
        '//Moveset
        For X = 1 To MAX_MOVESET
            .Moveset(X).Num = buffer.ReadLong
            .Moveset(X).CurPP = buffer.ReadByte
            .Moveset(X).TotalPP = buffer.ReadByte
        Next
        
        '//Ball Used
        .BallUsed = buffer.ReadByte
        
        '//Helditem
        .HeldItem = buffer.ReadLong
        
        '//Trade Slot
        .TradeSlot = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    '//Update
    If myTrade = YES Then '//Mine
        YourTrade.data(TradeSlot) = TmpTrade.data(TradeSlot)
    Else
        TheirTrade.data(TradeSlot) = TmpTrade.data(TradeSlot)
    End If
End Sub

Private Sub HandleTradeUpdateMoney(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim valMoney As Long, myTrade As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    myTrade = buffer.ReadByte
    valMoney = buffer.ReadLong
    Set buffer = Nothing
    
    If myTrade = YES Then
        YourTrade.TradeMoney = valMoney
    Else
        TheirTrade.TradeMoney = valMoney
    End If
End Sub

Private Sub HandleSetTradeState(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tState As Byte, myTrade As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    myTrade = buffer.ReadByte
    tState = buffer.ReadByte
    Set buffer = Nothing
    
    If myTrade = YES Then
        YourTrade.TradeSet = tState
        EditInputMoney = False
    Else
        TheirTrade.TradeSet = tState
    End If
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GUI(GuiEnum.GUI_TRADE).Visible Then
        GuiState GUI_TRADE, False
    End If
    If GUI(GuiEnum.GUI_INVENTORY).Visible Then
        GuiState GUI_INVENTORY, False
        Button(ButtonEnum.Game_Bag).State = 0
    End If
    TradeIndex = 0
    CheckingTrade = 0
    TradeInputMoney = "0"
    EditInputMoney = False
    Call ZeroMemory(ByVal VarPtr(YourTrade), LenB(YourTrade))
    Call ZeroMemory(ByVal VarPtr(TheirTrade), LenB(TheirTrade))
End Sub

Private Sub HandlePlayerPokedex(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_POKEMON
        With PlayerPokedex(i)
            .Scanned = buffer.ReadByte
            .Obtained = buffer.ReadByte
        End With
    Next
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokedexSlot(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    i = buffer.ReadLong
    With PlayerPokedex(i)
        .Scanned = buffer.ReadByte
        .Obtained = buffer.ReadByte
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePokemonStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    With MapPokemon(MapPokeNum)
        .Status = buffer.ReadByte
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleMapNpcPokemonStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MapPokeNum = buffer.ReadLong
    With MapNpcPokemon(MapPokeNum)
        .Status = buffer.ReadByte
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim thePlayer As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    thePlayer = buffer.ReadLong
    With PlayerPokemon(thePlayer)
        .Slot = buffer.ReadLong
        .Status = buffer.ReadByte
        .IsConfused = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    If thePlayer = MyIndex Then
        If PlayerPokemon(MyIndex).Slot > 0 Then
            PlayerPokemons(PlayerPokemon(MyIndex).Slot).Status = PlayerPokemon(MyIndex).Status
        End If
    End If
End Sub

Private Sub HandleClearPlayer(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ClearPlayer (buffer.ReadLong)
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonsStat(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte, X As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Slot = buffer.ReadByte
    If Slot > 0 Then
        If PlayerPokemons(Slot).Num > 0 Then
            For X = 1 To StatEnum.Stat_Count - 1
                PlayerPokemons(Slot).Stat(X) = buffer.ReadLong
                PlayerPokemons(Slot).StatIV(X) = buffer.ReadLong
                PlayerPokemons(Slot).StatEV(X) = buffer.ReadLong
            Next
        End If
    End If
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerPokemonStatBuff(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Byte

    If PlayerPokemon(MyIndex).Num <= 0 Then Exit Sub
    If PlayerPokemon(MyIndex).Slot <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For X = 1 To StatEnum.Stat_Count - 1
        PlayerPokemon(MyIndex).StatBuff(X) = buffer.ReadLong
    Next
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerStatus(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).Status = buffer.ReadByte
    Player(MyIndex).IsConfuse = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleWeather(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InitWeather buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleNpcPokemonData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim npcIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    npcIndex = buffer.ReadLong
    With MapNpcPokemon(npcIndex)
        .Init = buffer.ReadByte
        .State = buffer.ReadByte
        '//Ball Location
        .BallX = buffer.ReadLong
        .BallY = buffer.ReadLong
        
        '//General
        .Num = buffer.ReadLong

        '//Location
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Vital
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
        
        '//Shiny
        .IsShiny = buffer.ReadByte
        
        '//Happiness
        .Happiness = buffer.ReadByte
        
        '//Gender
        .Gender = buffer.ReadByte
        
        '//Status
        .Status = buffer.ReadByte
        
        '//Init
        If .Init = YES Then
            If .State = 0 Then
                .Frame = 0
            ElseIf .State = 1 Then
                .Frame = 2
            End If
            .FrameState = 0
            .FrameTimer = GetTickCount + 100
        End If
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleNpcPokemonMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim npcIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    npcIndex = buffer.ReadLong
    With MapNpcPokemon(npcIndex)
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = YES
        
        '//Set offset
        Select Case .Dir
            Case DIR_UP
                .yOffset = TILE_Y
            Case DIR_DOWN
                .yOffset = TILE_Y * -1
            Case DIR_LEFT
                .xOffset = TILE_X
            Case DIR_RIGHT
                .xOffset = TILE_X * -1
        End Select
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleNpcPokemonDir(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim npcIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    npcIndex = buffer.ReadLong
    With MapNpcPokemon(npcIndex)
        .Dir = buffer.ReadByte
        
        '//Clear moving attributes
        .xOffset = 0
        .yOffset = 0
        .Moving = NO
    End With
    Set buffer = Nothing
End Sub

Private Sub HandleNpcPokemonVital(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim npcIndex As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    npcIndex = buffer.ReadLong
    With MapNpcPokemon(npcIndex)
        .CurHP = buffer.ReadLong
        .MaxHP = buffer.ReadLong
    End With
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerNpcDuel(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InNpcDuel = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleReleaseMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MoveRelearnPokeNum = buffer.ReadLong
    MoveRelearnPokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If MoveRelearnPokeNum > 0 And MoveRelearnPokeSlot > 0 Then
        MoveRelearnMaxIndex = 0
        For i = 1 To MAX_POKEMON_MOVESET
            If Pokemon(MoveRelearnPokeNum).Moveset(i).MoveNum > 0 Then
                MoveRelearnMaxIndex = i
            End If
        Next
        If GUI(GuiEnum.GUI_RELEARN).Visible = False Then
            GuiState GUI_RELEARN, True
        End If
    End If
End Sub

Private Sub HandlePlayerAction(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).Action = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).CurExp = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleParty(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    InParty = buffer.ReadByte
    For i = 1 To MAX_PARTY
        PartyName(i) = Trim$(buffer.ReadString)
    Next
    Set buffer = Nothing
End Sub

'//Editor
Private Sub HandleInitMap(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_MAPPER Then Exit Sub
    ChatOn = False
    InitEditor_Map
End Sub

Private Sub HandleInitNpc(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Npc
End Sub

Private Sub HandleNpcs(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Npc(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Npc(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitPokemon(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Pokemon
End Sub

Private Sub HandlePokemons(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Pokemon(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Pokemon(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitItem(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Item
End Sub

Private Sub HandleItems(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Item(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitPokemonMove(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_PokemonMove
End Sub

Private Sub HandlePokemonMoves(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(PokemonMove(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(PokemonMove(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Animation
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Animation(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitSpawn(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Spawn
End Sub

Private Sub HandleSpawn(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Spawn(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Spawn(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitConversation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Conversation
End Sub

Private Sub HandleConversation(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

On Error Resume Next

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Conversation(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Conversation(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Shop
End Sub

Private Sub HandleShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Shop(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Shop(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleInitQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub
    ChatOn = False
    InitEditor_Quest
End Sub

Private Sub HandleQuest(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim n As Long
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    n = buffer.ReadLong
    dSize = LenB(Quest(n))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Quest(n)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
End Sub

Private Sub HandleRank(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    For i = 1 To MAX_RANK
        Rank(i).Name = buffer.ReadString
        Rank(i).Level = buffer.ReadLong
        Rank(i).Exp = buffer.ReadLong
    Next
    Set buffer = Nothing
End Sub

Private Sub HandleDataLimit(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    MAX_PLAYER = buffer.ReadInteger
    Set buffer = Nothing
    
    ReDim Player(1 To MAX_PLAYER) As PlayerRec
    ReDim PlayerPokemon(1 To MAX_PLAYER) As PlayerPokemonRec
End Sub

Private Sub HandlePlayerPvP(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).win = buffer.ReadLong
    Player(MyIndex).Lose = buffer.ReadLong
    Player(MyIndex).Tie = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerCash(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    Player(MyIndex).Cash = buffer.ReadLong
    Player(MyIndex).Money = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleRequestCash(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim value As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    value = buffer.ReadLong
    Set buffer = Nothing
    
    If frmAdmin.optCash Then
        frmAdmin.lblCash = "Player Cash: " & value
    Else
        frmAdmin.lblCash = "Player Money: " & value
    End If
End Sub

Private Sub HandleEventInfo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim value As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    ExpMultiply = buffer.ReadByte
    ExpSecs = buffer.ReadLong
    Set buffer = Nothing
End Sub

Private Sub HandleRequestServerInfo(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    If CurServerList = 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    ServerInfo(CurServerList).Status = buffer.ReadString
    ServerInfo(CurServerList).Player = buffer.ReadInteger
    ServerInfo(CurServerList).Colour = buffer.ReadInteger
    Set buffer = Nothing
    
    '//Close Socket
    DestroyTCP
End Sub

Private Sub HandleClientTime(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    GameWeek = buffer.ReadByte
    GameHour = buffer.ReadByte
    GameMinute = buffer.ReadByte
    GameSecond = buffer.ReadByte
    GameSecond_Velocity = buffer.ReadByte
    
    Set buffer = Nothing
End Sub

Private Sub HandleVirtualShop(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, X As Long, Matriz As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Matriz = buffer.ReadLong
        ReDim VirtualShop(i).Items(1 To Matriz)

        For X = LBound(VirtualShop(i).Items) To UBound(VirtualShop(i).Items)
            VirtualShop(i).Items(X).ItemNum = buffer.ReadLong
            VirtualShop(i).Items(X).ItemQuant = buffer.ReadLong
            VirtualShop(i).Items(X).ItemPrice = buffer.ReadLong
            VirtualShop(i).Items(X).CustomDesc = buffer.ReadByte
        Next X
    Next i
    Set buffer = Nothing

    SwitchTabFromVirtualShop Skins
End Sub

Private Sub HandleFishMode(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long, X As Long, Y As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    i = buffer.ReadLong
    X = buffer.ReadByte
    Y = buffer.ReadByte
    Set buffer = Nothing
    
    Player(i).FishMode = X
    Player(i).FishRod = Y
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, n As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    frmMapReport.lstIndex.Clear
    For i = 1 To MAX_MAP
        MapReport(i) = buffer.ReadString
        frmMapReport.lstIndex.AddItem i & ": " & MapReport(i)
    Next i

    frmMapReport.Show vbModeless, frmMain

    buffer.Flush: Set buffer = Nothing
End Sub

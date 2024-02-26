Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
Dim filename As String

    frmServer.Caption = GAME_NAME & " Server"
    frmServer.Show
    
    LoadOption
    
    Randomize
    
    TextAdd frmServer.txtLog, "Checking directory..."
    '//Let's check if our required Directory Exist
    ChkDir App.Path & "\", "data"
    ChkDir App.Path & "\data\", "logs"
    ChkDir App.Path & "\data\", "accounts"
    ChkDir App.Path & "\data\", "maps"
    ChkDir App.Path & "\data\", "npcs"
    ChkDir App.Path & "\data\", "pokemons"
    ChkDir App.Path & "\data\", "items"
    ChkDir App.Path & "\data\", "moves"
    ChkDir App.Path & "\data\", "animations"
    ChkDir App.Path & "\data\", "mappokemon"
    ChkDir App.Path & "\data\", "conversation"
    ChkDir App.Path & "\data\", "shop"
    ChkDir App.Path & "\data\", "quest"
    ChkDir App.Path & "\data\", "virtualshop"
    
    '//False = Get Error on Log text
    '//True = Get Error from IDE
    'DebugMode = False
    DebugMode = True
    
    filename = App.Path & "\data\data_limit.ini"
    If Not FileExist(filename) Then
        MAX_PLAYER = 100
        Call PutVar(filename, "Data", "Player", Str(MAX_PLAYER))
    Else
        MAX_PLAYER = Val(GetVar(filename, "Data", "Player"))
    End If
    
    Call UsersOnline_Start
    
    ' Data Limit
    ReDim Player(1 To MAX_PLAYER, 1 To MAX_PLAYERCHAR) As PlayerRec
    ReDim PlayerInv(1 To MAX_PLAYER) As PlayerInvRec
    ReDim PlayerPokemons(1 To MAX_PLAYER) As PlayerPokemonsRec
    ReDim PlayerInvStorage(1 To MAX_PLAYER) As PlayerInvStorageRec
    ReDim PlayerPokemonStorage(1 To MAX_PLAYER) As PlayerPokemonStorageRec
    ReDim PlayerPokedex(1 To MAX_PLAYER) As PlayerPokedexRec
    '//Player Pokemon
    ReDim PlayerPokemon(1 To MAX_PLAYER) As PlayerPokemonRec
    ReDim Account(1 To MAX_PLAYER) As AccountRec
    ReDim TempPlayer(1 To MAX_PLAYER) As TempPlayerRec
    
    LoadRank
    
    '//Clear Game Data
    TextAdd frmServer.txtLog, "Clearing Maps..."
    ClearMaps
    TextAdd frmServer.txtLog, "Clearing Npcs..."
    ClearNpcs
    TextAdd frmServer.txtLog, "Clearing Pokemons..."
    ClearPokemons
    TextAdd frmServer.txtLog, "Clearing Items..."
    ClearItems
    TextAdd frmServer.txtLog, "Clearing Moves..."
    ClearPokemonMoves
    TextAdd frmServer.txtLog, "Clearing Animations..."
    ClearAnimations
    TextAdd frmServer.txtLog, "Clearing Spawns..."
    ClearSpawns
    TextAdd frmServer.txtLog, "Clearing Conversations..."
    ClearConversations
    TextAdd frmServer.txtLog, "Clearing Shops..."
    ClearShops
    TextAdd frmServer.txtLog, "Clearing Quests..."
    ClearQuests
    TextAdd frmServer.txtLog, "Clearing Map Npcs..."
    ClearMapNpcs
    TextAdd frmServer.txtLog, "Clearing Virtual Shop..."
    ClearVirtualShop
    TextAdd frmServer.txtLog, "Loading Virtual Shop..."
    LoadVirtualShop
    TextAdd frmServer.txtLog, "Loading Maps..."
    LoadMaps
    TextAdd frmServer.txtLog, "Loading Npcs..."
    LoadNpcs
    TextAdd frmServer.txtLog, "Loading Pokemons..."
    LoadPokemons
    TextAdd frmServer.txtLog, "Loading Items..."
    LoadItems
    TextAdd frmServer.txtLog, "Loading Moves..."
    LoadPokemonMoves
    TextAdd frmServer.txtLog, "Loading Animations..."
    LoadAnimations
    TextAdd frmServer.txtLog, "Loading Spawns..."
    LoadSpawns
    TextAdd frmServer.txtLog, "Loading Conversations..."
    LoadConversations
    TextAdd frmServer.txtLog, "Loading Shops..."
    LoadShops
    TextAdd frmServer.txtLog, "Loading Quests..."
    LoadQuests
    TextAdd frmServer.txtLog, "Creating Map Cache..."
    CacheAllMaps
    TextAdd frmServer.txtLog, "Spawning All Map Npcs..."
    SpawnAllMapNpcs
    TextAdd frmServer.txtLog, "Spawning All Map Pokemons..."
    SpawnAllMapPokemon
    TextAdd frmServer.txtLog, "Add All Pokes Fish By Mapnum..."
    AddPokemonsFishing
    
    frmServer.Caption = GAME_NAME & " Server"
    
    '//Starting up Form Socket
    TcpInit
    InitMessages
    'MainTcpInit
    
    AddLog "Server started up"
    TextAdd frmServer.txtLog, "Initialization Complete..."
    
    '//Obter data e hora do sistema, OBS: O client trabalha com esses horários
    GameHour = hour(Now)
    GameMinute = Minute(Now)
    GameSecs = Second(Now)
    GameSecs_Velocity = 15
    
    AppRunning = True   '//Make sure that our application is actually running
    AppLoop             '//Start the loop
End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYER
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Sub DestroyServer()
    On Error Resume Next
    AddLog "Server closed"
    UpdateSavePlayers
    DestroyTCP
    Unload frmServer
    SaveRank
    End
End Sub

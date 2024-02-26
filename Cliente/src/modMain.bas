Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
Dim i As Long
Dim SetWidth As Long

    If App.PrevInstance = True Then
        MsgBox "Client is already running...", vbCritical
        End
    End If
    
    '// Inicializa a Cryptografia
    InitCryptographyKey
    
    ShowServerList = True
    ServerList = False
    
    ' //SERVER SETUP
    CurServerList = 1
    
    ServerName(1) = "Server 1"
    ServerIP(1) = "127.0.0.1"
    ServerPort(1) = 8090
    
    ServerName(2) = "Server 2"
    ServerIP(2) = "127.0.0.1"
    ServerPort(2) = 8001
    
    ServerName(3) = "Server 3"
    ServerIP(3) = "127.0.0.1"
    ServerPort(3) = 8001

    ' Configurações das Resoluções
    CurResolutionList = CurResolutionList
    ShowResolutionList = True
    ResolutionList = False
    ResolutionName(1) = "800x608"
    ResolutionName(2) = "1280x704"
    ResolutionName(3) = "1344x704"
    ResolutionName(4) = "1600x832"
    ResolutionName(5) = "1856x960"
    ResolutionName(6) = "2432x960"
    
    '//Initialize the random-number generator
    Randomize ', seed
    
    ' Preload the Computer Information
    ProcessorID = GetProcessorID
    
    StartingUp = True
    
    LoadSetting     '//Load All settings required on game
    ClearControlKey
    LoadControlKey  '//Load all key input
    
    '//Let's check if our required Directory Exist
    ChkDir App.path & "\", "data"
    ChkDir App.path & "\data\", "themes"
    ChkDir App.path & "\data\", "music"
    ChkDir App.path & "\data\", "sfx"
    ChkDir App.path & "\data\", "resources"
    ChkDir App.path & "\data\", "cache"
    ChkDir App.path & "\data\cache\", "maps"
    ChkDir App.path & "\data\themes\", Trim$(GameSetting.ThemePath)
    ChkDir App.path & "\data\themes\" & Trim$(GameSetting.ThemePath) & "\", "textures"
    ChkDir App.path & "\data\themes\" & Trim$(GameSetting.ThemePath) & "\", "ui"
    ChkDir App.path & "\data\resources\", "character-sprites"
    ChkDir App.path & "\data\resources\", "player-sprites"
    ChkDir App.path & "\data\resources\", "map-animation"
    ChkDir App.path & "\data\resources\", "world-tiles"
    ChkDir App.path & "\data\resources\", "pokemon"
    ChkDir App.path & "\data\resources\pokemon\", "portrait"
    ChkDir App.path & "\data\resources\", "item"
    ChkDir App.path & "\data\resources\", "misc"
    ChkDir App.path & "\data\resources\", "animation"
    ChkDir App.path & "\data\resources\", "weather"
    
    InitDirectX         '//Load the DirectX
    LoadGui
    
    ' // SERVER SETUP
    ServerMaxWidth = 0
    For i = 1 To MAX_SERVER_LIST
        SetWidth = GetTextWidth(Font_Default, "Server: " & ServerName(i))
        If SetWidth > ServerMaxWidth Then
            ServerMaxWidth = SetWidth
        End If
    Next
    
    Randomize
    
    '//Load Sockets
    TcpInit
    InitMessages
    
    '//Sound
    InitSound
    
    '//Clear Game Data
    ClearGameData
    
    LoadCredits
    
    InitSettingConfiguration
    
    ' Valor random backplayer
    RandBackPlayer = Format(RandomNumBetween(1, 8))
    
    frmMain.Width = Form_Width
    frmMain.Height = Form_Height
    frmMain.Caption = GAME_NAME
    frmMain.Show        '//Make sure that our window is visible
    
    StartingUp = False
    
    '//Set Game State
    CanShowCursor = False
    InitCursorTimer = False
    If GameSetting.SkipBootUp = YES Then
        InitGameState InMenu ', True
        CanShowCursor = True
        InitCursorTimer = True
    Else
        InitGameState InMenu, True
        CanShowCursor = False
        InitCursorTimer = False
    End If
    
    '// Open login window
    User = vbNullString: Pass = vbNullString '//Reset
    GuiState GUI_LOGIN, True
    
    CentralizeWindow frmMain
    
    ' Initialize, using in scrolling of Controls Option.
    ControlMaxViewLine = ControlEnum.Control_Count - 1 - ControlScrollViewLine
    
    ForceExit = False
    AppRunning = True   '//Make sure that our application is actually running
    AppLoop             '//Start the loop
End Sub

Private Sub CentralizeWindow(ByRef Form As Form)
    Form.Left = (Screen.Width / 2) - (Form.Width / 2)
    
    If GameSetting.Width > 800 Then
        Form.top = 0
    Else
        Form.top = (Screen.Height / 2) - (Form.Height / 2)
    End If
End Sub

Sub UnloadMain()
    On Error Resume Next
    
    ForceExit = True
    AppRunning = False
    DestroyDirectX      '//Unloading DirectX
    
    '//Sound
    UnloadSound
    
    '// Clear all Data
    ClearSetting
    
    UnloadAllForms      '//Closing all forms
    End                 '//Terminate the Program
End Sub

'//Close all available forms on the project
Private Sub UnloadAllForms()
Dim frm As Form
    
    On Error Resume Next
    
    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

'//Changing game state and loading/unload the required/not required data
Public Sub InitGameState(ByVal gState As GameStateEnum, Optional ByVal IsStart As Boolean = False)
    GameState = gState
    
    '//Load or Unload data
    Select Case GameState
        Case GameStateEnum.InMenu
            '//If App just opened, show company screen
            If IsStart Then
                MenuState = MenuStateEnum.StateCompanyScreen
                InitFade 2500, FadeIn, 1
            Else
                '//If it just go back to title screen, show menu screen
                MenuState = MenuStateEnum.StateNormal
                
                '//Play Menu Music
                If Trim$(GameSetting.MenuMusic) <> "None." Then
                    If CurMusic <> Trim$(GameSetting.MenuMusic) Then
                        PlayMusic Trim$(GameSetting.MenuMusic), False, True
                    End If
                Else
                    StopMusic True
                End If
            End If
            
            BackgroundXOffset = 640 '//Size of the background texture (Need to change if size changed)
        Case GameStateEnum.InGame
            
    End Select
End Sub

Public Sub UpdateGuiCount(ByVal vGui As GuiEnum, ByVal gVisible As Boolean)
Dim X As Byte

    '//Find gui position
    X = findDataArray(vGui, GuiZOrder)
    
    '//check for empty array
    If (Not GuiZOrder) = False Then
        '//Make sure gui doesn't exist
        If X <= 0 Then Exit Sub
    End If
    
    If gVisible Then
        '//Store it
        GuiVisibleCount = GuiVisibleCount + 1
        ReDim Preserve GuiZOrder(1 To GuiVisibleCount) As Byte
        'UpdateGuiOrder vGui
        GuiZOrder(GuiVisibleCount) = vGui
    Else
        '//Remove it
        If X > 0 Then
            byteArrRemoveData GuiZOrder, X
            GuiVisibleCount = GuiVisibleCount - 1
        End If
    End If
    
End Sub

Public Sub UpdateGuiOrder(ByVal vGui As GuiEnum, Optional ByVal AddGui As Boolean = False)
Dim i As Byte, X As Byte

    '//zOrdering of gui
    If AddGui Then
        GuiZOrder(GuiVisibleCount) = AddGui
    Else
        '//Find gui position
        X = findDataArray(vGui, GuiZOrder)
    
        If X > 0 Then
            For i = X To GuiVisibleCount - 1
                GuiZOrder(i) = GuiZOrder(i + 1)
            Next
            '//Set to top
            GuiZOrder(GuiVisibleCount) = vGui
        End If
    End If
End Sub

Public Sub SetStatus(ByVal lStatus As Boolean, Optional ByVal sText As String = vbNullString)
    IsLoading = lStatus
    LoadText = sText
End Sub

Public Sub Menu_State(ByVal State As Byte)
    SetStatus True, "Connecting to server..."
    
    If ConnectToServer Then
        Select Case State
            Case MENU_STATE_REGISTER
                SetStatus True, "Connected, Sending new account information..."
                SendNewAccount User, Pass, Email
                CheckPing
            Case MENU_STATE_LOGIN
                SetStatus True, "Connected, Sending account information..."
                SendLoginInfo User, Pass
                CheckPing
            Case MENU_STATE_ADDCHAR
                SetStatus True, "Connected, Sending new character information..."
                SendNewCharacter CharName, SelGender, CurChar
            Case MENU_STATE_USECHAR
                SetStatus True, "Connected, Receiving game data. This might take a while.."
                SendUseCharacter CurChar
            Case MENU_STATE_DELCHAR
                SetStatus True, "Connected, Deleting character data..."
                SendDelCharacter CurChar
        End Select
        ShowServerList = True
        ServerList = False
    End If
    
    If IsLoading Then
        If Not IsConnected Then
            SetStatus False
            ShowServerList = True
            ServerList = False
            AddAlert "Sorry, the server seems to be down. Please try to reconnect in a few minutes", White
        End If
    End If
End Sub

Public Sub GuiState(ByVal vGui As GuiEnum, ByVal vState As Boolean, Optional ByVal UpdateOrder As Boolean = False)
    '//Make sure it won't repeat the same state
    If GUI(vGui).Visible = vState Then Exit Sub
    
    GUI(vGui).Visible = vState
    UpdateGuiCount vGui, GUI(vGui).Visible
    
    If vState = False Then
        '//Reset Location
        ResetGuiLocation vGui
    End If
    
    If UpdateOrder Then UpdateGuiOrder vGui
    
    Select Case vGui
        Case GuiEnum.GUI_LOGIN
            CurTextbox = 1
            
            '//Check save password
            If GameSetting.SavePass = YES Then
                User = Trim$(GameSetting.Username)
                Pass = Trim$(GameSetting.Password)
            End If
        Case GuiEnum.GUI_REGISTER
            CurTextbox = 1
            ShowPass = NO
        Case GuiEnum.GUI_POKEDEX
            PokedexScrollY = 132
        Case GuiEnum.GUI_RANK
            RankingScrollY = RankingScrollLength
        Case GuiEnum.GUI_OPTION
            ControlScrollY = 121
    End Select
End Sub

Public Function CanShowButton(ByVal ButtonNum As ButtonEnum) As Boolean
    CanShowButton = True
    
    If GameState = GameStateEnum.InMenu Then
        If Not MenuState = MenuStateEnum.StateNormal Then CanShowButton = False
    End If
    If ButtonNum <> ChoiceBox_No And ButtonNum <> ChoiceBox_Yes Then
        If GettingMap Then CanShowButton = False
        If CreditVisible Then CanShowButton = False
    End If
    If GameState = GameStateEnum.InGame And Editor = EDITOR_MAP Then CanShowButton = False
    
    Select Case ButtonNum
        Case ButtonEnum.Game_Pokedex To ButtonEnum.Game_Menu
            If IsLoading Then CanShowButton = False
            If Fade Then CanShowButton = False
        Case ButtonEnum.Game_Evolve
            If IsLoading Then CanShowButton = False
            If Fade Then CanShowButton = False
            If Not CanPlayerPokemonEvolve Then CanShowButton = False
        Case ButtonEnum.Character_Use, ButtonEnum.Character_Delete
            If Not pCharInUsed(CurChar) Then CanShowButton = False
        Case ButtonEnum.Character_New
            If pCharInUsed(CurChar) Then CanShowButton = False
        Case ButtonEnum.Character_SwitchLeft
            If CurChar = 1 Then CanShowButton = False
        Case ButtonEnum.Character_SwitchRight
            If CurChar = MAX_PLAYERCHAR Then CanShowButton = False
        '//Control Key
        Case ButtonEnum.Option_cTabUp, ButtonEnum.Option_cTabDown
            If setWindow <> ButtonEnum.Option_Control Then CanShowButton = False
        '//Sound Key
        Case ButtonEnum.Option_sMusicUp To ButtonEnum.Option_sSoundDown
            If setWindow <> ButtonEnum.Option_Sound Then CanShowButton = False
        '//Convo
        Case ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
            '//Check if convo have text
            If ConvoNum <= 0 Then CanShowButton = False
            If ConvoData <= 0 Then CanShowButton = False
            If Not ConvoShowButton Then CanShowButton = False
            If Len(Trim$(ConvoReply((ButtonNum + 1) - ButtonEnum.Convo_Reply1))) <= 0 Then CanShowButton = False
        Case ButtonEnum.Trade_Accept, ButtonEnum.Trade_Decline
            If YourTrade.TradeSet = NO Or TheirTrade.TradeSet = NO Then CanShowButton = False
        Case ButtonEnum.Trade_Set
            If YourTrade.TradeSet = YES And TheirTrade.TradeSet = YES Then CanShowButton = False
        Case ButtonEnum.Trade_AddMoney
            If Not IsNumeric(TradeInputMoney) Then CanShowButton = False
            If TradeInputMoney = vbNullString Then
                CanShowButton = False
            Else
                If YourTrade.TradeMoney = Val(TradeInputMoney) Then CanShowButton = False
            End If
        Case Else: '//None
    End Select
End Function

Public Function CanShowGui(ByVal GuiNum As GuiEnum) As Boolean
    CanShowGui = True

    If Not MenuState = MenuStateEnum.StateNormal Then CanShowGui = False
    If IsLoading Then CanShowGui = False
    If GameState = GameStateEnum.InGame And Editor = EDITOR_MAP Then CanShowGui = False
    If GettingMap Then CanShowGui = False
    If CreditVisible Then CanShowGui = False
    If CreditVisible Then GUI(GuiEnum.GUI_GLOBALMENU).Visible = False
    
    Select Case GuiNum
        Case Else: '//None
    End Select
End Function

Public Sub OpenChoiceBox(ByVal cText As String, ByVal ctype As Byte)
    If GameState = GameStateEnum.InMenu Then
        If Not MenuState = MenuStateEnum.StateNormal Then Exit Sub
    End If
    
    GuiState GUI_CHOICEBOX, True
    ChoiceBoxText = cText
    ChoiceBoxType = ctype
End Sub

Public Sub ClearGameData()
    GettingMap = False
    Player_HighIndex = 0
    MyIndex = 0
    TradeIndex = 0
    PlayerRequest = 0
    PlayerRequest = 0
    ShopNum = 0
    ConvoNum = 0
    StorageType = 0
    DragInvSlot = 0
    DragStorageSlot = 0
    DragPokeSlot = 0
    SetAttackMove = 0
    InParty = 0
    '// Evento XP
    ExpMultiply = 0
    ExpSecs = 0
    
    Erase PokemonsStorage_Select

    '//Clear Chat
    ClearChat
    ClearNpcs
    ClearPokemons
    ClearItems
    ClearPokemonMoves
    ClearAnimations
    ClearSpawns
    ClearChatbubble
    ClearConversations
    ClearShops
    ClearQuests
    ClearRank
    ClearVirtualShop
End Sub

Public Sub ResetMenu()
Dim i As Long

    ClearGameData
    '//Close Socket
    DestroyTCP
    '//Set Game State
    InitGameState InMenu
    '//Hide all other window
    For i = GuiEnum.GUI_LOGIN To GuiEnum.Gui_Count - 1
        GuiState i, False
    Next
    '//ToAdd: Hide all in-game window
    IsLoggedIn = False
    '//Remove loading
    SetStatus False
    '// Open login window
    User = vbNullString: Pass = vbNullString '//Reset
    GuiState GUI_LOGIN, True
    
    '//destroy the animations loaded
    For i = 1 To 255
        ClearAnimInstance (i)
    Next
    
    '//Play Menu Music
    If Trim$(GameSetting.MenuMusic) <> "None." Then
        If CurMusic <> Trim$(GameSetting.MenuMusic) Then
            PlayMusic Trim$(GameSetting.MenuMusic), False, True
        End If
    Else
        StopMusic True
    End If
End Sub

Public Sub InitSettingConfiguration()
Dim z As Long

    setDidChange = False
    setWindow = ButtonEnum.Option_Video
    
    isFullscreen = GameSetting.Fullscreen
    WidthSize = GameSetting.Width
    HeightSize = GameSetting.Height
    FPSvisible = GameSetting.ShowFPS
    PingVisible = GameSetting.ShowPing
    tSkipBootUp = GameSetting.SkipBootUp
    Namevisible = GameSetting.ShowName
    PPBarvisible = GameSetting.ShowPP
    GuiPath = Trim$(GameSetting.ThemePath)
    tmpCurLanguage = GameSetting.CurLanguage
    
    BGVolume = CurMusicVolume
    SEVolume = CurSoundVolume
    
    '//Reset
    ControlViewCount = 0
    editKey = 0
    '//Set Key
    For z = 1 To ControlEnum.Control_Count - 1
        TmpKey(z) = ControlKey(z).cAsciiKey
    Next
End Sub

Public Sub SaveSettingConfiguration()
Dim restartToChange As Boolean

    restartToChange = False

    ' FullScreen
    If isFullscreen <> GameSetting.Fullscreen Then
        GameSetting.Fullscreen = isFullscreen
        restartToChange = True
    End If
    
    ' Resolution
    If WidthSize <> GameSetting.Width Then
        GameSetting.Width = WidthSize
        GameSetting.Height = HeightSize
        restartToChange = True
    End If
    
    ' Theme Path
    If Trim$(LCase$(GuiPath)) <> Trim$(LCase$(GameSetting.ThemePath)) Then
        If FileExist(App.path & "\data\themes\" & GuiPath & ".ini") Then
            GameSetting.ThemePath = Trim$(GuiPath)
            restartToChange = True
        Else
            AddAlert "Failed to load gui path", White
        End If
    End If
    
    ' Music Configuration
    ChangeVolume BGVolume, True
    ChangeVolume SEVolume, False
    GameSetting.Background = BGVolume
    GameSetting.SoundEffect = SEVolume
    
    '//Fps
    GameSetting.ShowFPS = FPSvisible
    
    '//Ping
    GameSetting.ShowPing = PingVisible
    
    '//Skip BootUp
    GameSetting.SkipBootUp = tSkipBootUp
    
    '//Name
    GameSetting.ShowName = Namevisible
    
    '//PP Bar
    GameSetting.ShowPP = PPBarvisible
    
    '//Language
    If tmpCurLanguage <> GameSetting.CurLanguage Then
        GameSetting.CurLanguage = tmpCurLanguage
        SendSetLanguage
    End If
    
    If restartToChange Then
        AddAlert "You must restart your client to apply some changed settings", White
    End If
    
    SaveSetting
End Sub

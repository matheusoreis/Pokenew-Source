Attribute VB_Name = "modGlobal"
Option Explicit

'//Server List
Public ServerMaxWidth As Long
Public ShowServerList As Boolean
Public ServerList As Boolean
Public CurServerList As Integer
Public ServerName(1 To MAX_SERVER_LIST) As String
Public ServerIP(1 To MAX_SERVER_LIST) As String
Public ServerPort(1 To MAX_SERVER_LIST) As Long

Public ServerInfo(1 To MAX_SERVER_LIST) As ServerInfoRec

Public MAX_PLAYER As Integer

Public PokedexShowInfo As Boolean
Public PokedexInfoIndex As Integer
Public PokedexShowTimer As Long

'//General
Public StartingUp As Boolean          '//This check whether the app is just starting up
Public AppRunning As Boolean          '//Controls whether the program is running or not...
Public GameState As Byte              '//Controls the current state of the game (In-Game, In-Menu, In-Loading)
Public MenuState As Byte              '//Controls the current state of the menu
Public Connected As Boolean
Public IsLoggedIn As Boolean
Public ForceExit As Boolean
Public GettingMap As Boolean
Public ReInit As Boolean

Public ProcessorID As String

'//Loading
Public IsLoading As Boolean           '//Controls if the game is currently loading or not
Public LoadText As String             '//This is the text that is being draw on loading screen

'//Mouse Icon Pointer
Public IsHovering As Boolean
Public MouseIcon As Byte

'//Background Offset
Public BackgroundXOffset As Long      '//This control the movement of the menu screen background

'//For Faders
Public Fade As Boolean                '//Check if we can fade
Public FadeWait As Long               '//Check how long before fade start/end
Public FadeState As Byte              '//Check if Fade In/Out
Public FadeAlpha As Byte              '//Control the alpha of fade screen
Public FadeType As Byte               '//Control the event process after the fade

'//Cursor
Public CanShowCursor As Boolean
Public InitCursorTimer As Boolean
Public CursorTimer As Long
Public oldCursorX As Long
Public oldCursorY As Long
Public CursorX As Long
Public CursorY As Long
Public curTileX As Long
Public curTileY As Long
Public CursorLoadAnim As Byte

'//Configuration of Screen Size base of Resolution
Public Form_Width As Long
Public Form_Height As Long
Public Screen_Width As Long
Public Screen_Height As Long

'//GUI
Public GuiVisibleCount As Long        '//This count all Gui that are visible
Public GuiZOrder() As Byte            '//Store the number of gui visible
Public ShortKeyTimer As Long

'//Menu Textbox
Public TextLine As String * 1
Public User As String
Public Pass As String
Public Pass2 As String
Public Email As String
Public ShowPass As Byte
Public CurTextbox As Byte

'//Index
Public MyIndex As Long                '//Socket Index of Client
Public Player_HighIndex As Long
Public Npc_HighIndex As Long
Public Pokemon_HighIndex As Long
Public ActionMsgIndex As Byte
Public Action_HighIndex As Byte
Public AnimationIndex As Byte

'//Ping
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long
Public PingToDraw As String

'//Character Selection
Public CurChar As Byte
Public pCharInUsed(1 To MAX_PLAYERCHAR) As Boolean
Public pCharName(1 To MAX_PLAYERCHAR) As String
Public pCharSprite(1 To MAX_PLAYERCHAR) As Long

'//Character Creation
Public SelGender As Byte
Public GenderAnim As Byte
Public CharName As String

'//ChoiceBox
Public ChoiceBoxText As String
Public ChoiceBoxType As Byte

'//InputBox
Public InputBoxHeader As String
Public InputBoxText As String
Public InputBoxType As Byte
Public InputBoxLen As Long
Public InputBoxData1 As Long
Public InputBoxData2 As Long

'//Password Changing
Public NewPassword As String
Public OldPassword As String

'//Camera
Public Camera As RECT              '//update the tile/characters/objects/etc position when player is moving
Public TileView As RECT            '//determine the range of the game screen
Public ViewPortInit As Boolean

'//View Port
Public ScreenX As Long
Public ScreenY As Long
Public StartXValue As Long
Public StartYValue As Long
Public EndXValue As Long
Public EndYValue As Long
Public GlobalMapX As Long
Public GlobalMapY As Long

'//FPS
Public GameFps As Long

'//Player Moving
Public CanMoveNow As Boolean

'//Settings
Public setDidChange As Boolean
Public setWindow As Byte

'//Settings Configuration
Public isFullscreen As Byte
Public WidthSize As Long
Public HeightSize As Long
Public FPSvisible As Byte
Public PingVisible As Byte
Public tSkipBootUp As Byte
Public Namevisible As Byte
Public PPBarvisible As Byte
Public GuiPath As String
Public GuiPathEdit As Boolean
Public BGVolume As Byte
Public SEVolume As Byte
Public tmpCurLanguage As Byte

'//Chatbox
Public ChatOn As Boolean
Public EditTab As Boolean
Public MyChat As String
Public ChatTab As String
Public ChatMinimize As Boolean

'//Menu
Public WaitTimer As Long

'//Shiny-Summary
Public ShinySummaryStep As Byte
Public ShinySummaryTmr As Long

'//Map
Public MapAnim As Byte
Public MapFrameAnim As Long

'//Info
Public ShowLoc As Boolean

'//chat bubble
Public chatBubble(1 To 255) As ChatBubbleRec
Public chatBubbleIndex As Long

'//Player Pokemon
Public SelPoke As Byte

'//Animation
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

'//Game Time
Public GameHour As Byte
Public GameMinute As Byte
Public GameSecond As Byte
Public GameSecond_Velocity As Byte
Public GameWeek As Byte

'//New Moves
Public MoveLearnPokeSlot As Byte
Public MoveLearnNum As Long
Public MoveLearnIndex As Byte

'//Evolve
Public EvolveSelect As Byte

'//Credit
Public CreditVisible As Boolean
Public CreditState As Byte
Public CreditOffset As Long
Public CreditTextCount As Long

'//Day And Night
Public DayAndNightARGB As Long
Public ShowLights As Boolean
Public LightAlpha As Byte
Public WeatherAlpha As Long
Public WeatherAlphaState As Byte

'//Move
Public SetAttackMove As Byte

'//Drag
Public DragInvSlot As Byte
Public DragStorageSlot As Byte
Public DragPokeSlot As Byte
Public InvUseSlot As Byte
Public InvUseDataType As Byte

'//Storage
Public StorageType As Byte
Public InvCurSlot As Byte
Public PokemonCurSlot As Byte

'//Window
Public WindowPriority As Byte

'//Convo
Public ConvoNum As Long
Public ConvoData As Byte
Public ConvoNpcNum As Long
Public ConvoText As String
Public ConvoShowButton As Boolean
Public ConvoRenderText As String
Public ConvoDrawTextLen As Long
Public ConvoNoReply As Byte
Public ConvoReply(1 To 3) As String

'//Shop
Public ShopNum As Long
Public ShopAddY As Byte
Public ShopCountItem As Byte
'//ForShopButton
Public ShopButtonHover As Byte
Public ShopButtonState As Byte

'//Duel Request
Public PlayerRequest As Long
Public RequestType As Byte

'//Trade
Public TradeIndex As Long
Public TheirTrade As TradeRec
Public YourTrade As TradeRec
Public CheckingTrade As Byte
Public TradeInputMoney As String
Public EditInputMoney As Boolean

'//Pokedex Scrolling
Public PokedexScrollHold As Boolean
Public PokedexScroll As Long
Public PokedexScrollY As Long
Public PokedexScrollCount As Long
Public PokedexViewCount As Long
Public MaxPokedexViewLine As Long
Public PokedexScrollUp As Boolean
Public PokedexScrollDown As Boolean
Public PokedexScrollTimer As Long
Public PokedexHighIndex As Long

'//Ranking Scrolling
Public RankingScrollHold As Boolean
Public RankingScroll As Long
Public RankingScrollY As Long
Public RankingScrollCount As Long
Public RankingViewCount As Long
Public RankingMaxViewLine As Long
Public RankingScrollUp As Boolean
Public RankingScrollDown As Boolean
Public RankingScrollTimer As Long
Public RankingHighIndex As Long

' Controle Scrooling
Public ControlScrollHold As Boolean
Public ControlScroll As Long
Public ControlScrollY As Long
Public ControlViewCount As Long
Public ControlMaxViewLine As Long
Public ControlScrollUp As Boolean
Public ControlScrollDown As Boolean
Public ControlScrollTimer As Long
'//ControlKey
Public editKey As Long
Public TmpKey(1 To ControlEnum.Control_Count - 1) As Long

'//Releasing thing
Public ReleaseStorageSlot As Byte
Public ReleaseStorageData As Byte

'//Fly
Public FlyBadgeSlot As Byte

'//Spawn Timer
Public SpawnTimer As Long

'//Summary
Public SummaryType As Byte
Public SummarySlot As Long
Public SummaryData As Long

'//Buying Slot
Public BuySlotType As Byte
Public BuySlotData As Byte

'//Duel
Public InNpcDuel As Long

'//Move Relearn
Public MoveRelearnPokeNum As Long
Public MoveRelearnPokeSlot As Byte
Public MoveRelearnCurPos As Byte
Public MoveRelearnMaxIndex As Byte

'//Party
Public InParty As Byte
Public PartyName(1 To MAX_PARTY) As String

'//Rank
Public RankScroll As Byte

'// Evento XP
Public ExpMultiply As Byte
Public ExpSecs As Long

'// Inv Desc
Public InvItemDescTimer As Long
Public InvItemDesc As Integer
Public InvItemDescShow As Boolean

'// Shop Desc
Public ShopItemDescTimer As Long
Public ShopItemDesc As Integer
Public ShopItemDescShow As Boolean

'// Storage Item Desc
Public StorageItemDescTimer As Long
Public StorageItemDesc As Integer
Public StorageItemDescShow As Boolean

'// Trade Item Desc
Public TradeItemDescTimer As Long
Public TradeItemDesc As Integer
Public TradeItemDescShow As Boolean
Public TradeItemDescType As Byte
Public TradeItemDescOwner As Byte '1= Me(Their) 2= You(Your)

' Variáveis globais da lista de resolução
Public ShowResolutionList As Boolean
Public ResolutionList As Boolean
Public CurResolutionList As Integer
Public ResolutionName(1 To MAX_RESOLUTION_LIST) As String

' Fundo do jogador aleatório
Public RandBackPlayer As Integer

' Texto da janela de login
Public TextUILoginUsername As String
Public TextUILoginPassword As String
Public TextUILoginServerList As String
Public TextUILoginCheckBox As String
Public TextUILoginEntryButton As String
Public TextUILoginInvalidUsername As String
Public TextUILoginInvalidPassword As String

' Texto da janela de registro
Public TextUIRegisterUsername As String
Public TextUIRegisterPassword As String
Public TextUIRegisterEmail As String
Public TextUIRegisterConfirm As String
Public TextUIRegisterCheckBox As String
Public TextUIRegisterUsernameLenght As String
Public TextUIRegisterPasswordLenght As String
Public TextUIRegisterPasswordMatch As String
Public TExtUIRegisterInvalidEmail As String

' Texto da janela de criação de porsonagem
Public TextUICreateCharacterCreateButton As String
Public TextUICreateCharacterUsername As String
Public TextUICreateCharacterUsernameLenght As String

' Texto Universal
Public TextUIWait As String

' Texto Footer
Public TextUIFooterCreateAccount As String
Public TextUIFooterCredits As String
Public TextUIFooterDeveloper As String
Public TextUIFooterChangePassword As String

' Texto Global menu
Public TextUIGlobalMenuReturn As String
Public TextUIGlobalMenuOptions As String
Public TextUIGlobalMenuReturnMenu As String
Public TextUIGlobalMenuExit As String

' Texto Choice Window
Public TextUIChoiceYes As String
Public TextUIChoiceNo As String
Public TextUIChoiceExit As String

' Texto Input Window
Public TextUIInputCancel As String
Public TextUIInputConfirm As String
Public TextUIInputAmountHeader As String
Public TextUIInputNewPasswordHeader As String

Public TextUIChoiceReturnMainMenu As String
Public TextUIChoiceBuyInvSlot As String
Public TextUIChoiceEvolve As String
Public TextUIChoiceRelease As String
Public TextUIChoiceBuySlot As String
Public TextUIChoiceFly As String
Public TextUIChoiceDuel As String
Public TextUIChoiceTrade As String
Public TextUIChoiceParty As String
Public TextUIChoiceSave As String
Public TextUIChoiceDeleteCharacter As String

' Texto Opções Menu
Public TextUIOptionVideoButton As String
Public TextUIOptionSoundButton As String
Public TextUIOptionGameButton As String
Public TextUIOptionControlButton As String
Public TextUIOptionFullscreen As String
Public TextUIOptionPath As String
Public TextUIOptionsFps As String
Public TextUIOptionsPing As String
Public TextUIOptionsFast As String
Public TextUIOptionName As String
Public TextUIOptionLanguage As String
Public TextUIOptionMusic As String
Public TextUIOptionSound As String
Public TextUIOptionPP As String
' Controles
Public TextUIOptionUp As String
Public TextUIOptionDown As String
Public TextUIOptionLeft As String
Public TextUIOptionRight As String
Public TextUIOptionCheckMove As String
Public TextUIOptionMoveSlot1 As String
Public TextUIOptionMoveSlot2 As String
Public TextUIOptionMoveSlot3 As String
Public TextUIOptionMoveSlot4 As String
Public TextUIOptionAttack As String
Public TextUIOptionPokeSlot1 As String
Public TextUIOptionPokeSlot2 As String
Public TextUIOptionPokeSlot3 As String
Public TextUIOptionPokeSlot4 As String
Public TextUIOptionPokeSlot5 As String
Public TextUIOptionPokeSlot6 As String
Public TextUIOptionHotbarSlot1 As String
Public TextUIOptionHotbarSlot2 As String
Public TextUIOptionHotbarSlot3 As String
Public TextUIOptionHotbarSlot4 As String
Public TextUIOptionHotbarSlot5 As String
Public TextUIOptionInventory As String
Public TextUIOptionPokedex As String
Public TextUIOptionInteract As String
Public TextUIOptionConvoChoice1 As String
Public TextUIOptionConvoChoice2 As String
Public TextUIOptionConvoChoice3 As String
Public TextUIOptionConvoChoice4 As String

' Seleção de personagem
Public TextUICharactersNew As String
Public TextUICharactersNone As String
Public TextUICharactersUse As String
Public TextUICharactersDelete As String

' Chat
Public TextEnterToChat As String

' Editor Spawn, Set Map XY on Click Map
Public SpawnSet As Boolean

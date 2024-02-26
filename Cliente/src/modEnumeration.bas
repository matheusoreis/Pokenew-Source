Attribute VB_Name = "modEnumeration"
Option Explicit

'//The order of the packets must match with the client's packet enumeration

'//Packets sent by server to client
Public Enum ServerPackets
    SSendPing = 1
    SHighIndex
    SAlertMsg
    SLoginOk
    SCharacters
    SInGame
    SPlayerData
    SMap
    SCheckForMap
    SMapDone
    SPlayerMove
    SPlayerXY
    SPlayerDir
    SLeftGame
    SPlayerMsg
    SSpawnMapNpc
    SMapNpcData
    SNpcMove
    SNpcDir
    SPokemonData
    SPokemonHighIndex
    SPokemonMove
    SPokemonDir
    SPokemonVital
    SChatbubble
    SPlayerPokemonData
    SPlayerPokemonMove
    SPlayerPokemonXY
    SPlayerPokemonDir
    SPlayerPokemonVital
    SPlayerPokemonPP
    SPlayerInv
    SPlayerInvSlot
    SPlayerPokemons
    SPlayerPokemonSlot
    SActionMsg
    SAttack
    SPlayAnimation
    SNpcAttack
    SNewMove
    SGetData
    SMapPokemonCatchState
    SPlayerVital
    SPlayerInvStorage
    SPlayerInvStorageSlot
    SPlayerPokemonStorage
    SPlayerPokemonStorageSlot
    SStorage
    SInitConvo
    SOpenShop
    SRequest
    SPlaySound
    SOpenTrade
    SUpdateTradeItem
    STradeUpdateMoney
    SSetTradeState
    SCloseTrade
    SPlayerPokedex
    SPlayerPokedexSlot
    SPokemonStatus
    SMapNpcPokemonStatus
    SPlayerPokemonStatus
    SClearPlayer
    SPlayerPokemonsStat
    SPlayerPokemonStatBuff
    SPlayerStatus
    SWeather
    SNpcPokemonData
    SNpcPokemonMove
    SNpcPokemonDir
    SNpcPokemonVital
    SPlayerNpcDuel
    SRelearnMove
    SPlayerAction
    SPlayerExp
    SParty
    '//Editors
    SInitMap
    SInitNpc
    SNpcs
    SInitPokemon
    SPokemons
    SInitItem
    SItems
    SInitPokemonMove
    SPokemonMoves
    SInitAnimation
    SAnimation
    SInitSpawn
    SSpawn
    SInitConversation
    SConversation
    SInitShop
    SShop
    SInitQuest
    SQuest
    SRank
    SDataLimit
    SPlayerPvP
    SPlayerCash
    SRequestCash
    SEventInfo
    SRequestServerInfo
    SClientTime
    SSendVirtualShop
    SFishMode
    SMapReport
    '//Make sure SMSG_COUNT is below everything else
    SMSG_Count
End Enum

'//Packets sent by client to server
Public Enum ClientPackets
    CCheckPing = 1
    CNewAccount
    CLoginInfo
    CNewCharacter
    CUseCharacter
    CDelCharacter
    CNeedMap
    CPlayerMove
    CPlayerDir
    CMapMsg
    CGlobalMsg
    CPartyMsg
    CPlayerMsg
    CWarpTo
    CAdminWarp
    CWarpToMe
    CWarpMeTo
    CSetAccess
    CPlayerPokemonMove
    CPlayerPokemonDir
    CGetItem
    CPlayerPokemonState
    CAttack
    CChangePassword
    CReplaceNewMove
    CEvolvePoke
    CUseItem
    CSwitchInvSlot
    CGotData
    COpenStorage
    CDepositItemTo
    CSwitchStorageSlot
    CSwitchStorageItem
    CWithdrawItemTo
    CConvo
    CProcessConvo
    CDepositPokemon
    CWithdrawPokemon
    CSwitchStoragePokeSlot
    CSwitchStoragePoke
    CCloseShop
    CBuyItem
    CSellItem
    CRequest
    CRequestState
    CAddTrade
    CRemoveTrade
    CTradeUpdateMoney
    CSetTradeState
    CTradeState
    CScanPokedex
    CMOTD
    CCopyMap
    CReleasePokemon
    CGiveItemTo
    CGivePokemonTo
    CSpawnPokemon
    CSetLanguage
    CBuyStorageSlot
    CSellPokeStorageSlot
    CChangeShinyRate
    CRelearnMove
    CUseRevive
    CAddHeld
    CRemoveHeld
    CStealthMode
    CWhosOnline

    CRequestRank

    CHotbarUpdate
    CUseHotbar
    CCreateParty
    CLeaveParty
    '//Editors
    CRequestEditMap
    CMap
    CRequestEditNpc
    CRequestNpc
    CSaveNpc
    CRequestEditPokemon
    CRequestPokemon
    CSavePokemon
    CRequestEditItem
    CRequestItem
    CSaveItem
    CRequestEditPokemonMove
    CRequestPokemonMove
    CSavePokemonMove
    CRequestEditAnimation
    CRequestAnimation
    CSaveAnimation
    CRequestEditSpawn
    CRequestSpawn
    CSaveSpawn
    CRequestEditConversation
    CRequestConversation
    CSaveConversation
    CRequestEditShop
    CRequestShop
    CSaveShop
    CRequestEditQuest
    CRequestQuest
    CSaveQuest
    CKickPlayer
    CBanPlayer
    CMutePlayer
    CUnmutePlayer
    CFlyToBadge
    CRequestCash
    CSetCash
    CRequestServerInfo
    CBuyInvSlot
    CRequestVirtualShop
    CPurchaseVirtualShop
    CMapReport
    '//Make sure CMSG_COUNT is below everything else
    CMSG_Count
End Enum

Public HandleDataSub(SMSG_Count) As Long

'//Map Layers
Public Enum MapLayer
    Ground = 0
    mask
    Mask2
    Fringe
    Fringe2
    Lights
    '//Make sure MapLayer_Count is below everything else
    MapLayer_Count
End Enum

'//Map Layer Type
Public Enum MapLayerType
    Normal = 0
    Animated
End Enum

'//Map Attributes
Public Enum MapAttribute
    Walkable = 0
    Blocked
    NpcSpawn
    NpcAvoid
    Warp
    HealPokemon
    BothStorage
    InvStorage
    PokemonStorage
    ConvoTile
    Slide
    Checkpoint
    WarpCheckpoint
    FishSpot
    '//Make sure MapAttribute_Count is below everything else
    MapAttribute_Count
End Enum

'//Stats
Public Enum StatEnum
    HP = 1
    Atk
    Def
    SpAtk
    SpDef
    Spd
    '//Make sure Stat_Count is below everything else
    Stat_Count
End Enum

'//Types of Database
Public Enum PokemonType
    typeNone = 0
    typeNormal
    typeFire
    typeWater
    typeElectric
    typeGrass
    typeIce
    typeFighting
    typePoison
    typeGround
    typeFlying
    typePsychic
    typeBug
    typeRock
    typeGhost
    typeDragon
    typeDark
    typeSteel
    typeFairy
    '//Make sure PokemonType_Count is below everything else
    PokemonType_Count
End Enum

'//Nature
Public Enum PokemonNature
    None = -1
    '//Neutral
    NatureHardy = 0
    NatureDocile
    NatureSerious
    NatureBashful
    NatureQuirky
    '//Others
    NatureLonely
    NatureBrave
    NatureAdamant
    NatureNaughty
    NatureBold
    NatureRelaxed
    NatureImpish
    NatureLax
    NatureTimid
    NatureHasty
    NatureJolly
    NatureNaive
    NatureModest
    NatureMild
    NatureQuiet
    NatureRash
    NatureCalm
    NatureGentle
    NatureSassy
    NatureCareful
    '//Make sure PokemonNature_Count is below everything else
    PokemonNature_Count
End Enum

'//Category
Public Enum MoveCategory
    Neutral = 0
    Physical
    Special
    Status
End Enum

'//EggGroup
Public Enum EggGroupEnum
    Amorphous = 0
    Bug
    Dragon
    Fairy
    Field
    Flying
    Grass
    HumanLike
    Mineral
    Monster
    Water1
    Water2
    Water3
    Ditto
    Undiscovered
End Enum

'//Growth Rate
Public Enum GrowthRateEnum
    Erratic = 0
    Fast
    MediumFast
    MediumSlow
    Slow
    Fluctuating
End Enum

'//Game Weather
Public Enum WeatherEnum
    None = 0
    Sunny
    Rain
    Snow
    SandStorm
    Hail
    '//weather count
    Count_Weather
End Enum

'//Item Type
Public Enum ItemTypeEnum
    None = 0
    PokeBall
    Medicine
    Berries
    keyItems
    TM_HM
    PowerBracer
    Items
    MysteryBox
End Enum

'//Ball Type
Public Enum BallEnum
    b_Pokeball = 0
    b_Greatball
    b_Ultraball
    b_Masterball
    b_Primerball
    b_CherishBall
    b_LuxuryBall
    b_FriendBall
    b_NetBall
    b_DiveBall
    b_RepeatBall
    b_TimerBall
    b_SafariBall
    b_QuickBall
    b_DuskBall
    b_LoveBall
    
    BallEnum_Count
End Enum

'//Status
Public Enum StatusEnum
    None = 0
    Poison
    Paralize
    Sleep
    Frozen
    Burn
End Enum

' *************
' ** General **
' *************
'//Game State
Public Enum GameStateEnum
    InMenu = 1
    InGame
End Enum

'//Menu State
Public Enum MenuStateEnum
    StateCompanyScreen = 1
    StateTitleScreen
    StateNormal
End Enum

'//Fade
Public Enum FadeStateEnum
    FadeOut = 0
    FadeIn
End Enum

' **************
' ** Graphics **
' **************
'//Gui
Public Enum GuiEnum
    GUI_LOGIN = 1
    GUI_REGISTER
    GUI_CHARACTERSELECT
    GUI_CHARACTERCREATE
    GUI_CHOICEBOX
    GUI_GLOBALMENU
    GUI_OPTION
    GUI_CHATBOX
    GUI_INVENTORY
    GUI_INPUTBOX
    GUI_MOVEREPLACE
    GUI_TRAINER
    GUI_INVSTORAGE
    GUI_POKEMONSTORAGE
    GUI_CONVO
    GUI_SHOP
    GUI_TRADE
    GUI_POKEDEX
    GUI_POKEMONSUMMARY
    GUI_RELEARN
    GUI_BADGE
    GUI_RANK
    GUI_VIRTUALSHOP
    '//Make sure that Gui_Count is below everything else
    Gui_Count
End Enum

'//Buttons
Public Enum ButtonEnum
    Login_Confirm = 1
    Register_Confirm
    Register_Close
    Character_SwitchLeft
    Character_SwitchRight
    Character_New
    Character_Use
    Character_Delete
    CharCreate_Confirm
    CharCreate_Close
    ChoiceBox_Yes
    ChoiceBox_No
    GlobalMenu_Return
    GlobalMenu_Option
    GlobalMenu_Back
    GlobalMenu_Exit
    Option_Close
    Option_Video
    Option_Sound
    Option_Game
    Option_Control
    Option_cTabUp
    Option_cTabDown
    Option_sMusicUp
    Option_sMusicDown
    Option_sSoundUp
    Option_sSoundDown
    Chatbox_ScrollUp
    Chatbox_ScrollDown
    Chatbox_Minimize
    Game_Pokedex
    Game_Bag
    Game_Card
    Game_CheckIn
    Game_Rank
    Game_VirtualShop
    Game_Menu
    Game_Evolve
    Inventory_Close
    InputBox_Okay
    InputBox_Cancel
    MoveReplace_Slot1
    MoveReplace_Slot2
    MoveReplace_Slot3
    MoveReplace_Slot4
    MoveReplace_Cancel
    Trainer_Close
    Trainer_Badge
    InvStorage_Close
    InvStorage_Slot1
    InvStorage_Slot2
    InvStorage_Slot3
    InvStorage_Slot4
    InvStorage_Slot5
    PokemonStorage_Close
    PokemonStorage_Slot1
    PokemonStorage_Slot2
    PokemonStorage_Slot3
    PokemonStorage_Slot4
    PokemonStorage_Slot5
    Convo_Reply1
    Convo_Reply2
    Convo_Reply3
    Shop_Close
    Shop_ScrollUp
    Shop_ScrollDown
    Trade_Close
    Trade_Accept
    Trade_Decline
    Trade_Set
    Trade_AddMoney
    Pokedex_Close
    Pokedex_ScrollUp
    Pokedex_ScrollDown
    PokemonSummary_Close
    Relearn_Close
    Relearn_ScrollDown
    Relearn_ScrollUp
    Badge_Close
    Rank_Close
    Rank_ScrollUp
    Rank_ScrollDown
    VirtualShop_Close
    VirtualShop_Buy
    VirtualShop_ScrollDown
    VirtualShop_ScrollUp
    '//Make sure that Gui_Count is below everything else
    Button_Count
End Enum

'//Button State
Public Enum ButtonState
    StateNormal = 0
    StateHover
    StateClick
End Enum

'//System Texture
Public Enum gSystemEnum
    UserInterface = 1
    CursorIcon
    CursorLoad
End Enum

'//Surface Texture
Public Enum gSurfaceEnum
    CompanyScreen = 1
    TitleScreen
    Background
End Enum

'//Control KeyEnum
Public Enum ControlEnum
    KeyUp = 1
    KeyDown
    KeyLeft
    KeyRight
    KeyCheckMove
    KeyMoveUp
    KeyMoveDown
    KeyMoveLeft
    KeyMoveRight
    KeyAttack
    KeyPokeSlot1
    KeyPokeSlot2
    KeyPokeSlot3
    KeyPokeSlot4
    KeyPokeSlot5
    KeyPokeSlot6
    KeyHotbarSlot1
    KeyHotbarSlot2
    KeyHotbarSlot3
    KeyHotbarSlot4
    KeyHotbarSlot5
    KeyInventory
    KeyPokedex
    KeyInteract
    KeyConvo1
    KeyConvo2
    KeyConvo3
    KeyConvo4
    '//Make sure that Control_Count is below everything else
    Control_Count
End Enum

'//SelMenu
Public Enum SelMenuType
    Inv = 1
    SpawnPokes
    RevivePokes
    PlayerPokes
    Evolve
    Storage
    NPCChat
    InvStorage
    PokeStorage
    PlayerMenu
    TradeItem
    PokedexMapPokemon
    PokedexPlayerPokemon
    ConvoTileCheck
End Enum

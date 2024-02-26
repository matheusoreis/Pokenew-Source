Attribute VB_Name = "modType"
Option Explicit

' *************
' ** General **
' *************
Public Player() As PlayerRec
Public PlayerInv(1 To MAX_PLAYER_INV) As PlayerInvRec
Public PlayerPokemons(1 To MAX_PLAYER_POKEMON) As PlayerPokemonsRec
Public PlayerInvStorage(1 To MAX_STORAGE_SLOT) As PlayerInvStorageSlotRec
Public PlayerPokemonStorage(1 To MAX_STORAGE_SLOT) As PlayerPokemonStorageSlotRec
Public PlayerPokedex(1 To MAX_POKEMON) As PlayerPokedexRec
'//Player Pokemon
Public PlayerPokemon() As PlayerPokemonRec
Public Map As MapRec
Public Npc(1 To MAX_NPC) As NpcRec
Public MapNpcPokemon(1 To MAX_MAP_NPC) As MapNpcPokemonRec
Public MapNpc(1 To MAX_MAP_NPC) As MapNpcRec
Public Pokemon(1 To MAX_POKEMON) As PokemonRec
Public MapPokemon(1 To MAX_GAME_POKEMON) As MapPokemonRec
'Public Item(1 To MAX_ITEM) As ItemRec
Public PokemonMove(1 To MAX_POKEMON_MOVE) As PokemonMoveRec
Public Animation(1 To MAX_ANIMATION) As AnimationRec
Public Spawn(1 To MAX_GAME_POKEMON) As SpawnRec
Public Conversation(1 To MAX_CONVERSATION) As ConversationRec
Public Shop(1 To MAX_SHOP) As ShopRec
Public Quest(1 To MAX_QUEST) As QuestRec

Public Rank(1 To MAX_RANK) As RankRec

Public MapReport(1 To MAX_MAP) As String


Private Type RankRec
    Name As String
    Level As Long
    Exp As Long
End Type

'//Server info
Public Type ServerInfoRec
    Player As Integer
    Status As String
    Colour As Integer
End Type

' **************
' ** Map Data **
' **************
Private Type MapPokemonRec
    '//General
    Num As Long
    
    '//Location
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    
    '//Vital
    CurHP As Long
    MaxHP As Long
    
    '//Shiny
    IsShiny As Byte
    
    '//Happiness
    Happiness As Byte
    
    '//Gender
    Gender As Byte
    
    '//Status
    Status As Byte

    '//Client Only
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
    Attacking As Byte
    AttackTimer As Long
    IdleTimer As Long
    IdleFrameTmr As Long
    IdleAnim As Byte
    MoveSpeed As Long
End Type

Private Type MapNpcPokemonRec
    '//General
    Num As Long     '//Index of the npc
    
    '//Location
    X As Long
    Y As Long
    Dir As Byte
    
    '//Vital
    CurHP As Long
    MaxHP As Long
    
    '//Shiny
    IsShiny As Byte
    
    '//Happiness
    Happiness As Byte
    
    '//Gender
    Gender As Byte
    
    '//Status
    Status As Byte
    
    '//Client Only
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
    Attacking As Byte
    AttackTimer As Long
    IdleTimer As Long
    IdleFrameTmr As Long
    IdleAnim As Byte
    '//Pokeball
    Init As Byte
    State As Byte
    Frame As Byte
    FrameState As Byte
    FrameTimer As Long
    BallX As Long
    BallY As Long
    MoveSpeed As Long
    
End Type

Private Type MapNpcRec
    '//General
    Num As Long     '//Index of the npc
    
    '//Location
    X As Long
    Y As Long
    Dir As Byte
    
    '//Client Only
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
End Type

Public Type PlayerPokemonRec
    '//General
    Num As Long
    
    '//Location
    X As Long
    Y As Long
    Dir As Byte
    
    '//For own index
    Slot As Byte
    
    '//Stat
    Stat(1 To StatEnum.Stat_Count - 1) As Long
    StatIV(1 To StatEnum.Stat_Count - 1) As Long
    StatEV(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Vital
    CurHP As Long
    MaxHP As Long
    
    '//Shiny
    IsShiny As Byte
    
    '//Happiness
    Happiness As Byte
    
    '//Gender
    Gender As Byte
    
    '//Status
    Status As Byte
    
    '//Ball Used
    BallUsed As Byte
    
    '//Confuse
    IsConfused As Byte
    
    '//Buff
    StatBuff(1 To StatEnum.Stat_Count - 1) As Long
    
    '//HeldItem
    HeldItem As Long
    
    '//Client Only
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
    Attacking As Byte
    AttackTimer As Long
    IdleTimer As Long
    IdleFrameTmr As Long
    IdleAnim As Byte
    '//Pokeball
    Init As Byte
    State As Byte
    Frame As Byte
    FrameState As Byte
    FrameTimer As Long
    BallX As Long
    BallY As Long
    MoveSpeed As Long
End Type

Private Type HotbarRec
    Num As Long
    TmrCooldown As Long
End Type

Public Type PlayerRec
    '//Identification
    Name As String * NAME_LENGTH
    
    '//General
    Sprite As Long
    Access As Byte
    
    '//Location
    Map As Long
    X As Long
    Y As Long
    Dir As Byte

    '//Vital
    CurHP As Long
    
    '//Level
    Level As Long
    CurExp As Long
    
    '//Game Data
    Money As Long
    
    '//Temp Sprite
    TempSprite As Long
    TempSpriteID As Long
    TempSpritePassiva As Long
    
    '//confuse
    IsConfuse As Byte
    
    '//Status
    Status As Byte
    
    '//Badge
    Badge(1 To MAX_BADGE) As Byte
    
    '//Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    '//Stealth Mode
    StealthMode As Byte
    
    '//Action
    Action As Byte
    ActionTmr As Long
    
    ' PvP
    win As Long
    Lose As Long
    Tie As Long
    
    '//Cash
    Cash As Long
    
    Started As Date
    TimePlay As Long
    
    '//Fish System
    FishMode As Byte
    FishRod As Byte
    
    '//Client Only
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
    
    '//Animações de Montarias
    IdleTimer As Long
    IdleFrameTmr As Long
    IdleAnim As Byte
End Type

' *****************
' ** Player Data **
' *****************
Private Type LockedRec
    Locked As Byte
    Opacity As Byte
End Type

Private Type PlayerInvRec
    Num As Long
    Value As Long
    Status As LockedRec
    ItemCooldown As Long
End Type

Public Type PokemonMovesetRec
    Num As Long
    CurPP As Byte
    TotalPP As Byte
End Type

Public Type PlayerInvStorageDataRec
    Num As Long
    Value As Long
End Type

Private Type PlayerInvStorageSlotRec
    Unlocked As Byte
    data(1 To MAX_STORAGE) As PlayerInvStorageDataRec
End Type

Public Type PlayerPokemonsRec
    Num As Long
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As Long
    StatIV(1 To StatEnum.Stat_Count - 1) As Long
    StatEV(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Vital
    CurHP As Long
    MaxHP As Long
    
    '//Nature
    Nature As Byte
    
    '//Shiny
    IsShiny As Byte
    
    '//Happiness
    Happiness As Byte
    
    '//Gender
    Gender As Byte
    
    '//Status
    Status As Byte
    
    '//Exp
    CurExp As Long
    NextExp As Long
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
    
    '//Ball Used
    BallUsed As Byte
    
    '//Held Item
    HeldItem As Long
End Type

Private Type PlayerPokemonStorageSlotRec
    Unlocked As Byte
    data(1 To MAX_STORAGE) As PlayerPokemonsRec
End Type

Private Type PlayerPokedexRec
    Scanned As Byte
    Obtained As Byte
End Type

' ************
' ** Editor **
' ************
Public Type LayerRec
    Tile As Long
    TileX As Long
    TileY As Long
    '//Animation
    MapAnim As Long
End Type

Public Type TileRec
    '//Layer
    Layer(MapLayer.Ground To MapLayer.MapLayer_Count - 1, MapLayerType.Normal To MapLayerType.Animated) As LayerRec
    '//Tile Data
    Attribute As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
End Type

Private Type MapRec
    '//General
    Revision As Long
    Name As String * NAME_LENGTH
    Moral As Byte
    
    '//Size
    MaxX As Long
    MaxY As Long
    
    '//Tiles
    Tile() As TileRec
    
    '//Map Link
    LinkUp As Long
    LinkDown As Long
    LinkLeft As Long
    LinkRight As Long
    
    '//Map Data
    Music As String * NAME_LENGTH
    
    '//Npc
    Npc(1 To MAX_MAP_NPC) As Long
    
    '//Moral
    KillPlayer As Byte
    IsCave As Byte
    CaveLight As Byte
    SpriteType As Byte
    StartWeather As Byte
    
    NoCure As Byte
End Type

' Client Only | Fill Bucket Array
Public Type TilePosRec
    Used As Boolean
    ' Position of tile
    X As Integer
    Y As Integer
End Type

Private Type NpcRec
    '//General
    Name As String * NAME_LENGTH
    Sprite As Long
    Behaviour As Byte
    Convo As Byte
    
    PokemonNum(1 To MAX_PLAYER_POKEMON) As Long
    PokemonLevel(1 To MAX_PLAYER_POKEMON) As Long
    PokemonMoveset(1 To MAX_PLAYER_POKEMON, 1 To MAX_MOVESET) As Long
    Reward As Long
    WinEvent As Long
    RewardExp As Long
    PokemonItem(1 To MAX_PLAYER_POKEMON) As Long
    PokemonNature(1 To MAX_PLAYER_POKEMON) As Integer
    PokemonIsShiny(1 To MAX_PLAYER_POKEMON) As Byte
    PokemonIvFull(1 To MAX_PLAYER_POKEMON) As Byte
    Rebatle As Byte
    SpawnWeekDay(1 To 7) As Byte
End Type

' Public Type ItemRec
'     '//General
'     Name As String * NAME_LENGTH
'     Sprite As Long
'     Stock As Byte
'     Type As Byte
'     Data1 As Long
'     Data2 As Long
'     Price As Long
'     Data3 As Long
'     Desc As String * 255
'     IsCash As Byte          'Novo método de cash no shop!
'     Linked As Byte          'Vinculado ao jogador!
'     NotEquipable As Byte    'Não equipavel ao poke.
'     Delay As Long           'Items que utilizam de Delay
'     Item(1 To MAX_MYSTERY_BOX) As Integer 'Utiliza no MysteryBox
'     ItemValue(1 To MAX_MYSTERY_BOX) As Long 'Utiliza no MysteryBox
'     ItemChance(1 To MAX_MYSTERY_BOX) As Double 'Utiliza no MysteryBox
'     Data4 As Long 'Bonus de exp no item, está sendo usado apenas em montaria
'     Data5 As Long 'Adicionado pra ser usado como um checkbox, pra ver se a montaria tem a passiva ou não
' End Type

Private Type MovesetRec
    MoveNum As Long
    MoveLevel As Long
End Type

Private Type PokemonRec
    '//General
    Name As String * NAME_LENGTH
    Sprite As Long
    ScaleSprite As Byte
    Behaviour As Byte
    
    '//Stats
    BaseStat(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Types
    PrimaryType As Byte
    SecondaryType As Byte
    
    '//Other
    CatchRate As Long
    FemaleRate As Long
    EggCycle As Byte
    EggGroup As Byte
    EvYeildType As Byte
    EvYeildVal As Byte
    BaseExp As Long
    GrowthRate As Byte
    Height As Long
    Weight As Long
    Species As String * NAME_LENGTH
    PokeDexEntry As String * 250
    
    '//Evolution
    evolveNum(1 To MAX_EVOLVE) As Long
    EvolveLevel(1 To MAX_EVOLVE) As Long
    EvolveCondition(1 To MAX_EVOLVE) As Byte
    EvolveConditionData(1 To MAX_EVOLVE) As Long
    
    '//Moveset
    Moveset(1 To MAX_POKEMON_MOVESET) As MovesetRec
    EggMoveset(1 To MAX_POKEMON_MOVESET) As Long
    Range As Byte
    DropNum(1 To MAX_DROP) As Long
    DropRate(1 To MAX_DROP) As Byte
    ItemMoveset(1 To 110) As Long
    '//Offset
    NameOffSetY As Integer
    '//Cries
    Sound As String * NAME_LENGTH
    '//Lendary
    Lendary As Byte
End Type

Private Type PokemonMoveRec
    '//General
    Name As String * NAME_LENGTH
    Type As Byte
    Category As Byte
    PP As Byte
    MaxPP As Byte
    Power As Long
    Range As Byte
    Description As String * 150
    dStat(1 To StatEnum.Stat_Count - 1) As Long
    AttackType As Byte
    targetType As Byte
    Animation As Long
    Interval As Long
    Duration As Long
    Cooldown As Long
    CastTime As Long
    AmountOfAttack As Long
    SelfAnim As Byte
    Sound As String * NAME_LENGTH
    pStatus As Byte
    pStatusChance As Byte
    RecoilDamage As Byte
    AbsorbDamage As Byte
    ChangeWeather As Byte
    BoostWeather As Byte
    StatusReq As Byte
    DecreaseWeather As Byte
    StatusToSelf As Byte
    ReflectType As Byte
    CastProtect As Byte
    SelfStatusReq As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Public Type SpawnRec
    PokeNum As Long
    MinLevel As Byte
    MaxLevel As Byte
    Respawn As Long
    SpawnTimeMin As Long
    SpawnTimeMax As Long
    
    '//Location
    MapNum As Long
    randomMap As Byte
    randomXY As Byte
    MapX As Long
    MapY As Long
    Rarity As Long
    CanCatch As Byte
    NoExp As Byte
    PokeBuff As Byte
    '//HeldItem
    HeldItem As Integer
    '//Nature
    Nature As Integer
    '//Fishing?
    Fishing As Byte
End Type

Private Type TextLangRec
    Text As String * 255
    tReply(1 To 3) As String * 100
End Type

Private Type ConvDataRec
    TextLang(1 To MAX_LANGUAGE) As TextLangRec
    '//Others
    NoText As Byte
    NoReply As Byte
    CustomScript As Byte
    CustomScriptData As Long
    CustomScriptData2 As Long
    MoveNext As Byte
    tReplyMove(1 To 3) As Byte
    CustomScriptData3 As Long
End Type

Private Type ConversationRec
    Name As String * NAME_LENGTH
    
    '//Data
    ConvData(1 To MAX_CONV_DATA) As ConvDataRec
End Type

Private Type ShopItemRec
    Num As Long
    Price As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    
    ShopItem(1 To MAX_SHOP_ITEM) As ShopItemRec
End Type

Private Type QuestRec
    Name As String * NAME_LENGTH
End Type

' *****************
' ** Client Only **
' *****************
'//Interface
Public GUI(1 To GuiEnum.Gui_Count - 1) As GuiRec
Public Button(1 To ButtonEnum.Button_Count - 1) As ButtonRec
'//Configuration
Public Resolution As ResolutionRec
Public GameSetting As SettingRec
'//Others
Public AlertWindow(1 To MAX_ALERT) As AlertRec
'//Control Key
Public ControlKey(1 To ControlEnum.Control_Count - 1) As ControlKeyRec
'//SelMenu
Public SelMenu As SelMenuRec
'//ActionMsg
Public ActionMsg(1 To 255) As ActionMsgRec
'//Animation
Public AnimInstance(1 To 255) As AnimInstanceRec
'//Credit
Public Credit() As CreditRec
'//Weather
Public Weather As WeatherRec
'//Catching
Public CatchBall(1 To MAX_GAME_POKEMON) As CatchBallRec
'//Trade

Public Type TradeDataRec
    TradeType As Byte
    
    Num As Long
    Value As Long
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As Long
    StatIV(1 To StatEnum.Stat_Count - 1) As Long
    StatEV(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Vital
    CurHP As Long
    MaxHP As Long
    
    '//Nature
    Nature As Byte
    
    '//Shiny
    IsShiny As Byte
    
    '//Happiness
    Happiness As Byte
    
    '//Gender
    Gender As Byte
    
    '//Status
    Status As Byte
    
    '//Exp
    CurExp As Long
    NextExp As Long
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
    
    '//Ball Used
    BallUsed As Byte
    
    '//Held Item
    HeldItem As Long
    
    '//Trade Slot
    TradeSlot As Byte
End Type

Public Type TradeRec
    data(1 To MAX_TRADE) As TradeDataRec
    TradeMoney As Long
    TradeSet As Byte
End Type

Private Type CatchBallRec
    InUsed As Boolean
    Pic As Byte
    X As Long
    Y As Long
    State As Byte
    Frame As Byte
    FrameState As Byte
    FrameTimer As Long
End Type

Private Type WeatherDropRec
    Pic As Long
    PicType As Byte
    X As Long
    Y As Long
    SpeedY As Long
End Type

Private Type WeatherRec
    Type As Byte
    
    InitDrop As Boolean
    MaxDrop As Long
    Drop() As WeatherDropRec
End Type

Private Type CreditRec
    Text As String
    Y As Long
    StartY As Long
End Type

Private Type ControlKeyRec
    keyName As String
    cAsciiKey As Long
End Type

Private Type ButtonRec
    '//General
    StartX(0 To 2) As Long
    StartY(0 To 2) As Long
    
    '//Location
    X As Long
    Y As Long
    
    '//Size
    Height As Long
    Width As Long
    
    '//State
    State As Byte
End Type

Private Type GuiRec
    '//General
    Visible As Boolean
    Pic As Byte
    
    '//Location
    X As Long
    Y As Long
    OrigX As Long
    OrigY As Long
    
    '//Size
    StartX As Long
    StartY As Long
    Height As Long
    Width As Long
    
    '//Dragable
    InDrag As Boolean
    OldMouseX As Long
    OldMouseY As Long
End Type

Private Type ResolutionDataRec
    Width As Long
    Height As Long
End Type

Private Type ResolutionRec
    MaxResolution As Integer
    ResolutionSize() As ResolutionDataRec
End Type

Private Type SettingRec
    '//GUI
    ThemePath As String
    
    '//Video
    'Resolution As Byte
    Fullscreen As Byte
    Width As Long
    Height As Long
    
    '//Network
    RemoteHost As String
    RemotePort As Long
    
    '//Others
    SkipBootUp As Byte
    ShowFPS As Byte
    ShowPing As Byte
    ShowName As Byte
    ShowPP As Byte
    
    '//Account
    Username As String
    Password As String
    SavePass As Byte
    
    '//Sound
    Background As Byte
    SoundEffect As Byte
    MenuMusic As String
    
    '//Language
    CurLanguage As Byte
End Type


Private Type AlertRec
    '//General
    IsUsed As Boolean
    
    Text As String
    Color As Long
    
    '//Size and Location
    Width As Long
    Height As Long
    
    SetYPos As Long
    CurYPos As Long
    
    '//Determine how long the window will be visible
    AlertTimer As Long
End Type

Public Type ChatBubbleRec
    Msg As String
    Colour As Long
    target As Long
    targetType As Byte
    X As Long
    Y As Long
    
    '//Client data only
    timer As Long
    active As Boolean
End Type

Private Type SelMenuRec
    '//General
    Visible As Boolean
    Type As Byte
    
    '//Location
    X As Long
    Y As Long
    
    '//Text
    MaxText As Byte
    Text() As String
    MaxWidth As Long
    CurPick As Byte
    
    '//Data
    Data1 As Long
End Type

Private Type ActionMsgRec
    Msg As String
    Created As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    timer As Long
    Alpha As Long
End Type

'//Animation
Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    '//timing
    timer(0 To 1) As Long
    '//rendering check
    Used(0 To 1) As Boolean
    '//counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

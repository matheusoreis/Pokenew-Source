Attribute VB_Name = "modType"
Option Explicit

' *************
' ** General **
' *************
Public Player() As PlayerRec
Public PlayerInv() As PlayerInvRec
Public PlayerPokemons() As PlayerPokemonsRec
Public PlayerInvStorage() As PlayerInvStorageRec
Public PlayerPokemonStorage() As PlayerPokemonStorageRec
Public PlayerPokedex() As PlayerPokedexRec
'//Player Pokemon
Public PlayerPokemon() As PlayerPokemonRec
Public Map(1 To MAX_MAP) As MapRec
Public Npc(1 To MAX_NPC) As NpcRec
Public MapNpc(1 To MAX_MAP, 1 To MAX_MAP_NPC) As MapNpcRec
Public MapNpcPokemon(1 To MAX_MAP, 1 To MAX_MAP_NPC) As MapNpcPokemonRec
Public Pokemon(1 To MAX_POKEMON) As PokemonRec
Public MapPokemon(1 To MAX_GAME_POKEMON) As MapPokemonRec

Public PokemonMove(1 To MAX_POKEMON_MOVE) As PokemonMoveRec
Public Animation(1 To MAX_ANIMATION) As AnimationRec
Public Spawn(1 To MAX_GAME_POKEMON) As SpawnRec
Public Conversation(1 To MAX_CONVERSATION) As ConversationRec
Public Shop(1 To MAX_SHOP) As ShopRec
Public Quest(1 To MAX_QUEST) As QuestRec
Public Rank(1 To MAX_RANK) As RankRec

'//Fishing system
Public Fishing(1 To MAX_MAP) As FishingRec

'//Event Xp
Public EventExp As EventExpRec
'//Virtual Shop
Public VirtualShop(1 To VirtualShopTabsRec.CountTabs - 1) As VirtualShopDataRec

Public Type TempSpriteRec
    TempSpriteType As Long
    TempSpriteID As Long
    TempSpriteExp As Long
    TempSpritePassiva As Long
End Type

Private Type FishingRec
    Pokemon() As Long
End Type

Public Type MysteryBoxRec
    Num As Integer
    Quant As Long
    Chance As Double
End Type

Private Type VirtualShopRec
    ItemNum As Long
    ItemQuant As Long
    ItemPrice As Long
    CustomDesc As Byte
End Type

Private Type VirtualShopDataRec
    Items() As VirtualShopRec
    Max_Slots As Integer
End Type

Private Type EventExpRec
    ExpEvent As Boolean
    ExpMultiply As Byte
    ExpSecs As Long
End Type

Private Type RankRec
    Name As String
    Level As Long
    Exp As Long
End Type

'//Stats
Private Type StatDataRec
    Value As Long
    EV As Long
    IV As Long
End Type

' **************
' ** Map Data **
' **************
Private Type PokemonMovesetRec
    Num As Long
    CurPP As Byte
    TotalPP As Byte
    CD As Long
End Type

Private Type MapPokemonRec
    '//General
    Num As Long
    
    '//Location
    Map As Long
    x As Long
    Y As Long
    Dir As Byte
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As StatDataRec
    
    '//Vital
    CurHp As Long
    MaxHp As Long
    
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
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
    
    '//Server Side
    PokemonIndex As Long
    Respawn As Long
    MoveTmr As Long
    AtkTmr As Long
    targetType As Long
    TargetIndex As Long
    
    '//Buff/Debuff
    StatBuff(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Move
    QueueMove As Long
    QueueMoveSlot As Byte
    MoveDuration As Long
    MoveInterval As Long
    MoveCastTime As Long
    MoveAttackCount As Long
    NextCritical As Byte
    
    '//Catch
    InCatch As Byte
    
    '//Status
    StatusDamage As Long
    StatusMove As Byte
    LastAttacker As Long
    
    '//Confuse
    IsConfuse As Byte
    
    '//Reflect
    ReflectMove As Byte
    IsProtect As Byte
    
    '//HeldItem
    HeldItem As Long
End Type

Private Type MapNpcPokemonRec
    '//General
    Num As Long     '//Index of the npc
    
    '//Location
    x As Long
    Y As Long
    Dir As Byte
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As StatDataRec
    
    '//Vital
    CurHp As Long
    MaxHp As Long
    
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
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
    
    '//Data
    MoveTmr As Long
    AtkTmr As Long
    
    '//Buff/Debuff
    StatBuff(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Move
    QueueMove As Long
    QueueMoveSlot As Byte
    MoveDuration As Long
    MoveInterval As Long
    MoveCastTime As Long
    MoveAttackCount As Long
    NextCritical As Byte
    
    '//Status
    StatusDamage As Long
    StatusMove As Byte
    
    '//Confuse
    IsConfuse As Byte
    
    '//Reflect
    ReflectMove As Byte
    IsProtect As Byte
    
    '//Held Item
    HeldItem As Long
End Type

Private Type MapNpcRec
    '//General
    Num As Long     '//Index of the npc
    
    '//Location
    x As Long
    Y As Long
    Dir As Byte
    
    '//Data
    MoveTmr As Long
    
    '//Pokemon
    InBattle As Long
    PokemonAlive(1 To MAX_PLAYER_POKEMON) As Byte
    CurPokemon As Byte
    FaintWaitTimer As Long
End Type

Public Type PlayerPokemonRec
    '//General
    Num As Long
    
    '//Location
    x As Long
    Y As Long
    Dir As Byte
    
    '//Own index
    slot As Byte
    MoveTmr As Long
    AtkTmr As Long
    
    '//Move
    QueueMove As Long
    QueueMoveSlot As Byte
    MoveDuration As Long
    MoveInterval As Long
    MoveCastTime As Long
    MoveAttackCount As Long
    NextCritical As Byte
    
    '//Buff/Debuff
    StatBuff(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Status
    StatusDamage As Long
    StatusMove As Byte
    
    '//Confuse
    IsConfuse As Byte
    
    '//Reflect
    ReflectMove As Byte
    IsProtect As Byte
End Type

Private Type NpcBattledRec
    NpcBattledAt As Byte
    Win As Byte
End Type

Public Type PlayerRec
    '//Identification
    Name As String * NAME_LENGTH
    
    '//General
    Sprite As Long
    Access As Byte
    
    '//Location
    Map As Long
    x As Long
    Y As Long
    Dir As Byte
    
    '//Vital
    CurHp As Long
    
    '//Level
    Level As Long
    CurExp As Long
    
    '//Money
    Money As Long
    
    '//PvP
    Win As Long
    Lose As Long
    Tie As Long
    
    '//Temp Sprite
    KeyItemNum As Long
    
    '//Status
    Status As Byte
    
    '//Confuse
    IsConfuse As Byte
    
    Muted As Byte
    
    '//Action
    Action As Byte
    ActionTmr As Long
    
    '//Checkpoint
    CheckMap As Long
    CheckX As Long
    CheckY As Long
    CheckDir As Byte
    
    '//Stealth Mode
    StealthMode As Byte
    
    '//Badge
    Badge(1 To MAX_BADGE) As Byte
    
    '//NPC Battle
    NpcBattledMonth(1 To MAX_NPC) As NpcBattledRec
    NpcBattledDay(1 To MAX_NPC) As NpcBattledRec

    '//Switches
    Switches(1 To MAX_SWITCH) As Byte
    
    '//Hotbar
    Hotbar(1 To MAX_HOTBAR) As Long
    
    '//ServerSide
    DidStart As Byte
    MoveTmr As Long
    
    '//Cash
    Cash As Long
    
    '//Jour Init
    Started As Date
    TimePlay As Long
    
    '//Fish System
    FishMode As Byte
    FishRod As Byte
End Type

' *****************
' ** Player Data **
' *****************
'//Player Inv
Public Type PlayerInvDataRec
    Num As Long
    Value As Long
    Locked As Byte
    TmrCooldown As Long
End Type

Public Type PlayerInvRec
    Data(1 To MAX_PLAYER_INV) As PlayerInvDataRec
End Type

Public Type PlayerInvStorageDataRec
    Num As Long
    Value As Long
    TmrCooldown As Long
End Type

Private Type PlayerInvStorageSlotRec
    Unlocked As Byte
    Data(1 To MAX_STORAGE) As PlayerInvStorageDataRec
End Type

Public Type PlayerInvStorageRec
    slot(1 To MAX_STORAGE_SLOT) As PlayerInvStorageSlotRec
End Type

Public Type PlayerPokemonStorageDataRec
    Num As Long
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As StatDataRec
    
    '//Vital
    CurHp As Long
    MaxHp As Long
    
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
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
        
    '//Ball Used
    BallUsed As Byte
    
    '//Held Item
    HeldItem As Long
End Type

Private Type PlayerPokemonStorageSlotRec
    Unlocked As Byte
    Data(1 To MAX_STORAGE) As PlayerPokemonStorageDataRec
End Type

Public Type PlayerPokemonStorageRec
    slot(1 To MAX_STORAGE_SLOT) As PlayerPokemonStorageSlotRec
End Type

'//Player Pokemon
Private Type PlayerPokemonsDataRec
    Num As Long
    
    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As StatDataRec
    
    '//Vital
    CurHp As Long
    MaxHp As Long
    
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
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
        
    '//Ball Used
    BallUsed As Byte
    
    '//Held Item
    HeldItem As Long
End Type

Public Type PlayerPokemonsRec
    Data(1 To MAX_PLAYER_POKEMON) As PlayerPokemonsDataRec
End Type

Private Type PokedexDataRec
    Scanned As Byte
    Obtained As Byte
End Type

Public Type PlayerPokedexRec
    PokemonIndex(1 To MAX_POKEMON) As PokedexDataRec
End Type

' ************
' ** Editor **
' ************
Private Type LayerRec
    Tile As Long
    TileX As Long
    TileY As Long
    '//Animation
    MapAnim As Long
End Type

Private Type TileRec
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
    
    '//Server Side
    CurWeather As Byte
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
    EvolveNum(1 To MAX_EVOLVE) As Long
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
    
    '//OffSet
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

Private Type SpawnRec
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
    pokeBuff As Byte
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
' ** Server Only **
' *****************
Public Account() As AccountRec
Public TempPlayer() As TempPlayerRec
Public MapCache(1 To MAX_MAP) As DataRec
Public Options As OptionRec

Private Type OptionRec
    '//Network
    Port As Long
    
    '//Debug Mode
    DebugMode As Byte
    
    '//Starting Position
    StartMap As Long
    startX As Long
    startY As Long
    StartDir As Byte
    ExpRate As Long
    
    '//MOTD
    MOTD As String
    
    '//Other
    ShinyRarity As Long
    
    '//Requerimentos pra trade
    TradeLvlMin As Integer
    SameIp As Byte
    
    '//Dados Pokemon p/ GlobalMsg
    Rarity As Integer
End Type

Private Type DataRec
    Data() As Byte
End Type

Public Type AccountRec
    Username As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String * TEXT_LENGTH
    
    '//Character Count
    CharCount As Byte
End Type

Private Type TradeDataRec
    '//Determine if Pokemon or Item
    Type As Byte
    
    '//Data
    Num As Long
    Value As Long

    '//Stats
    Level As Byte
    Stat(1 To StatEnum.Stat_Count - 1) As Long
    StatIV(1 To StatEnum.Stat_Count - 1) As Long
    StatEV(1 To StatEnum.Stat_Count - 1) As Long
    
    '//Vital
    CurHp As Long
    MaxHp As Long
    
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
    nextExp As Long
    
    '//Moveset
    Moveset(1 To MAX_MOVESET) As PokemonMovesetRec
    
    '//Ball Used
    BallUsed As Byte
    
    '//Held Item
    HeldItem As Long
    
    '//TradeSlot
    TradeSlot As Byte
End Type

Public Type TempPlayerRec
    InGame As Boolean
    
    UseChar As Byte
    
    ' Non saved local vars
    buffer As clsBuffer

    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    
    GettingMap As Boolean
    
    '//New Pokemon Move
    MoveLearnPokeSlot As Byte
    MoveLearnNum As Long
    MoveLearnIndex As Byte
    
    '//Use Item
    TmpUseInvSlot As Byte
    
    '//Catching
    TmpCatchPokeNum As Long
    TmpCatchTries As Byte
    TmpCatchTimer As Long
    TmpCatchValue As Long
    TmpCatchUseBall As Byte
    
    '//Storage
    StorageType As Byte
    
    '//Shop
    InShop As Long
    
    '//Language
    CurLanguage As Byte
    
    '//Convo
    CurConvoNum As Long
    CurConvoData As Byte
    CurConvoNpc As Long
    CurConvoMapNpc As Long
    
    PlayerRequest As Long
    RequestType As Byte
    '//Duel
    InDuel As Long
    DuelTime As Long
    DuelTimeTmr As Long
    WarningTimer As Long
    InNpcDuel As Long
    DuelReset As Byte
    
    '//Trade
    InTrade As Long
    TradeItem(1 To MAX_TRADE) As TradeDataRec
    TradeMoney As Long
    TradeSet As Byte
    TradeAccept As Byte
    
    '//Map Switch Timer
    MapSwitchTmr As Byte
    
    '//Party
    InParty As Byte
    PartyIndex(1 To MAX_PARTY) As Long
    
    '//ProcessorID
    ProcessorID As String
    PingTimer As Long
    PingCount As Long
    
    '//Temp Sprite Bonuses
    TempSprite As TempSpriteRec
End Type

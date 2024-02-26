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
    
    ' Não editores
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

Public HandleDataSub(CMSG_Count) As Long

'//Map Layers
Public Enum MapLayer
    Ground = 0
    Mask
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

'//Types
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

'//Ball Type
Public Enum BallEnum
    b_Pokeball = 0
    b_Greatball
    b_Ultraball
    b_Masterball
    b_Primerball
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

'//Loja Virtual
Public Enum VirtualShopTabsRec
    Skins = 1
    Mounts
    Items
    Vips
    
    CountTabs
End Enum

'//Spawn Npc
Public Enum WeekDayEnum
    Domingo = 1
    Segunda
    Terça
    Quarta
    Quinta
    Sexta
    Sabado
    
    Count_WeekDay
End Enum










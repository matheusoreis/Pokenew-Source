Attribute VB_Name = "SharedConstants"
Option Explicit

' ******************
' ** API Declares **
' ******************
'//Use for copying data
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'//This use for clearing data
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
'//Text API
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
'//Socket
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lparam As Long) As Long


' *************
' ** General **
' *************
'//Text Color Pointers
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16

'//Use to declare as True or False for numerical variable
Public Const NO As Byte = 0
Public Const YES As Byte = 1

'//Constant
Public Const MAX_PLAYERCHAR As Byte = 3
Public Const MAX_NPC As Byte = 10
Public Const MAX_MAP_NPC As Byte = 10
Public Const MAX_POKEMON As Long = 10 '151
Public Const MAX_GAME_POKEMON As Long = 10
Public Const MAX_ITEM As Long = 10 '100
Public Const MAX_PLAYER_INV As Byte = 35
Public Const MAX_PLAYER_POKEMON As Byte = 6
Public Const MAX_POKEMON_MOVE As Long = 10 '100
Public Const MAX_POKEMON_MOVESET As Byte = 30
Public Const MAX_ANIMATION As Byte = 10 '100
Public Const MAX_MOVESET As Byte = 4
Public Const MAX_LEVEL As Byte = 100
Public Const MAX_EVOLVE As Byte = 10
Public Const MAX_DISTANCE As Byte = 9
Public Const MAX_STORAGE_SLOT As Byte = 5
Public Const MAX_STORAGE As Byte = 42
Public Const MAX_CONVERSATION As Long = 10
Public Const MAX_CONV_DATA As Byte = 10
Public Const MAX_LANGUAGE As Byte = 3
Public Const MAX_SHOP As Byte = 10
Public Const MAX_SHOP_ITEM As Byte = 20
Public Const MAX_TRADE As Byte = 16
Public Const MAX_SWITCH As Long = 10
Public Const MAX_DROP As Byte = 5
Public Const MAX_MONEY As Long = 999999999
Public Const MAX_CASH As Long = 9999
Public Const MAX_AMOUNT As Long = 999
Public Const MAX_BADGE As Byte = 32
Public Const MAX_HOTBAR As Byte = 5
Public Const MAX_QUEST As String = 10
Public Const MAX_PARTY As Byte = 4
Public Const MAX_PLAYER_LEVEL As Byte = 250
Public Const MAX_RANK As Byte = 20

'//String constants
Public Const NAME_LENGTH As Byte = 25
Public Const TEXT_LENGTH As Byte = 150

'//Map Constant
Public Const MAX_MAP As Long = 5
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18

'//Gender
Public Const GENDER_MALE As Byte = 0
Public Const GENDER_FEMALE As Byte = 1

'//Direction
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

'//Access
Public Const ACCESS_NONE As Byte = 0
Public Const ACCESS_MODERATOR As Byte = 1
Public Const ACCESS_MAPPER As Byte = 2
Public Const ACCESS_DEVELOPER As Byte = 3
Public Const ACCESS_CREATOR As Byte = 4
'Public Const ACCESS_HIDDEN As Byte = 5

'//Npc Behaviour
Public Const BEHAVIOUR_NONE As Byte = 0
Public Const BEHAVIOUR_MOVE As Byte = 1

'//Target Type
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2
Public Const TARGET_TYPE_PLAYERPOKEMON As Byte = 3

'//Language
Public Const LANG_PT As Byte = 0
Public Const LANG_EN As Byte = 1
Public Const LANG_ES As Byte = 2

'//Conversation Custom Script
Public Const CONVO_SCRIPT_NONE As Byte = 0
Public Const CONVO_SCRIPT_INVSTORAGE As Byte = 1
Public Const CONVO_SCRIPT_POKESTORAGE As Byte = 2
Public Const CONVO_SCRIPT_HEAL As Byte = 3
Public Const CONVO_SCRIPT_SHOP As Byte = 4
Public Const CONVO_SCRIPT_SETSWITCH As Byte = 5
Public Const CONVO_SCRIPT_GIVEPOKE As Byte = 6
Public Const CONVO_SCRIPT_GIVEITEM As Byte = 7
Public Const CONVO_SCRIPT_WARPTO As Byte = 8
Public Const CONVO_SCRIPT_CHECKMONEY As Byte = 9
Public Const CONVO_SCRIPT_TAKEMONEY As Byte = 10
Public Const CONVO_SCRIPT_STARTBATTLE As Byte = 11
Public Const CONVO_SCRIPT_RELEARN As Byte = 12
Public Const CONVO_SCRIPT_GIVEBADGE As Byte = 13
Public Const CONVO_SCRIPT_CHECKBADGE As Byte = 14
Public Const CONVO_SCRIPT_BEATPOKE As Byte = 15
Public Const CONVO_SCRIPT_CHECKITEM As Byte = 16
Public Const CONVO_SCRIPT_TAKEITEM As Byte = 17
Public Const CONVO_SCRIPT_RESPAWNPOKE As Byte = 18
Public Const CONVO_SCRIPT_CHECKLEVEL As Byte = 19
Public Const MAX_CONVO_SCRIPT As Byte = 19

'//Evolve Condition
Public Const EVOLVE_CONDT_NONE As Byte = 0
Public Const EVOLVE_CONDT_TIME As Byte = 1
Public Const EVOLVE_CONDT_HAPPINESS As Byte = 2
Public Const EVOLVE_CONDT_TRADE As Byte = 3
Public Const EVOLVE_CONDT_GENDER As Byte = 4
Public Const EVOLVE_CONDT_ITEM As Byte = 5
Public Const EVOLVE_CONDT_KNOWMOVE As Byte = 6
Public Const EVOLVE_CONDT_AREA As Byte = 7
Public Const MAX_EVOLVE_CONDT As Byte = 7

'//Temp Sprite Const
'Public Const TEMP_SPRITE_GROUP_NONE As Byte = 0
'Public Const TEMP_SPRITE_GROUP_DIVE As Byte = 1
'Public Const TEMP_SPRITE_GROUP_BIKE As Byte = 2
'Public Const TEMP_SPRITE_GROUP_SURF As Byte = 3
'Public Const TEMP_SPRITE_GROUP_MOUNT As Byte = 4
'Public Const TEMP_FISH_MODE As Byte = 5

'//Player Action
Public Const ACTION_SLIDE As Byte = 1

' **********************
' ** Server Side Only **
' **********************
'//Server Data
Public Const GAME_NAME As String = "PokeNew"
Public Const GAME_PORT As Long = 8007

Public Const DC_TIMER As Long = 30000 ' 30 seconds waiting before disconnection

'//Starting Location
Public Const START_MAP As Byte = 1
Public Const START_X As Byte = 11
Public Const START_Y As Byte = 7

Public Const MAX_EV As Long = 510

'//Map Constants
Public Const MAP_MORAL_DANGER As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_ARENA As Byte = 2
Public Const MAP_MORAL_SAFARI As Byte = 3
Public Const MAP_MORAL_PVP As Byte = 4

'//Game Constants
Public Const INV_SLOTS_LOCKED As Byte = 10 ' Slots
Public Const INV_SLOTS_PRICE As Byte = 5 ' Cash

Public Const ColourChar As String * 1 = "½"

'//Rebatle Options
Public Const REBATLE_NONE As Byte = 0
Public Const REBATLE_LOSE As Byte = 1
Public Const REBATLE_NEVER As Byte = 2

'//MysteryBox
Public Const MAX_MYSTERY_BOX As Byte = 30

' Editor de Itens
Public Const MAX_SPRITE_ITENS As Long = 547

Public Const MAX_MAPS_REQUIREMENTS As Byte = 50

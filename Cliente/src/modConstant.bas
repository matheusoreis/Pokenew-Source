Attribute VB_Name = "modConstant"
Option Explicit

'// Server List
Public Const MAX_SERVER_LIST As Integer = 3

' ******************
' ** API Declares **
' ******************
'//Use for copying data
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'//This use for clearing data
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
'//Use for setting Window zOrder
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'//Text API
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
'//Socket
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lparam As Long) As Long
'//Resolution
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'//Checking Keyboard Press
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'//Checking of window is currently active
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

' ******************
' ** API Variable **
' ******************
'//Keeping form on top
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

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
Public Const Dark As Byte = 17

'//Use to declare as True or False for numerical variable
Public Const NO As Byte = 0
Public Const YES As Byte = 1

'//Constant
Public Const MAX_PLAYERCHAR As Byte = 3
Public Const MAX_NPC As Byte = 100
Public Const MAX_MAP_NPC As Byte = 35
Public Const MAX_POKEMON As Long = 1000 '151
Public Const MAX_GAME_POKEMON As Long = 2000
Public Const MAX_ITEM As Long = 600 '100
Public Const MAX_PLAYER_INV As Byte = 35
Public Const MAX_PLAYER_POKEMON As Byte = 6
Public Const MAX_POKEMON_MOVE As Long = 1000 '100
Public Const MAX_POKEMON_MOVESET As Byte = 30
Public Const MAX_ANIMATION As Byte = 200 '100
Public Const MAX_MOVESET As Byte = 4
Public Const MAX_LEVEL As Byte = 100
Public Const MAX_EVOLVE As Byte = 10
Public Const MAX_DISTANCE As Byte = 9
Public Const MAX_STORAGE_SLOT As Byte = 5
Public Const MAX_STORAGE As Byte = 42
Public Const MAX_CONVERSATION As Long = 500
Public Const MAX_CONV_DATA As Byte = 10
Public Const MAX_LANGUAGE As Byte = 3
Public Const MAX_SHOP As Byte = 100
Public Const MAX_SHOP_ITEM As Byte = 20
Public Const MAX_TRADE As Byte = 16
Public Const MAX_SWITCH As Long = 1000
Public Const MAX_DROP As Byte = 5
Public Const MAX_BADGE As Byte = 32
Public Const MAX_HOTBAR As Byte = 5
Public Const MAX_QUEST As String = 100
Public Const MAX_PARTY As Byte = 4
Public Const MAX_RANK As Byte = 20

'//String constants
Public Const NAME_LENGTH As Byte = 25
Public Const TEXT_LENGTH As Byte = 150

'//Map Constant
Public Const MAX_MAP As Long = 200
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
Public Const ACCESS_HIDDEN As Byte = 5

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
Public Const TEMP_SPRITE_GROUP_NONE As Byte = 0
Public Const TEMP_SPRITE_GROUP_DIVE As Byte = 1
Public Const TEMP_SPRITE_GROUP_BIKE As Byte = 2
Public Const TEMP_SPRITE_GROUP_SURF As Byte = 3
Public Const TEMP_SPRITE_GROUP_MOUNT As Byte = 4
Public Const TEMP_FISH_MODE As Byte = 5

'//Player Action
Public Const ACTION_SLIDE As Byte = 1

' *****************
' ** Client Only **
' *****************
'//Client Data
Public Const GAME_NAME As String = "PokeReborn"

'//Default Resolution Screen
Public Const Default_ScreenWidth As Long = 1280
Public Const Default_ScreenHeight As Long = 720

'//File Extension Name
Public Const GFX_EXT As String = ".png"
Public Const DATA_EXT As String = ".dat"

'//Load Stuff
Public Const LOAD_STRING_LENGTH As Long = 350

'//Alert Stuff
Public Const MAX_ALERT As Byte = 10
Public Const ALERT_STRING_LENGTH As Long = 350
Public Const ALERT_TIMER As Long = 5000 ' 5 Seconds

'//Menu State
Public Const MENU_STATE_REGISTER As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_ADDCHAR As Byte = 3
Public Const MENU_STATE_USECHAR As Byte = 4
Public Const MENU_STATE_DELCHAR As Byte = 5

'//ChoiceBox
Public Const CB_EXIT As Byte = 1
Public Const CB_CHARDEL As Byte = 2
Public Const CB_RETURNMENU As Byte = 3
Public Const CB_SAVESETTING As Byte = 4
Public Const CB_EVOLVE As Byte = 5
Public Const CB_REQUEST As Byte = 6
Public Const CB_RELEASE As Byte = 7
Public Const CB_BUYSLOT As Byte = 8
Public Const CB_FLY As Byte = 9
Public Const CB_BUYINV As Byte = 10

'//InputBox
Public Const IB_NEWPASSWORD As Byte = 1
Public Const IB_PASSWORDCONFIRM As Byte = 2
Public Const IB_OLDPASSWORD As Byte = 3
Public Const IB_DEPOSIT As Byte = 4
Public Const IB_WITHDRAW As Byte = 5
Public Const IB_BUYITEM As Byte = 6
Public Const IB_SELLITEM As Byte = 7
Public Const IB_ADDTRADE As Byte = 8

'//Tile Size
Public Const TILE_X As Long = 32
Public Const TILE_Y As Long = 32
Public Const PIC_X As Byte = 16
Public Const PIC_Y As Byte = 16

'//Chatbox
Public Const MAX_CHAT_TEXT As Long = 180

'//Map Anim
Public Const MAX_MAP_FRAME As Byte = 100

'//Chatbubble
Public Const ChatBubbleWidth As Long = 200

'//App Version
Public Const APP_MAJOR As Long = 1
Public Const APP_MINOR As Long = 1
Public Const APP_REVISION As Long = 85

'//Player Inv
Public Const MAX_INV_VISIBLE As Byte = 7

'//Pokedex Scrolling
Public Const PokedexScrollLength As Byte = 167
Public Const PokedexScrollSize As Byte = 35
Public Const PokedexScrollStartY As Byte = 55
Public Const PokedexScrollEndY As Byte = 167

'//Ranking Scrolling
Public Const RankingScrollViewLine As Byte = 8  ' Quantidade de linha para visualizar
Public Const RankingScrollLength As Byte = 164  ' Tamanho do Scroll do começo dele a parte superior do final
Public Const RankingScrollSize As Byte = 35     ' Tamanho do botão do Scroll
Public Const RankingScrollStartY As Byte = 63   ' Posição Inicial do Scroll
Public Const RankingScrollEndY As Byte = 199    ' Posição final do Scroll

' Controles Scrool
Public Const ControlScrollViewLine As Byte = 7
Public Const ControlScrollLength As Byte = 185
Public Const ControlScrollSize As Byte = 35
Public Const ControlScrollStartY As Byte = 48
Public Const ControlScrollEndY As Byte = 176

' Quantidade da lista de resolução
Public Const MAX_RESOLUTION_LIST As Integer = 6

'//Items texture ID
Public Const IDMoney As Integer = 526
Public Const IDCash As Integer = 527

'//Poke using a Item Texture
Public Const PokeUseHeld As Integer = 531

Public Const ColourChar As String * 1 = "½"

'//Map Constants
Public Const MAP_MORAL_DANGER As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_ARENA As Byte = 2
Public Const MAP_MORAL_SAFARI As Byte = 3
Public Const MAP_MORAL_PVP As Byte = 4

'//Game Constants
Public Const INV_SLOTS_LOCKED As Byte = 10 ' Slots
Public Const INV_SLOTS_PRICE As Byte = 5 ' Cash


'//MysteryBox
Public Const MAX_MYSTERY_BOX As Byte = 30
Public Const MAX_MAPS_REQUIREMENTS As Byte = 50

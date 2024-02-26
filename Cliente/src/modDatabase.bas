Attribute VB_Name = "modDatabase"
Option Explicit

'//Gets a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String   '//Max string length
Dim szReturn As String  '//Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'//Writes a variable to a text file
Public Sub PutVar(file As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString$(Header, Var, value, file)
End Sub

'//This check the directory if exist, if not, then create one
Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' Checking of Directory Exist, Create if not
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

'//This check if the file exist
Public Function FileExist(ByVal FileName As String) As Boolean
    ' Checking if File Exist
    If LenB(Dir(FileName)) > 0 Then FileExist = True
End Function

Public Function DirExist(ByVal tDir As String) As Boolean
    Dim strPastaExiste As String
    
    strPastaExiste = Dir(tDir, vbDirectory)
    
    If strPastaExiste = "" Then
        DirExist = False
    Else
        DirExist = True
    End If
End Function

'//This sub delete the file, if it doesn't exist then it will ignore it
Public Sub DeleteFile(ByVal FileName As String)
    On Error Resume Next
    Kill FileName
End Sub

' ************************
' ** Game Configuration **
' ************************
Public Sub ClearSetting()
    Call ZeroMemory(ByVal VarPtr(GameSetting), LenB(GameSetting))
    GameSetting.ThemePath = vbNullString
    GameSetting.MenuMusic = "None."
End Sub

Public Sub LoadSetting()
Dim FileName As String

    On Error GoTo errorHandler

    '//Find our file
    FileName = App.path & "\data\config\setting.ini"
  
    '//Check if our file exist
    If Not FileExist(FileName) Then
        '//Clear the data first so we will have the default data before saving
        Call ClearSetting
        Call SaveSetting
        Exit Sub
    End If

    With GameSetting
        '//GUI
        .ThemePath = Trim$(GetVar(FileName, "GUI", "Path"))

        '//Video
        .Fullscreen = GetVar(FileName, "Video", "Fullscreen")
        .Width = GetVar(FileName, "Video", "Width")
        .Height = GetVar(FileName, "Video", "Height")
  
        '//Others
        .SkipBootUp = GetVar(FileName, "Other", "SkipBootUp")
        .ShowFPS = GetVar(FileName, "Other", "ShowFPS")
        .ShowPing = GetVar(FileName, "Other", "ShowPing")
        .ShowName = GetVar(FileName, "Other", "ShowName")
        .ShowPP = GetVar(FileName, "Other", "ShowPP")

        '//Account
        .Username = Trim$(GetVar(FileName, "Account", "Username"))
        .Password = Trim$(GetVar(FileName, "Account", "Password"))
        .SavePass = GetVar(FileName, "Account", "SavePass")

        '//Sound
        .Background = GetVar(FileName, "Sound", "Background")
        .SoundEffect = GetVar(FileName, "Sound", "SoundEffect")
        .MenuMusic = Trim$(GetVar(FileName, "Sound", "MenuMusic"))
    
        '//Language
        .CurLanguage = GetVar(FileName, "Language", "CurLanguage")
    End With
    
    '//Update interface base on settings
    If GameSetting.Fullscreen = YES Then
        frmMain.BorderStyle = 0 ' None
    ElseIf GameSetting.Fullscreen = NO Then
        frmMain.BorderStyle = 1 ' Fixed
    End If
    
    Exit Sub
errorHandler:
    MsgBox Err.Number & vbNewLine & Err.Description
    MsgBox "Failed to load game settings. Exiting...", vbCritical
    UnloadMain
End Sub

Public Sub SaveSetting()
Dim FileName As String

    '//Find our file
    FileName = App.path & "\data\config\setting.ini"
    
    '//Make sure that file doesn't duplicate
    If FileExist(FileName) Then DeleteFile FileName
    
    With GameSetting
        '//GUI
        Call PutVar(FileName, "GUI", "Path", Trim$(.ThemePath))
        
        '//Video
        Call PutVar(FileName, "Video", "Fullscreen", Str(.Fullscreen))
        Call PutVar(FileName, "Video", "Width", Str(.Width))
        Call PutVar(FileName, "Video", "Height", Str(.Height))
        
        '//Others
        Call PutVar(FileName, "Other", "SkipBootUp", Str(.SkipBootUp))
        Call PutVar(FileName, "Other", "ShowFPS", Str(.ShowFPS))
        Call PutVar(FileName, "Other", "ShowPing", Str(.ShowPing))
        Call PutVar(FileName, "Other", "ShowName", Str(.ShowName))
        Call PutVar(FileName, "Other", "ShowPP", Str(.ShowPP))
        
        '//Account
        Call PutVar(FileName, "Account", "Username", Trim$(.Username))
        Call PutVar(FileName, "Account", "Password", Trim$(.Password))
        Call PutVar(FileName, "Account", "SavePass", Str(.SavePass))
        
        '//Sound
        Call PutVar(FileName, "Sound", "Background", Str(.Background))
        Call PutVar(FileName, "Sound", "SoundEffect", Str(.SoundEffect))
        Call PutVar(FileName, "Sound", "MenuMusic", Trim$(.MenuMusic))
        
        '//Language
        Call PutVar(FileName, "Language", "CurLanguage", Str(.CurLanguage))
    End With
End Sub

' *********
' ** GUI **
' *********
Public Sub ResetAllGuiLocation()
Dim i As Long

    '//Loop through all items
    For i = 1 To GuiEnum.Gui_Count - 1
        ResetGuiLocation i
    Next
End Sub

Public Sub ResetGuiLocation(ByVal vGui As GuiEnum)
    '// Starting Location will be given by code
    GUI(vGui).X = GUI(vGui).OrigX
    GUI(vGui).Y = GUI(vGui).OrigY
End Sub

Public Sub ResetGui()
    '//GUI
    With GUI(GuiEnum.GUI_LOGIN)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2) '+ 100
    End With
    With GUI(GuiEnum.GUI_REGISTER)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2) '+ 100
    End With
    With GUI(GuiEnum.GUI_CHARACTERSELECT)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2) '+ 100
    End With
    With GUI(GuiEnum.GUI_CHARACTERCREATE)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2) '+ 100
    End With
    With GUI(GuiEnum.GUI_CHOICEBOX)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_GLOBALMENU)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_OPTION)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_CHATBOX)
        .OrigX = 10
        .OrigY = Screen_Height - .Height - 10
    End With
    With GUI(GuiEnum.GUI_MOVEREPLACE)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_INPUTBOX)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_INVENTORY)
        .OrigX = Screen_Width - .Width - Rand(10, 30) - 25
        .OrigY = 120
    End With
    With GUI(GuiEnum.GUI_TRAINER)
        .OrigX = Screen_Width - .Width - Rand(10, 30)
        .OrigY = 100 + Rand(10, 40)
    End With
    With GUI(GuiEnum.GUI_INVSTORAGE)
        .OrigX = Rand(10, 30) + 25
        .OrigY = 20 + Rand(10, 40) + 25
    End With
    With GUI(GuiEnum.GUI_POKEMONSTORAGE)
        .OrigX = Rand(10, 30)
        .OrigY = 20 + Rand(10, 40)
    End With
    With GUI(GuiEnum.GUI_CONVO)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = Screen_Height - .Height - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_SHOP)
        .OrigX = Rand(10, 30) + 25
        .OrigY = 20 + Rand(10, 40) + 25
    End With
    With GUI(GuiEnum.GUI_TRADE)
        .OrigX = Rand(10, 30) + 25
        .OrigY = 20 + Rand(10, 40) + 25
    End With
    With GUI(GuiEnum.GUI_POKEDEX)
        .OrigX = Rand(10, 30) + 25
        .OrigY = 20 + Rand(10, 40) + 25
    End With
    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        .OrigX = Rand(10, 30) + 25
        .OrigY = 20 + Rand(10, 40) + 25
    End With
    With GUI(GuiEnum.GUI_RELEARN)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2)
    End With
    With GUI(GuiEnum.GUI_BADGE)
        .OrigX = Screen_Width - .Width - Rand(10, 30)
        .OrigY = 100 + Rand(10, 40)
    End With
    With GUI(GuiEnum.GUI_RANK)
        .OrigX = (Screen_Width - .Width)
        .OrigY = 120
    End With
    With GUI(GuiEnum.GUI_VIRTUALSHOP)
        .OrigX = (Screen_Width / 2) - (.Width / 2)
        .OrigY = (Screen_Height / 2) - (.Height / 2) '+ 100
    End With
    With Button(ButtonEnum.Game_Pokedex)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 6)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_Bag)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 5)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_Card)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 4)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_CheckIn)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 3)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_Rank)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 2)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_VirtualShop)
        .X = Screen_Width - .Width - 10 - ((.Width + 5) * 1)
        .Y = Screen_Height - .Height - 10
    End With
    With Button(ButtonEnum.Game_Menu)
        .X = Screen_Width - .Width - 10
        .Y = Screen_Height - .Height - 10
    End With
    
    With Button(ButtonEnum.Game_Evolve)
        .X = 180
        .Y = 19
    End With
    
    With Button(ButtonEnum.Convo_Reply1)
        .X = (Screen_Width / 2) - (.Width / 2)
        .Y = (Screen_Height / 2) - (111 / 2)
    End With
    With Button(ButtonEnum.Convo_Reply2)
        .X = (Screen_Width / 2) - (.Width / 2)
        .Y = (Screen_Height / 2) - (111 / 2) + 37
    End With
    With Button(ButtonEnum.Convo_Reply3)
        .X = (Screen_Width / 2) - (.Width / 2)
        .Y = (Screen_Height / 2) - (111 / 2) + 74
    End With
End Sub

'//Once we get error in loading gui that means, gui file doesn't exist
Public Sub LoadGui()
Dim i As Long, X As Byte
Dim FileTitle As String
Dim FileName As String

    For i = 1 To GuiEnum.Gui_Count - 1
        Select Case i
            Case GuiEnum.GUI_LOGIN: FileTitle = "window-login"
            Case GuiEnum.GUI_REGISTER: FileTitle = "register-window"
            Case GuiEnum.GUI_CHARACTERSELECT: FileTitle = "character-selection"
            Case GuiEnum.GUI_CHARACTERCREATE: FileTitle = "character-creation"
            Case GuiEnum.GUI_CHOICEBOX: FileTitle = "choice-box"
            Case GuiEnum.GUI_GLOBALMENU: FileTitle = "global-menu"
            Case GuiEnum.GUI_OPTION: FileTitle = "option-window"
            Case GuiEnum.GUI_CHATBOX: FileTitle = "chatbox"
            Case GuiEnum.GUI_INVENTORY: FileTitle = "inventory"
            Case GuiEnum.GUI_INPUTBOX: FileTitle = "input-box"
            Case GuiEnum.GUI_MOVEREPLACE: FileTitle = "move-replace"
            Case GuiEnum.GUI_TRAINER: FileTitle = "trainer"
            Case GuiEnum.GUI_INVSTORAGE: FileTitle = "storage"
            Case GuiEnum.GUI_POKEMONSTORAGE: FileTitle = "storage"
            Case GuiEnum.GUI_CONVO: FileTitle = "convo"
            Case GuiEnum.GUI_SHOP: FileTitle = "shop"
            Case GuiEnum.GUI_TRADE: FileTitle = "trade"
            Case GuiEnum.GUI_POKEDEX: FileTitle = "pokedex"
            Case GuiEnum.GUI_POKEMONSUMMARY: FileTitle = "pokemon-summary"
            Case GuiEnum.GUI_RELEARN: FileTitle = "relearn"
            Case GuiEnum.GUI_BADGE: FileTitle = "badge"
            Case GuiEnum.GUI_RANK: FileTitle = "rank"
            Case GuiEnum.GUI_VIRTUALSHOP: FileTitle = "virtualShop-window"
        End Select
        FileName = App.path & Texture_Path & Trim$(GameSetting.ThemePath) & "\ui\" & FileTitle & ".ini"
        If FileExist(FileName) Then
            With GUI(i)
                .Pic = Val(GetVar(FileName, "GENERAL", "Pic"))
                
                .StartX = Val(GetVar(FileName, "GENERAL", "StartX"))
                .StartY = Val(GetVar(FileName, "GENERAL", "StartY"))
                
                .Width = Val(GetVar(FileName, "GENERAL", "Width"))
                .Height = Val(GetVar(FileName, "GENERAL", "Height"))
            End With
        End If
    Next
    
    FileName = App.path & Texture_Path & Trim$(GameSetting.ThemePath) & "\ui\buttons-setup.ini"
    If FileExist(FileName) Then
        For i = 1 To ButtonEnum.Button_Count - 1
            Select Case i
                Case ButtonEnum.Login_Confirm: FileTitle = "Login_Confirm"
                Case ButtonEnum.Register_Confirm: FileTitle = "Register_Confirm"
                Case ButtonEnum.Register_Close: FileTitle = "Register_Close"
                Case ButtonEnum.Character_SwitchLeft: FileTitle = "CharSel_SwitchLeft"
                Case ButtonEnum.Character_SwitchRight: FileTitle = "CharSel_SwitchRight"
                Case ButtonEnum.Character_New: FileTitle = "CharSel_New"
                Case ButtonEnum.Character_Use: FileTitle = "CharSel_Use"
                Case ButtonEnum.Character_Delete: FileTitle = "CharSel_Delete"
                Case ButtonEnum.CharCreate_Confirm: FileTitle = "CharCreate_Confirm"
                Case ButtonEnum.CharCreate_Close: FileTitle = "CharCreate_Close"
                Case ButtonEnum.ChoiceBox_Yes: FileTitle = "ChoiceBox_Yes"
                Case ButtonEnum.ChoiceBox_No: FileTitle = "ChoiceBox_No"
                Case ButtonEnum.GlobalMenu_Return: FileTitle = "GlobalMenu_Return"
                Case ButtonEnum.GlobalMenu_Option: FileTitle = "GlobalMenu_Option"
                Case ButtonEnum.GlobalMenu_Back: FileTitle = "GlobalMenu_Back"
                Case ButtonEnum.GlobalMenu_Exit: FileTitle = "GlobalMenu_Exit"
                Case ButtonEnum.Option_Close: FileTitle = "Option_Close"
                Case ButtonEnum.Option_Video: FileTitle = "Option_Video"
                Case ButtonEnum.Option_Sound: FileTitle = "Option_Sound"
                Case ButtonEnum.Option_Game: FileTitle = "Option_Game"
                Case ButtonEnum.Option_Control: FileTitle = "Option_Control"
                Case ButtonEnum.Option_cTabUp: FileTitle = "Option_cTabUp"
                Case ButtonEnum.Option_cTabDown: FileTitle = "Option_cTabDown"
                Case ButtonEnum.Option_sMusicUp: FileTitle = "Option_sMusicUp"
                Case ButtonEnum.Option_sMusicDown: FileTitle = "Option_sMusicDown"
                Case ButtonEnum.Option_sSoundUp: FileTitle = "Option_sSoundUp"
                Case ButtonEnum.Option_sSoundDown: FileTitle = "Option_sSoundDown"
                Case ButtonEnum.Chatbox_ScrollUp: FileTitle = "Chatbox_ScrollUp"
                Case ButtonEnum.Chatbox_ScrollDown: FileTitle = "Chatbox_ScrollDown"
                Case ButtonEnum.Chatbox_Minimize: FileTitle = "Chatbox_Minimize"
                Case ButtonEnum.Game_Pokedex: FileTitle = "Game_Pokedex"
                Case ButtonEnum.Game_Bag: FileTitle = "Game_Bag"
                Case ButtonEnum.Game_Card: FileTitle = "Game_Card"
                Case ButtonEnum.Game_CheckIn: FileTitle = "Game_CheckIn"
                Case ButtonEnum.Game_Rank: FileTitle = "Game_Rank"
                Case ButtonEnum.Game_VirtualShop: FileTitle = "Game_VirtualShop"
                Case ButtonEnum.Game_Menu: FileTitle = "Game_Menu"
                Case ButtonEnum.Game_Evolve: FileTitle = "Game_Evolve"
                Case ButtonEnum.Inventory_Close: FileTitle = "Inventory_Close"
                Case ButtonEnum.InputBox_Okay: FileTitle = "InputBox_Okay"
                Case ButtonEnum.InputBox_Cancel: FileTitle = "InputBox_Cancel"
                Case ButtonEnum.MoveReplace_Slot1: FileTitle = "MoveReplace_Slot1"
                Case ButtonEnum.MoveReplace_Slot2: FileTitle = "MoveReplace_Slot2"
                Case ButtonEnum.MoveReplace_Slot3: FileTitle = "MoveReplace_Slot3"
                Case ButtonEnum.MoveReplace_Slot4: FileTitle = "MoveReplace_Slot4"
                Case ButtonEnum.MoveReplace_Cancel: FileTitle = "MoveReplace_Cancel"
                Case ButtonEnum.Trainer_Close: FileTitle = "Trainer_Close"
                Case ButtonEnum.Trainer_Badge: FileTitle = "Trainer_Badge"
                Case ButtonEnum.InvStorage_Close: FileTitle = "InvStorage_Close"
                Case ButtonEnum.InvStorage_Slot1: FileTitle = "InvStorage_Slot1"
                Case ButtonEnum.InvStorage_Slot2: FileTitle = "InvStorage_Slot2"
                Case ButtonEnum.InvStorage_Slot3: FileTitle = "InvStorage_Slot3"
                Case ButtonEnum.InvStorage_Slot4: FileTitle = "InvStorage_Slot4"
                Case ButtonEnum.InvStorage_Slot5: FileTitle = "InvStorage_Slot5"
                Case ButtonEnum.PokemonStorage_Close: FileTitle = "PokemonStorage_Close"
                Case ButtonEnum.PokemonStorage_Slot1: FileTitle = "PokemonStorage_Slot1"
                Case ButtonEnum.PokemonStorage_Slot2: FileTitle = "PokemonStorage_Slot2"
                Case ButtonEnum.PokemonStorage_Slot3: FileTitle = "PokemonStorage_Slot3"
                Case ButtonEnum.PokemonStorage_Slot4: FileTitle = "PokemonStorage_Slot4"
                Case ButtonEnum.PokemonStorage_Slot5: FileTitle = "PokemonStorage_Slot5"
                Case ButtonEnum.Convo_Reply1: FileTitle = "Convo_Reply1"
                Case ButtonEnum.Convo_Reply2: FileTitle = "Convo_Reply2"
                Case ButtonEnum.Convo_Reply3: FileTitle = "Convo_Reply3"
                Case ButtonEnum.Shop_Close: FileTitle = "Shop_Close"
                Case ButtonEnum.Shop_ScrollUp: FileTitle = "Shop_ScrollUp"
                Case ButtonEnum.Shop_ScrollDown: FileTitle = "Shop_ScrollDown"
                Case ButtonEnum.Trade_Close: FileTitle = "Trade_Close"
                Case ButtonEnum.Trade_Accept: FileTitle = "Trade_Accept"
                Case ButtonEnum.Trade_Decline: FileTitle = "Trade_Decline"
                Case ButtonEnum.Trade_Set: FileTitle = "Trade_Set"
                Case ButtonEnum.Trade_AddMoney: FileTitle = "Trade_AddMoney"
                Case ButtonEnum.Pokedex_Close: FileTitle = "Pokedex_Close"
                Case ButtonEnum.Pokedex_ScrollUp: FileTitle = "Pokedex_ScrollUp"
                Case ButtonEnum.Pokedex_ScrollDown: FileTitle = "Pokedex_ScrollDown"
                Case ButtonEnum.PokemonSummary_Close: FileTitle = "PokemonSummary_Close"
                Case ButtonEnum.Relearn_Close: FileTitle = "Relearn_Close"
                Case ButtonEnum.Relearn_ScrollDown: FileTitle = "Relearn_ScrollDown"
                Case ButtonEnum.Relearn_ScrollUp: FileTitle = "Relearn_ScrollUp"
                Case ButtonEnum.Badge_Close: FileTitle = "Badge_Close"
                Case ButtonEnum.Rank_Close: FileTitle = "Rank_Close"
                Case ButtonEnum.Rank_ScrollUp: FileTitle = "Rank_ScrollUp"
                Case ButtonEnum.Rank_ScrollDown: FileTitle = "Rank_ScrollDown"
                Case ButtonEnum.VirtualShop_Close: FileTitle = "VirtualShop_Close"
                Case ButtonEnum.VirtualShop_Buy: FileTitle = "VirtualShop_Buy"
                Case ButtonEnum.VirtualShop_ScrollDown: FileTitle = "VirtualShop_ScrollDown"
                Case ButtonEnum.VirtualShop_ScrollUp: FileTitle = "VirtualShop_ScrollUp"
            End Select
            
            With Button(i)
                For X = ButtonState.StateNormal To ButtonState.StateClick
                    .StartX(X) = Val(GetVar(FileName, FileTitle, "StartX_" & X))
                    .StartY(X) = Val(GetVar(FileName, FileTitle, "StartY_" & X))
                Next
                
                .Width = Val(GetVar(FileName, FileTitle, "Width"))
                .Height = Val(GetVar(FileName, FileTitle, "Height"))
                
                .X = Val(GetVar(FileName, FileTitle, "X"))
                .Y = Val(GetVar(FileName, FileTitle, "Y"))
            End With
        Next
    End If

    ResetGui
    ResetAllGuiLocation
End Sub

' ********************
' ** Map Properties **
' ********************
Public Function CheckRev(ByVal MapNum As Long, ByVal Rev As Long) As Boolean
Dim FileName As String
Dim f As Long
Dim GotRev As Long

    On Error GoTo errorHandler

    FileName = App.path & "\data\cache\maps\map_cache_" & MapNum & ".dat"
    f = FreeFile
    
    If Not FileExist(FileName) Then
        CheckRev = False
        Exit Function
    End If
    
    Open FileName For Binary As #f
        Get #f, , GotRev
    Close #f
    
    If GotRev = Rev Then
        CheckRev = True
    End If
    
    Exit Function
errorHandler:
    CheckRev = False
End Function

Public Sub ClearMap()
    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    'Map.MaxX = MAX_MAPX
    'Map.MaxY = MAX_MAPY
    'ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    Map.Music = "None."
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long, Y As Long
Dim i As Long, a As Byte

    FileName = App.path & "\data\cache\maps\map_cache_" & MapNum & ".dat"
    f = FreeFile
    
    If Not FileExist(FileName) Then
        MsgBox "Failed to load map cache. Exiting...", vbCritical
        UnloadMain
        Exit Sub
    End If
    
    Open FileName For Binary As #f
        With Map
            '//General
            Get #f, , .Revision
            Get #f, , .Name
            Get #f, , .Moral
            
            '//Size
            Get #f, , .MaxX
            Get #f, , .MaxY
            
            '//Redim the size
            If .MaxX < MAX_MAPX Then .MaxX = MAX_MAPX
            If .MaxY < MAX_MAPY Then .MaxY = MAX_MAPY
            ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)
        End With
        
        '//Tiles
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//Layer
                    For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                        For a = MapLayerType.Normal To MapLayerType.Animated
                            Get #f, , .Layer(i, a).Tile
                            Get #f, , .Layer(i, a).TileX
                            Get #f, , .Layer(i, a).TileY
                            '//Map anim
                            Get #f, , .Layer(i, a).MapAnim
                        Next
                    Next
                    '//Tile Data
                    Get #f, , .Attribute
                    Get #f, , .Data1
                    Get #f, , .Data2
                    Get #f, , .Data3
                    Get #f, , .Data4
                End With
            Next
        Next
        
        With Map
            '//Map Link
            Get #f, , .LinkUp
            Get #f, , .LinkDown
            Get #f, , .LinkLeft
            Get #f, , .LinkRight
            
            '//Map Data
            Get #f, , .Music
            
            '//Npc
            For i = 1 To MAX_MAP_NPC
                Get #f, , .Npc(i)
            Next
            
            '//Moral
            Get #f, , .KillPlayer
            Get #f, , .IsCave
            Get #f, , .CaveLight
            Get #f, , .SpriteType
            Get #f, , .StartWeather
            Get #f, , .NoCure
        End With
    Close #f
    DoEvents
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long, Y As Long
Dim i As Long, a As Long

    FileName = App.path & "\data\cache\maps\map_cache_" & MapNum & ".dat"
    f = FreeFile
    
    If FileExist(FileName) Then DeleteFile FileName
    
    Open FileName For Binary As #f
        With Map
            '//General
            Put #f, , .Revision
            Put #f, , .Name
            Put #f, , .Moral
            
            '//Size
            Put #f, , .MaxX
            Put #f, , .MaxY
        End With
        
        '//Tiles
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    '//Layer
                    For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                        For a = MapLayerType.Normal To MapLayerType.Animated
                            Put #f, , .Layer(i, a).Tile
                            Put #f, , .Layer(i, a).TileX
                            Put #f, , .Layer(i, a).TileY
                            '//Map anim
                            Put #f, , .Layer(i, a).MapAnim
                        Next
                    Next
                    
                    '//Tile Data
                    Put #f, , .Attribute
                    Put #f, , .Data1
                    Put #f, , .Data2
                    Put #f, , .Data3
                    Put #f, , .Data4
                End With
            Next
        Next
        
        With Map
            '//Map Link
            Put #f, , .LinkUp
            Put #f, , .LinkDown
            Put #f, , .LinkLeft
            Put #f, , .LinkRight
            
            '//Map Data
            Put #f, , .Music
            
            '//Npc
            For i = 1 To MAX_MAP_NPC
                Put #f, , .Npc(i)
            Next
            
            '//Moral
            Put #f, , .KillPlayer
            Put #f, , .IsCave
            Put #f, , .CaveLight
            Put #f, , .SpriteType
            Put #f, , .StartWeather
            Put #f, , .NoCure
        End With
    Close #f
    DoEvents
End Sub

' ***********************
' ** Player Properties **
' ***********************
Public Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Call ZeroMemory(ByVal VarPtr(PlayerPokemon(Index)), LenB(PlayerPokemon(Index)))
    Player(Index).Name = vbNullString
End Sub

Public Sub ClearPlayers()
Dim i As Long

    For i = 1 To MAX_PLAYER
        ClearPlayer i
    Next
End Sub

' *****************
' ** Control Key **
' *****************
Public Sub ClearControlKey()
Dim i As Long
Dim Control As Long

    '//Set Key to default
    
    Language
    
    '//Movement Key
    With ControlKey(ControlEnum.KeyUp)
        .keyName = TextUIOptionUp
        .cAsciiKey = 38
    End With
    With ControlKey(ControlEnum.KeyDown)
        .keyName = TextUIOptionDown
        .cAsciiKey = 40
    End With
    With ControlKey(ControlEnum.KeyLeft)
        .keyName = TextUIOptionLeft
        .cAsciiKey = 37
    End With
    With ControlKey(ControlEnum.KeyRight)
        .keyName = TextUIOptionRight
        .cAsciiKey = 39
    End With
    With ControlKey(ControlEnum.KeyCheckMove)
        .keyName = TextUIOptionCheckMove
        .cAsciiKey = 17
    End With
    With ControlKey(ControlEnum.KeyMoveUp)
        .keyName = TextUIOptionMoveSlot1
        .cAsciiKey = 87
    End With
    With ControlKey(ControlEnum.KeyMoveDown)
        .keyName = TextUIOptionMoveSlot2
        .cAsciiKey = 83
    End With
    With ControlKey(ControlEnum.KeyMoveLeft)
        .keyName = TextUIOptionMoveSlot3
        .cAsciiKey = 65
    End With
    With ControlKey(ControlEnum.KeyMoveRight)
        .keyName = TextUIOptionMoveSlot4
        .cAsciiKey = 68
    End With
    With ControlKey(ControlEnum.KeyAttack)
        .keyName = TextUIOptionAttack
        .cAsciiKey = 32
    End With
    
    With ControlKey(ControlEnum.KeyPokeSlot1)
        .keyName = TextUIOptionPokeSlot1
        .cAsciiKey = 49
    End With
    With ControlKey(ControlEnum.KeyPokeSlot2)
        .keyName = TextUIOptionPokeSlot2
        .cAsciiKey = 50
    End With
    With ControlKey(ControlEnum.KeyPokeSlot3)
        .keyName = TextUIOptionPokeSlot3
        .cAsciiKey = 51
    End With
    With ControlKey(ControlEnum.KeyPokeSlot4)
        .keyName = TextUIOptionPokeSlot4
        .cAsciiKey = 52
    End With
    With ControlKey(ControlEnum.KeyPokeSlot5)
        .keyName = TextUIOptionPokeSlot5
        .cAsciiKey = 53
    End With
    With ControlKey(ControlEnum.KeyPokeSlot6)
        .keyName = TextUIOptionPokeSlot6
        .cAsciiKey = 54
    End With
    
    With ControlKey(ControlEnum.KeyHotbarSlot1)
        .keyName = TextUIOptionHotbarSlot1
        .cAsciiKey = 112
    End With
    With ControlKey(ControlEnum.KeyHotbarSlot2)
        .keyName = TextUIOptionHotbarSlot2
        .cAsciiKey = 113
    End With
    With ControlKey(ControlEnum.KeyHotbarSlot3)
        .keyName = TextUIOptionHotbarSlot3
        .cAsciiKey = 114
    End With
    With ControlKey(ControlEnum.KeyHotbarSlot4)
        .keyName = TextUIOptionHotbarSlot4
        .cAsciiKey = 115
    End With
    With ControlKey(ControlEnum.KeyHotbarSlot5)
        .keyName = TextUIOptionHotbarSlot5
        .cAsciiKey = 116
    End With
    
    With ControlKey(ControlEnum.KeyInventory)
        .keyName = TextUIOptionInventory
        .cAsciiKey = 73
    End With
    With ControlKey(ControlEnum.KeyPokedex)
        .keyName = TextUIOptionPokedex
        .cAsciiKey = 80
    End With
    With ControlKey(ControlEnum.KeyInteract)
        .keyName = TextUIOptionInteract
        .cAsciiKey = 32
    End With
    
    With ControlKey(ControlEnum.KeyConvo1)
        .keyName = TextUIOptionConvoChoice1
        .cAsciiKey = 90
    End With
    With ControlKey(ControlEnum.KeyConvo2)
        .keyName = TextUIOptionConvoChoice2
        .cAsciiKey = 88
    End With
    With ControlKey(ControlEnum.KeyConvo3)
        .keyName = TextUIOptionConvoChoice3
        .cAsciiKey = 67
    End With
    With ControlKey(ControlEnum.KeyConvo4)
        .keyName = TextUIOptionConvoChoice4
        .cAsciiKey = 86
    End With
End Sub

Public Sub LoadControlKey()
Dim FileName As String
Dim i As Long

    On Error GoTo errorHandler

    FileName = App.path & "\data\config\controlkey.ini"
    
    If Not FileExist(FileName) Then
        Call ClearControlKey
        Call SaveControlKey
        Exit Sub
    End If
    
    For i = 1 To ControlEnum.Control_Count - 1
        With ControlKey(i)
            .cAsciiKey = GetVar(FileName, Trim$(.keyName), "KeyAscii")
        End With
    Next
    
    Exit Sub
errorHandler:
    Call ClearControlKey
    Call SaveControlKey
    Exit Sub
End Sub

Public Sub SaveControlKey()
Dim FileName As String
Dim i As Long

    FileName = App.path & "\data\config\controlkey.ini"
    
    If FileExist(FileName) Then
        DeleteFile FileName
    End If
    
    For i = 1 To ControlEnum.Control_Count - 1
        With ControlKey(i)
            Call PutVar(FileName, Trim$(.keyName), "KeyAscii", Str(.cAsciiKey))
        End With
    Next
End Sub

' *********
' ** Npc **
' *********
Public Sub ClearNpc(ByVal NpcNum As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(NpcNum)), LenB(Npc(NpcNum)))
    Npc(NpcNum).Name = vbNullString
End Sub

Public Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPC
        ClearNpc i
    Next
End Sub

' *************
' ** Pokemon **
' *************
Public Sub ClearPokemon(ByVal PokemonNum As Long)
    Call ZeroMemory(ByVal VarPtr(Pokemon(PokemonNum)), LenB(Pokemon(PokemonNum)))
    Pokemon(PokemonNum).Name = vbNullString
    Pokemon(PokemonNum).Species = vbNullString
    Pokemon(PokemonNum).PokeDexEntry = vbNullString
    Pokemon(PokemonNum).Sound = "None."
End Sub

Public Sub ClearPokemons()
Dim i As Long

    For i = 1 To MAX_POKEMON
        ClearPokemon i
    Next
End Sub

' *************
' ** Map Npc **
' *************
Public Sub ClearMapNpc(ByVal MapNpcNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNpcNum)), LenB(MapNpc(MapNpcNum)))
End Sub

Public Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPC
        ClearMapNpc i
    Next
End Sub

Public Sub ClearMapNpcPokemons()
Dim i As Long

    For i = 1 To MAX_MAP_NPC
        Call ZeroMemory(ByVal VarPtr(MapNpcPokemon(i)), LenB(MapNpcPokemon(i)))
    Next
End Sub

' *****************
' ** Map Pokemon **
' *****************
Public Sub ClearMapPokemon(ByVal MapPokeNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapPokemon(MapPokeNum)), LenB(MapPokemon(MapPokeNum)))
End Sub

Public Sub ClearMapPokemons()
Dim i As Long

    For i = 1 To Pokemon_HighIndex
        ClearMapPokemon i
    Next
End Sub

Public Sub ClearPlayerPokemon(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PlayerPokemon(Index)), LenB(PlayerPokemon(Index)))
End Sub

Public Sub ClearPlayerPokemons()
Dim i As Long

    For i = 1 To Player_HighIndex
        ClearPlayerPokemon i
    Next
End Sub

Public Sub ClearNpcPokemon(ByVal MapNpcNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpcPokemon(MapNpcNum)), LenB(MapNpcPokemon(MapNpcNum)))
End Sub

Public Sub ClearNpcPokemons()
Dim i As Long

    For i = 1 To MAX_MAP_NPC
        ClearNpcPokemon i
    Next
End Sub

' **********
' ** Item **
' **********
Public Sub ClearItem(ByVal ItemNum As Long)
    Call ZeroMemory(ByVal VarPtr(Item(ItemNum)), LenB(Item(ItemNum)))
    Item(ItemNum).Name = vbNullString
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEM
        ClearItem i
    Next
End Sub

' *****************
' ** PokemonMove **
' *****************
Public Sub ClearPokemonMove(ByVal PokemonMoveNum As Long)
    Call ZeroMemory(ByVal VarPtr(PokemonMove(PokemonMoveNum)), LenB(PokemonMove(PokemonMoveNum)))
    PokemonMove(PokemonMoveNum).Name = vbNullString
    PokemonMove(PokemonMoveNum).Sound = "None."
End Sub

Public Sub ClearPokemonMoves()
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        ClearPokemonMove i
    Next
End Sub

' ***************
' ** Animation **
' ***************
Public Sub ClearAnimation(ByVal AnimationNum As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(AnimationNum)), LenB(Animation(AnimationNum)))
    Animation(AnimationNum).Name = vbNullString
End Sub

Public Sub ClearAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATION
        ClearAnimation i
    Next
End Sub

' ***************
' ** Spawn **
' ***************
Public Sub ClearSpawn(ByVal SpawnNum As Long)
    Call ZeroMemory(ByVal VarPtr(Spawn(SpawnNum)), LenB(Spawn(SpawnNum)))
    Spawn(SpawnNum).SpawnTimeMax = 23
End Sub

Public Sub ClearSpawns()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        ClearSpawn i
    Next
End Sub

' ***************
' ** Conversation **
' ***************
Public Sub ClearConversation(ByVal ConversationNum As Long)
Dim X As Byte, Y As Byte, z As Byte

    Call ZeroMemory(ByVal VarPtr(Conversation(ConversationNum)), LenB(Conversation(ConversationNum)))
    For X = 1 To MAX_CONV_DATA
        For Y = 1 To MAX_LANGUAGE
            Conversation(ConversationNum).ConvData(X).TextLang(Y).Text = vbNullString
            For z = 1 To 3
                Conversation(ConversationNum).ConvData(X).TextLang(Y).tReply(z) = vbNullString
            Next
        Next
    Next
End Sub

Public Sub ClearConversations()
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        ClearConversation i
    Next
End Sub

' ***************
' ** Shop **
' ***************
Public Sub ClearShop(ByVal ShopNum As Long)
Dim X As Byte, Y As Byte, z As Byte

    Call ZeroMemory(ByVal VarPtr(Shop(ShopNum)), LenB(Shop(ShopNum)))
End Sub

Public Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOP
        ClearShop i
    Next
End Sub

' ***************
' ** Quest **
' ***************
Public Sub ClearQuest(ByVal QuestNum As Long)
Dim X As Byte, Y As Byte, z As Byte

    Call ZeroMemory(ByVal VarPtr(Quest(QuestNum)), LenB(Quest(QuestNum)))
End Sub

Public Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUEST
        ClearQuest i
    Next
End Sub

' ****************
' ** Chatbubble **
' ****************
Public Sub ClearChatbubble()
Dim i As Long

    For i = 1 To 255
        Call ZeroMemory(ByVal VarPtr(chatBubble(i)), LenB(chatBubble(i)))
        chatBubble(i).Msg = vbNullString
    Next
End Sub

' *************
' ** SelMenu **
' *************
Public Sub ClearSelMenu()
    Call ZeroMemory(ByVal VarPtr(SelMenu), LenB(SelMenu))
    SelMenu.Visible = False
End Sub

'//Animation
Public Sub ClearAnimInstance(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
End Sub

' *************
' ** SelMenu **
' *************
Public Sub LoadCredits()
Dim tmpCreditText As String
Dim CreditText() As String
Dim i As Long
    
    ' Load on file
    'If Not FileExist(App.Path & "\config\credits.txt") Then
    '    Open App.Path & "\config\credits.txt" For Output As #1
    '    Close #1
    '    tmpCreditText = vbNullString
    'Else
    '    Open App.Path & "\config\credits.txt" For Input As #1
    '        Line Input #1, tmpCreditText
    '    Close #1
    'End If
    '//Static
    tmpCreditText = "#h PokeReborn Team//#h Owner/Philips//#h Head Programmer/Leahos, Dragonick//#h Moderators/Flares///#h Special Thanks/xxxx/xxxx"
    
    ' Split Text
    CreditText() = Split(tmpCreditText, "/")
    CreditTextCount = UBound(CreditText)
    ReDim Credit(CreditTextCount)
    
    For i = 0 To CreditTextCount
        Credit(i).Text = CreditText(i)
        Credit(i).Y = (Screen_Height - 40) + (20 * i)
        Credit(i).StartY = (Screen_Height - 40) + (20 * i)
    Next
End Sub

'//Rank
Public Sub ClearRank()
    Dim i As Long
    
    For i = 1 To UBound(Rank)
        Call ZeroMemory(ByVal VarPtr(Rank(i)), LenB(Rank(i)))
    Next i
End Sub

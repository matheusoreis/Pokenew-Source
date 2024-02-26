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
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, file)
End Sub

'//This check the directory if exist, if not, then create one
Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' Checking of Directory Exist, Create if not
    If Not LCase$(Dir(tDir & tName, vbDirectory)) = tName Then Call MkDir(tDir & tName)
End Sub

'//This check if the file exist
Public Function FileExist(ByVal filename As String) As Boolean
    ' Checking if File Exist
    If LenB(Dir(filename)) > 0 Then FileExist = True
End Function

'//This sub delete the file, if it doesn't exist then it will ignore it
Public Sub DeleteFile(ByVal filename As String)
    On Error Resume Next
    Kill filename
End Sub

Public Sub AddLog(ByVal sString As String)
Dim filename As String

    '//Write error on logs
    filename = App.Path & "\data\logs\server_log_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".txt"
    If Not FileExist(filename) Then
        Open filename For Output As #1
        Close #1
    End If
    Open filename For Append As #1
        Print #1, "[" & KeepTwoDigit(hour(Now)) & ":" & KeepTwoDigit(Minute(Now)) & "]" & " " & sString
    Close #1
End Sub

Public Sub AddIPLog(ByVal sString As String)
Dim filename As String

    '//Write error on logs
    ChkDir App.Path & "\data\", "iplogs"
    filename = App.Path & "\data\iplogs\log_" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".txt"
    If Not FileExist(filename) Then
        Open filename For Output As #1
        Close #1
    End If
    Open filename For Append As #1
        Print #1, "[" & KeepTwoDigit(hour(Now)) & ":" & KeepTwoDigit(Minute(Now)) & "]" & " " & sString
    Close #1
End Sub

' *************
' ** Options **
' *************
Public Sub LoadOption()
Dim filename As String

    filename = App.Path & "\data\option.ini"
    
    If Not FileExist(filename) Then
        With Options
            '//Network
            .Port = GAME_PORT
            
            '//Debug Mode
            .DebugMode = NO
            
            '//Starting Location
            .StartMap = START_MAP
            .startX = START_X
            .startY = START_Y
            .StartDir = DIR_UP
            .ExpRate = 1
            
            .ShinyRarity = 500
            
            .TradeLvlMin = 30
            
            .SameIp = YES
            
            .Rarity = 50
            
            '//MOTD
            .MOTD = "PokeReborn official beta release is near"
            
            Call PutVar(filename, "Network", "Port", Str(.Port))
            Call PutVar(filename, "DebugMode", "DebugMode", Str(.DebugMode))
            Call PutVar(filename, "StartingLocation", "StartMap", Str(.StartMap))
            Call PutVar(filename, "StartingLocation", "StartX", Str(.startX))
            Call PutVar(filename, "StartingLocation", "StartY", Str(.startY))
            Call PutVar(filename, "StartingLocation", "StartDir", Str(.StartDir))
            Call PutVar(filename, "MOTD", "MOTD", Trim$(.MOTD))
            Call PutVar(filename, "Others", "ShinyRarity", Trim$(.ShinyRarity))
            Call PutVar(filename, "Others", "ExpRate", Str(.ExpRate))
            Call PutVar(filename, "Others", "TradeLvlMin", Str(.TradeLvlMin))
            Call PutVar(filename, "Others", "SameIp", Str(.SameIp))
            Call PutVar(filename, "Others", "CheckRarity", Str(.Rarity))
        End With
    Else
        With Options
            '//Network
            .Port = Val(GetVar(filename, "Network", "Port"))
            
            '//Debug Mode
            .DebugMode = Val(GetVar(filename, "DebugMode", "DebugMode"))
            
            '//Starting Location
            .StartMap = Val(GetVar(filename, "StartingLocation", "StartMap"))
            .startX = Val(GetVar(filename, "StartingLocation", "StartX"))
            .startY = Val(GetVar(filename, "StartingLocation", "StartY"))
            .StartDir = Val(GetVar(filename, "StartingLocation", "StartDir"))
            
            '//MOTD
            .MOTD = Trim$(GetVar(filename, "MOTD", "MOTD"))
            
            .ShinyRarity = Val(GetVar(filename, "Others", "ShinyRarity"))
            .ExpRate = Val(GetVar(filename, "Others", "ExpRate"))
            .TradeLvlMin = Val(GetVar(filename, "Others", "TradeLvlMin"))
            .SameIp = Val(GetVar(filename, "Others", "SameIp"))
            .Rarity = Val(GetVar(filename, "Others", "CheckRarity"))
        End With
    End If
End Sub

Public Sub SaveOption()
Dim filename As String

    filename = App.Path & "\data\option.ini"
    
    With Options
        '//Network
        Call PutVar(filename, "Network", "Port", Str(.Port))
        
        '//Debug Mode
        Call PutVar(filename, "Network", "DebugMode", Str(.DebugMode))
        
        '//Starting Location
        Call PutVar(filename, "StartingLocation", "StartMap", Str(.StartMap))
        Call PutVar(filename, "StartingLocation", "StartX", Str(.startX))
        Call PutVar(filename, "StartingLocation", "StartY", Str(.startY))
        Call PutVar(filename, "StartingLocation", "StartDir", Str(.StartDir))
        
        '//MOTD
        Call PutVar(filename, "MOTD", "MOTD", Trim$(.MOTD))
    End With
End Sub

' ***********************
' ** Player Properties **
' ***********************
Public Sub ClearTempPlayer(ByVal Index As Long)
    ' Temp Player Data
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).buffer = New clsBuffer
    TempPlayer(Index).DataTimer = GetTickCount
End Sub

Public Function AccountExist(ByVal User As String) As Boolean
Dim filename As String

    '//Reset
    AccountExist = False
    
    '//Set file destination
    filename = App.Path & "\data\accounts\" & Trim$(User) & "\account.ini"
    
    '//Check if exist
    If FileExist(filename) Then
        AccountExist = True
    End If
End Function

Public Sub AddAccount(ByVal User As String, ByVal Pass As String, ByVal Email As String)
Dim filename As String

    '//Create the file destination folder
    ChkDir App.Path & "\data\accounts\", Trim$(User)

    '//Create account file
    filename = App.Path & "\data\accounts\" & Trim$(User) & "\account.ini"
    Call PutVar(filename, "Account", "Username", User)
    Call PutVar(filename, "Account", "Password", Pass)
    Call PutVar(filename, "Account", "Email", Email)
End Sub

Public Function isPasswordOK(ByVal User As String, ByVal Pass As String) As Boolean
Dim filename As String
Dim Pass2 As String * NAME_LENGTH

    If AccountExist(User) Then
        filename = App.Path & "\data\accounts\" & Trim$(User) & "\account.ini"
        Pass2 = Trim$(GetVar(filename, "Account", "Password"))
        If Trim$(Pass) = Trim$(Pass2) Then isPasswordOK = True
    End If
End Function

Public Function LoadAccount(ByVal Index As Long, ByVal User As String) As Boolean
Dim filename As String

    '//Clear
    Call ZeroMemory(ByVal VarPtr(Account(Index)), LenB(Account(Index)))
    
    '//Create account file
    filename = App.Path & "\data\accounts\" & Trim$(User) & "\account.ini"
    With Account(Index)
        .Username = Trim$(GetVar(filename, "Account", "Username"))
        .Password = Trim$(GetVar(filename, "Account", "Password"))
        .Email = Trim$(GetVar(filename, "Account", "Email"))
    End With
    
    LoadPlayerDatas Index
    
    '//Success
    LoadAccount = True
End Function

Public Sub ClearAccount(ByVal Index As Long)
    '//Clear
    Call ZeroMemory(ByVal VarPtr(Account(Index)), LenB(Account(Index)))
    Account(Index).Username = vbNullString
    Account(Index).Password = vbNullString
    Account(Index).Email = vbNullString
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_PLAYERCHAR
        '//Clear
        Call ZeroMemory(ByVal VarPtr(Player(Index, i)), LenB(Player(Index, i)))
        Call ZeroMemory(ByVal VarPtr(PlayerPokemon(Index)), LenB(PlayerPokemon(Index)))
        Player(Index, i).Name = vbNullString
    Next

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

Public Sub AddPlayerData(ByVal Index As Long, ByVal CharSlot As Byte, ByVal Name As String, ByVal Sprite As Long)
    Dim filename As String
    Dim f As Long

    '//Determine the file location
    'FileName = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & ".ini"

    With Player(Index, CharSlot)
        .Name = Trim$(Name)
        .Sprite = Sprite
        .Access = 0
        .Map = Options.StartMap
        .x = Options.startX
        .Y = Options.startY
        .Dir = Options.StartDir
        .CurHp = 255
        .Money = 3000
        .Muted = 0
        .DidStart = YES
        .Started = CStr(Date)
        .TimePlay = 0
    End With

    'Call PutVar(FileName, "General", "Name", Trim$(Name))
    'Call PutVar(FileName, "General", "Sprite", Str(Sprite))
    'Call PutVar(FileName, "General", "Access", Str(ACCESS_NONE))
    'Call PutVar(FileName, "Location", "Map", Str(Options.StartMap))
    'Call PutVar(FileName, "Location", "X", Str(Options.startX))
    'Call PutVar(FileName, "Location", "Y", Str(Options.startY))
    'Call PutVar(FileName, "Location", "Dir", Str(Options.StartDir))
    'Call PutVar(FileName, "Vital", "CurHP", Str(255))
    'Call PutVar(FileName, "GameData", "Money", Str(3000))
    'Call PutVar(FileName, "Others", "Muted", "0")

    '// For tutorial mode
    'Call PutVar(FileName, "Tutorial", "DidStart", YES)

    SavePlayerData Index, CharSlot

    '//Append name to file
    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Append As #f
    Print #f, Trim$(Name)
    Close #f
End Sub

Public Sub DeletePlayerData(ByVal Index As Long, ByVal CharSlot As Byte)
    '//Make sure data in used
    If Len(Player(Index, CharSlot).Name) <= 0 Then Exit Sub
    
    DeleteName Trim$(Player(Index, CharSlot).Name)
    '//Clear data
    Call ZeroMemory(ByVal VarPtr(Player(Index, CharSlot)), LenB(Player(Index, CharSlot)))
    Player(Index, CharSlot).Name = vbNullString
    
    '//Delete file
    Call DeleteFile(App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & ".ini")
    Call DeleteFile(App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_inv.ini")
    Call DeleteFile(App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokemon.ini")
End Sub

Public Sub SavePlayerData(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & ".ini"

    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , Player(Index, CharSlot)
    Close #f
    
    Debug.Print Player(Index, CharSlot).Name & " [Save]"
    
    'With Player(Index, CharSlot)
    '    Call PutVar(FileName, "General", "Name", Trim$(.Name))
    '    Call PutVar(FileName, "General", "Sprite", Str(.Sprite))
    '    Call PutVar(FileName, "General", "Access", Str(.Access))
    '    Call PutVar(FileName, "Location", "Map", Str(.Map))
    '    Call PutVar(FileName, "Location", "X", Str(.x))
    '    Call PutVar(FileName, "Location", "Y", Str(.y))
    '    Call PutVar(FileName, "Location", "Dir", Str(.Dir))
    '    Call PutVar(FileName, "Vital", "CurHP", Str(.CurHP))
    '    Call PutVar(FileName, "GameData", "Money", Str(.Money))
    '    Call PutVar(FileName, "Others", "Muted", Str(.Muted))
    '    For i = 1 To MAX_NPC
    '        Call PutVar(FileName, "NpcBattledMonth", "NPC#" & i, Str(.NpcBattledMonth(i)))
    '        Call PutVar(FileName, "NpcBattledDay", "NPC#" & i, Str(.NpcBattledDay(i)))
    '    Next
    '    Call PutVar(FileName, "Checkpoint", "Map", Str(.CheckMap))
    '    Call PutVar(FileName, "Checkpoint", "X", Str(.CheckX))
    '    Call PutVar(FileName, "Checkpoint", "Y", Str(.CheckY))
    '    Call PutVar(FileName, "Checkpoint", "Dir", Str(.CheckDir))
    '    For i = 1 To MAX_BADGE
    '        Call PutVar(FileName, "Badge", "Badge #" & i, Str(.Badge(i)))
    '    Next
    '    Call PutVar(FileName, "Status", "Level", Str(.Level))
    '    Call PutVar(FileName, "Status", "Exp", Str(.CurExp))
    '
    '    '//For tutorial mode
    '    Call PutVar(FileName, "Tutorial", "DidStart", Str(.DidStart))
    '    For i = 1 To MAX_HOTBAR
    '        Call PutVar(FileName, "Hotbar", "Hotbar#" & i, Str(.Hotbar(i)))
    '    Next
    'End With
End Sub

Public Sub SavePlayerDatas(ByVal Index As Long)
    If TempPlayer(Index).UseChar > 0 Then
        SavePlayerData Index, TempPlayer(Index).UseChar
        SavePlayerInv Index, TempPlayer(Index).UseChar
        SavePlayerPokemons Index, TempPlayer(Index).UseChar
        SavePlayerInvStorage Index, TempPlayer(Index).UseChar
        SavePlayerPokemonStorage Index, TempPlayer(Index).UseChar
        'SavePlayerSwitch Index, TempPlayer(Index).UseChar
        SavePlayerPokedex Index, TempPlayer(Index).UseChar
        DoEvents
    End If
End Sub

Public Sub LoadPlayerData(ByVal Index As Long, ByVal CharSlot As Long)
Dim filename As String
Dim f As Long
Dim i As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & ".ini"
    
    If Not FileExist(filename) Then Exit Sub
    f = FreeFile
    
    Open filename For Binary As #f
        Get #f, , Player(Index, CharSlot)
    Close #f
    
    Debug.Print Player(Index, CharSlot).Name & " [Load]"
    
    'With Player(Index, CharSlot)
    '    .Name = Trim$(GetVar(FileName, "General", "Name"))
    '    .Sprite = Val(GetVar(FileName, "General", "Sprite"))
    '    .Access = Val(GetVar(FileName, "General", "Access"))
    '    .Map = Val(GetVar(FileName, "Location", "Map"))
    '    .x = Val(GetVar(FileName, "Location", "X"))
    '    .y = Val(GetVar(FileName, "Location", "Y"))
    '    .Dir = Val(GetVar(FileName, "Location", "Dir"))
    '    .CurHP = Val(GetVar(FileName, "Vital", "CurHP"))
    '    .Money = Val(GetVar(FileName, "GameData", "Money"))
    '    .Muted = Val(GetVar(FileName, "Others", "Muted"))
    '    For i = 1 To MAX_NPC
    '        .NpcBattledMonth(i) = Val(GetVar(FileName, "NpcBattledMonth", "NPC#" & i))
    '        .NpcBattledDay(i) = Val(GetVar(FileName, "NpcBattledDay", "NPC#" & i))
    '    Next
    '    .CheckMap = Val(GetVar(FileName, "Checkpoint", "Map"))
    '    .CheckX = Val(GetVar(FileName, "Checkpoint", "X"))
    '    .CheckY = Val(GetVar(FileName, "Checkpoint", "Y"))
    '    .CheckDir = Val(GetVar(FileName, "Checkpoint", "Dir"))
    '    For i = 1 To MAX_BADGE
    '        .Badge(i) = Val(GetVar(FileName, "Badge", "Badge #" & i))
    '    Next
    '    .Level = Val(GetVar(FileName, "Status", "Level"))
    '    .CurExp = Val(GetVar(FileName, "Status", "Exp"))
        
        '//For tutorial mode
    '    .DidStart = Val(GetVar(FileName, "Tutorial", "DidStart"))
    '    For i = 1 To MAX_HOTBAR
    '        .Hotbar(i) = Val(GetVar(FileName, "Hotbar", "Hotbar#" & i))
    '    Next
        
        '//Check for error
        If Player(Index, CharSlot).Level <= 0 Then
            Player(Index, CharSlot).Level = 1
            Player(Index, CharSlot).CurExp = 0
            Player(Index, CharSlot).CurHp = GetPlayerHP(Player(Index, CharSlot).Level)
        End If
    'End With
End Sub

Public Sub LoadPlayerDatas(ByVal Index As Long)
Dim i As Long

    '//Clear data first
    Call ClearPlayer(Index)
    
    For i = 1 To MAX_PLAYERCHAR
        LoadPlayerData Index, i
    Next
End Sub

Public Function CheckSameName(ByVal Name As String) As Boolean
Dim filename As String
Dim f As Long
Dim s As String
    
    filename = App.Path & "\data\accounts\charlist.txt"
    f = FreeFile
    
    '//Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist(filename) Then
        Open App.Path & "\data\accounts\charlist.txt" For Output As #f
        Close #f
        CheckSameName = False
        Exit Function
    End If
    
    Open filename For Input As #f
        Do While Not EOF(f)
            Input #f, s
    
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                CheckSameName = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Public Sub DeleteName(ByVal Name As String)
Dim f1 As Long
Dim f2 As Long
Dim s As String

    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    '//Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Not Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop

    Close #f1
    Close #f2
    Call DeleteFile(App.Path & "\data\accounts\chartemp.txt")
End Sub

Public Sub ClearPlayerInv(ByVal Index As Long)
    Dim i As Byte
    Call ZeroMemory(ByVal VarPtr(PlayerInv(Index)), LenB(PlayerInv(Index)))
    
    ' Lock lasts 10 InvSlots and buying by cash.
    For i = (MAX_PLAYER_INV - INV_SLOTS_LOCKED + 1) To MAX_PLAYER_INV
        PlayerInv(Index).Data(i).Locked = YES
    Next i
End Sub

Public Sub LoadPlayerInv(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim i As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_inv.ini"
    
    If Not FileExist(filename) Then
        ClearPlayerInv Index
        SavePlayerInv Index, CharSlot
        Exit Sub
    End If
    
    f = FreeFile
    
    Open filename For Binary As #f
        Get #f, , PlayerInv(Index)
    Close #f
    
    'With PlayerInv(Index)
    '    For i = 1 To MAX_PLAYER_INV
    '        .Data(i).Num = Val(GetVar(FileName, "Inv_Slot_" & i, "Num"))
    '        .Data(i).Value = Val(GetVar(FileName, "Inv_Slot_" & i, "Value"))
    '    Next
    'End With
End Sub

Public Sub SavePlayerInv(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim i As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_inv.ini"
    
    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , PlayerInv(Index)
    Close #f
    
    'With PlayerInv(Index)
    '    For i = 1 To MAX_PLAYER_INV
    '        Call PutVar(FileName, "Inv_Slot_" & i, "Num", Str(.Data(i).Num))
    '        Call PutVar(FileName, "Inv_Slot_" & i, "Value", Str(.Data(i).Value))
    '    Next
    'End With
End Sub

Public Sub ClearPlayerPokemons(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PlayerPokemons(Index)), LenB(PlayerPokemons(Index)))
End Sub

Public Sub LoadPlayerPokemons(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim i As Byte, x As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokemon.ini"
    
    If Not FileExist(filename) Then
        ClearPlayerPokemons Index
        SavePlayerPokemons Index, CharSlot
        Exit Sub
    End If
    
    f = FreeFile
    
    Open filename For Binary As #f
        Get #f, , PlayerPokemons(Index)
    Close #f
    
    'With PlayerPokemons(Index)
    '    For i = 1 To MAX_PLAYER_POKEMON
    '        .Data(i).Num = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Num"))
    '
    '        .Data(i).Level = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Level"))
    '
    '        .Data(i).Nature = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Nature"))
    '
    '        .Data(i).isShiny = Val(GetVar(FileName, "Pokemon_Slot_" & i, "IsShiny"))
    '
    '        .Data(i).Happiness = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Happiness"))
    '
    '        .Data(i).Gender = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Gender"))
    '
    '        .Data(i).Status = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Status"))
            
    '        For x = 1 To StatEnum.Stat_Count - 1
    '            .Data(i).Stat(x).Value = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_val"))
    '            .Data(i).Stat(x).EV = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_EV"))
    '            .Data(i).Stat(x).IV = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_IV"))
    '        Next
            
            '//Vital
    '        .Data(i).CurHP = Val(GetVar(FileName, "Pokemon_Slot_" & i, "CurHP"))
    '        .Data(i).MaxHP = Val(GetVar(FileName, "Pokemon_Slot_" & i, "MaxHP"))
            
            '//Exp
    '        .Data(i).CurExp = Val(GetVar(FileName, "Pokemon_Slot_" & i, "CurExp"))
            
            '//Moveset
    '        For x = 1 To MAX_MOVESET
    '            .Data(i).Moveset(x).Num = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "Num"))
    '            .Data(i).Moveset(x).CurPP = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "CurPP"))
    '            .Data(i).Moveset(x).TotalPP = Val(GetVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "TotalPP"))
    '        Next
            
            '//Ball Used
    '        .Data(i).BallUsed = Val(GetVar(FileName, "Pokemon_Slot_" & i, "BallUsed"))
            
            '//Held Item
    '        .Data(i).HeldItem = Val(GetVar(FileName, "Pokemon_Slot_" & i, "HeldItem"))
    '    Next
    'End With
End Sub

Public Sub SavePlayerPokemons(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim i As Byte, x As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokemon.ini"
    
    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , PlayerPokemons(Index)
    Close #f
    
    'With PlayerPokemons(Index)
    '    For i = 1 To MAX_PLAYER_POKEMON
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Num", Str(.Data(i).Num))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Level", Str(.Data(i).Level))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Nature", Str(.Data(i).Nature))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "IsShiny", Str(.Data(i).isShiny))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Happiness", Str(.Data(i).Happiness))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Gender", Str(.Data(i).Gender))
    '
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "Status", Str(.Data(i).Status))
    '
    '        For x = 1 To StatEnum.Stat_Count - 1
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_val", Str(.Data(i).Stat(x).Value))
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_EV", Str(.Data(i).Stat(x).EV))
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Stat_" & x & "_IV", Str(.Data(i).Stat(x).IV))
    '        Next
    '
    '        '//Vital
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "CurHP", Str(.Data(i).CurHP))
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "MaxHP", Str(.Data(i).MaxHP))
    '
    '        '//Exp
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "CurExp", Str(.Data(i).CurExp))
    '
    '        '//Moveset
    '        For x = 1 To MAX_MOVESET
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "Num", Str(.Data(i).Moveset(x).Num))
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "CurPP", Str(.Data(i).Moveset(x).CurPP))
    '            Call PutVar(FileName, "Pokemon_Slot_" & i, "Moveset_" & x & "TotalPP", Str(.Data(i).Moveset(x).TotalPP))
    '        Next
    '
    '        '//Ball Used
    ''        Call PutVar(FileName, "Pokemon_Slot_" & i, "BallUsed", Str(.Data(i).BallUsed))
    '
            '//Held Item
    '        Call PutVar(FileName, "Pokemon_Slot_" & i, "HeldItem", Str(.Data(i).HeldItem))
     '   Next
    'End With
End Sub

Public Sub ClearPlayerInvStorage(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PlayerInvStorage(Index)), LenB(PlayerInvStorage(Index)))
    PlayerInvStorage(Index).slot(1).Unlocked = YES
    PlayerInvStorage(Index).slot(2).Unlocked = YES
End Sub

Public Sub LoadPlayerInvStorage(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim x As Byte, Y As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_invstorage.dat"
    f = FreeFile
    
    If Not FileExist(filename) Then
        ClearPlayerInvStorage Index
        SavePlayerInvStorage Index, CharSlot
        Exit Sub
    End If
    
    Open filename For Binary As #f
        Get #f, , PlayerInvStorage(Index)
    Close #f
End Sub

Public Sub SavePlayerInvStorage(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim x As Byte, Y As Byte
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_invstorage.dat"
    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , PlayerInvStorage(Index)
    Close #f
End Sub

Public Sub ClearPlayerPokemonStorage(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PlayerPokemonStorage(Index)), LenB(PlayerPokemonStorage(Index)))
    PlayerPokemonStorage(Index).slot(1).Unlocked = YES
    PlayerPokemonStorage(Index).slot(2).Unlocked = YES
End Sub

Public Sub LoadPlayerPokemonStorage(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokemonstorage.dat"
    f = FreeFile
    
    If Not FileExist(filename) Then
        ClearPlayerPokemonStorage Index
        SavePlayerPokemonStorage Index, CharSlot
        Exit Sub
    End If
    
    Open filename For Binary As #f
        Get #f, , PlayerPokemonStorage(Index)
    Close #f
End Sub

Public Sub SavePlayerPokemonStorage(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokemonstorage.dat"
    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , PlayerPokemonStorage(Index)
    Close #f
End Sub

Public Sub ClearPlayerPokedex(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PlayerPokedex(Index)), LenB(PlayerPokedex(Index)))
End Sub

Public Sub LoadPlayerPokedex(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokedex.dat"
    f = FreeFile
    
    If Not FileExist(filename) Then
        ClearPlayerPokedex Index
        SavePlayerPokedex Index, CharSlot
        Exit Sub
    End If
    
    Open filename For Binary As #f
        Get #f, , PlayerPokedex(Index)
    Close #f
End Sub

Public Sub SavePlayerPokedex(ByVal Index As Long, ByVal CharSlot As Byte)
Dim filename As String
Dim f As Long

    '//Determine the file location
    filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\character_slot_" & CharSlot & "_pokedex.dat"
    f = FreeFile
    
    Open filename For Binary As #f
        Put #f, , PlayerPokedex(Index)
    Close #f
End Sub

' ********************
' ** Map Properties **
' ********************
Public Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    Map(MapNum).Music = "None."
End Sub

Public Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAP
        ClearMap i
    Next
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim x As Long, Y As Long
Dim filename As String
Dim f As Long
Dim i As Long
Dim a As Byte

    filename = App.Path & "\data\maps\mapdata_" & MapNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearMap MapNum
        SaveMap MapNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        With Map(MapNum)
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
            ReDim Map(MapNum).Tile(0 To .MaxX, 0 To .MaxY)
        End With
        
        '//Tiles
        For x = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                With Map(MapNum).Tile(x, Y)
                    '//Layer
                    For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                        For a = MapLayerType.Normal To MapLayerType.Animated
                            Get #f, , .Layer(i, a).Tile
                            Get #f, , .Layer(i, a).TileX
                            Get #f, , .Layer(i, a).TileY
                            '//Map Anim
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
        
        With Map(MapNum)
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
    frmServer.Caption = "Loading Map#" & MapNum & ".."
    DoEvents
End Sub

Public Sub LoadMaps()
Dim i As Long

    For i = 1 To MAX_MAP
        LoadMap i
    Next
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim x As Long, Y As Long
Dim filename As String
Dim f As Long
Dim i As Long
Dim a As Byte

    filename = App.Path & "\data\maps\mapdata_" & MapNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        With Map(MapNum)
            '//General
            Put #f, , .Revision
            Put #f, , .Name
            Put #f, , .Moral
            
            '//Size
            Put #f, , .MaxX
            Put #f, , .MaxY
        End With
        
        '//Tiles
        For x = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                With Map(MapNum).Tile(x, Y)
                    '//Layer
                    For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                        For a = MapLayerType.Normal To MapLayerType.Animated
                            Put #f, , .Layer(i, a).Tile
                            Put #f, , .Layer(i, a).TileX
                            Put #f, , .Layer(i, a).TileY
                            '//Map Anim
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
        
        With Map(MapNum)
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

Public Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAP
        SaveMap i
    Next
End Sub

' ********************
' ** Npc Properties **
' ********************
Public Sub ClearNpc(ByVal NpcNum As Long)
    Dim i As Byte
    Call ZeroMemory(ByVal VarPtr(Npc(NpcNum)), LenB(Npc(NpcNum)))
    Npc(NpcNum).Name = vbNullString
    
    For i = 1 To MAX_PLAYER_POKEMON
        Npc(NpcNum).PokemonNature(i) = -1
    Next i
End Sub

Public Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPC
        ClearNpc i
    Next
End Sub

Public Sub LoadNpc(ByVal NpcNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\npcs\npcdata_" & NpcNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearNpc NpcNum
        SaveNpc NpcNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Npc(NpcNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadNpcs()
Dim i As Long

    For i = 1 To MAX_NPC
        LoadNpc i
    Next
End Sub

Public Sub SaveNpc(ByVal NpcNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\npcs\npcdata_" & NpcNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveNpcs()
Dim i As Long

    For i = 1 To MAX_NPC
        SaveNpc i
    Next
End Sub

' ********************
' ** Pokemon Properties **
' ********************
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

Public Sub LoadPokemon(ByVal PokemonNum As Long)
Dim filename As String
Dim f As Long, x As Byte

    filename = App.Path & "\data\pokemons\pokemondata_" & PokemonNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearPokemon PokemonNum
        SavePokemon PokemonNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Pokemon(PokemonNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadPokemons()
Dim i As Long

    For i = 1 To MAX_POKEMON
        LoadPokemon i
    Next
End Sub

Public Sub SavePokemon(ByVal PokemonNum As Long)
Dim filename As String
Dim f As Long, x As Byte

    filename = App.Path & "\data\pokemons\pokemondata_" & PokemonNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Pokemon(PokemonNum)
    Close #f
    DoEvents
End Sub

Public Sub SavePokemons()
Dim i As Long

    For i = 1 To MAX_POKEMON
        SavePokemon i
    Next
End Sub

' ********************
' ** Shop Properties **
' ********************
Public Sub ClearShop(ByVal ShopNum As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(ShopNum)), LenB(Shop(ShopNum)))
End Sub

Public Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOP
        ClearShop i
    Next
End Sub

Public Sub LoadShop(ByVal ShopNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\shop\shopdata_" & ShopNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearShop ShopNum
        SaveShop ShopNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Shop(ShopNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadShops()
Dim i As Long

    For i = 1 To MAX_SHOP
        LoadShop i
    Next
End Sub

Public Sub SaveShop(ByVal ShopNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\shop\shopdata_" & ShopNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOP
        SaveShop i
    Next
End Sub

' ********************
' ** PokemonMove Properties **
' ********************
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

Public Sub LoadPokemonMove(ByVal PokemonMoveNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\moves\movedata_" & PokemonMoveNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearPokemonMove PokemonMoveNum
        SavePokemonMove PokemonMoveNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , PokemonMove(PokemonMoveNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadPokemonMoves()
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        LoadPokemonMove i
    Next
End Sub

Public Sub SavePokemonMove(ByVal PokemonMoveNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\moves\movedata_" & PokemonMoveNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , PokemonMove(PokemonMoveNum)
    Close #f
    DoEvents
End Sub

Public Sub SavePokemonMoves()
Dim i As Long

    For i = 1 To MAX_POKEMON_MOVE
        SavePokemonMove i
    Next
End Sub

' ********************
' ** Animation Properties **
' ********************
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

Public Sub LoadAnimation(ByVal AnimationNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\animations\animationdata_" & AnimationNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearAnimation AnimationNum
        SaveAnimation AnimationNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Animation(AnimationNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATION
        LoadAnimation i
    Next
End Sub

Public Sub SaveAnimation(ByVal AnimationNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\animations\animationdata_" & AnimationNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATION
        SaveAnimation i
    Next
End Sub

' ********************
' ** Spawn Properties **
' ********************
Public Sub ClearSpawn(ByVal SpawnNum As Long)
    Call ZeroMemory(ByVal VarPtr(Spawn(SpawnNum)), LenB(Spawn(SpawnNum)))
    Spawn(SpawnNum).SpawnTimeMax = 23
    Spawn(SpawnNum).Nature = -1
End Sub

Public Sub ClearSpawns()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        ClearSpawn i
    Next
End Sub

Public Sub LoadSpawn(ByVal SpawnNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\mappokemon\spawndata_" & SpawnNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearSpawn SpawnNum
        SaveSpawn SpawnNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Spawn(SpawnNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadSpawns()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        LoadSpawn i
    Next
End Sub

Public Sub SaveSpawn(ByVal SpawnNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\mappokemon\spawndata_" & SpawnNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Spawn(SpawnNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveSpawns()
Dim i As Long

    For i = 1 To MAX_GAME_POKEMON
        SaveSpawn i
    Next
End Sub

' ********************
' ** Conversation Properties **
' ********************
Public Sub ClearConversation(ByVal ConversationNum As Long)
Dim x As Byte, Y As Byte, z As Byte

    Call ZeroMemory(ByVal VarPtr(Conversation(ConversationNum)), LenB(Conversation(ConversationNum)))
    For x = 1 To MAX_CONV_DATA
        For Y = 1 To MAX_LANGUAGE
            Conversation(ConversationNum).ConvData(x).TextLang(Y).Text = vbNullString
            For z = 1 To 3
                Conversation(ConversationNum).ConvData(x).TextLang(Y).tReply(z) = vbNullString
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

Public Sub LoadConversation(ByVal ConversationNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\conversation\conversationdata_" & ConversationNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearConversation ConversationNum
        SaveConversation ConversationNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Conversation(ConversationNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadConversations()
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        LoadConversation i
    Next
End Sub

Public Sub SaveConversation(ByVal ConversationNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\conversation\conversationdata_" & ConversationNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Conversation(ConversationNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveConversations()
Dim i As Long

    For i = 1 To MAX_CONVERSATION
        SaveConversation i
    Next
End Sub

' ********************
' ** Quest Properties **
' ********************
Public Sub ClearQuest(ByVal QuestNum As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(QuestNum)), LenB(Quest(QuestNum)))
End Sub

Public Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUEST
        ClearQuest i
    Next
End Sub

Public Sub LoadQuest(ByVal QuestNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\quest\Questdata_" & QuestNum & ".dat"
    f = FreeFile
    
    '//Check if file exist, if not, create new
    If Not FileExist(filename) Then
        ClearQuest QuestNum
        SaveQuest QuestNum
        Exit Sub
    End If
        
    Open filename For Binary As #f
        Get #f, , Quest(QuestNum)
    Close #f
    DoEvents
End Sub

Public Sub LoadQuests()
Dim i As Long

    For i = 1 To MAX_QUEST
        LoadQuest i
    Next
End Sub

Public Sub SaveQuest(ByVal QuestNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\data\quest\Questdata_" & QuestNum & ".dat"
    If FileExist(filename) Then
        Kill filename
    End If
    f = FreeFile

    Open filename For Binary As #f
        Put #f, , Quest(QuestNum)
    Close #f
    DoEvents
End Sub

Public Sub SaveQuests()
Dim i As Long

    For i = 1 To MAX_QUEST
        SaveQuest i
    Next
End Sub

Public Sub SaveRank()
Dim filename As String, i As Byte

    filename = App.Path & "\data\rank.ini"
    
    For i = 1 To MAX_RANK
        PutVar filename, "RANK", "Name" & i, Trim$(Rank(i).Name)
        PutVar filename, "RANK", "Level" & i, Trim$(Rank(i).Level)
    Next

End Sub

Public Sub LoadRank()
Dim filename As String
Dim i As Byte

    filename = App.Path & "\data\rank.ini"
    
    If Not FileExist(filename) Then Exit Sub
    
    For i = 1 To MAX_RANK
        Rank(i).Name = GetVar(filename, "RANK", "Name" & i)
        Rank(i).Level = Val(GetVar(filename, "RANK", "Level" & i))
    Next
End Sub

Private Sub CheckVirtualShop()
    Dim filename As String, i As Byte, x As Long

    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Select Case i
        Case VirtualShopTabsRec.Skins: filename = App.Path & "\data\virtualshop\skins.ini"
        Case VirtualShopTabsRec.Mounts: filename = App.Path & "\data\virtualshop\mounts.ini"
        Case VirtualShopTabsRec.Items: filename = App.Path & "\data\virtualshop\items.ini"
        Case VirtualShopTabsRec.Vips: filename = App.Path & "\data\virtualshop\vips.ini"
        End Select
        If Not FileExist(filename) Then
            ClearVirtualShop
            Call SaveVirtualShop
            Exit Sub
        End If
    Next i
End Sub

Public Sub SaveVirtualShop()
    Dim filename As String, i As Long, x As Long
    
    
    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Select Case i
        Case VirtualShopTabsRec.Skins: filename = App.Path & "\data\virtualshop\skins.ini"
        Case VirtualShopTabsRec.Mounts: filename = App.Path & "\data\virtualshop\mounts.ini"
        Case VirtualShopTabsRec.Items: filename = App.Path & "\data\virtualshop\items.ini"
        Case VirtualShopTabsRec.Vips: filename = App.Path & "\data\virtualshop\vips.ini"
        End Select
        
        If FileExist(filename) Then
            Kill filename
        End If

        PutVar filename, "INIT", "MAX_SLOTS", CStr(VirtualShop(i).Max_Slots)
        
        For x = 1 To UBound(VirtualShop(i).Items)
            PutVar filename, "Slot" & x, "ItemNum", CStr(VirtualShop(i).Items(x).ItemNum)
            PutVar filename, "Slot" & x, "ItemQuant", CStr(VirtualShop(i).Items(x).ItemQuant)
            PutVar filename, "Slot" & x, "ItemPrice", CStr(VirtualShop(i).Items(x).ItemPrice)
            PutVar filename, "Slot" & x, "CustomDesc", CStr(VirtualShop(i).Items(x).CustomDesc)
        Next x
    Next

End Sub

'
Public Sub LoadVirtualShop()
Dim filename As String
Dim i As Byte
Dim x As Long

    Call CheckVirtualShop

    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Select Case i
        Case VirtualShopTabsRec.Skins: filename = App.Path & "\data\virtualshop\skins.ini"
        Case VirtualShopTabsRec.Mounts: filename = App.Path & "\data\virtualshop\mounts.ini"
        Case VirtualShopTabsRec.Items: filename = App.Path & "\data\virtualshop\items.ini"
        Case VirtualShopTabsRec.Vips: filename = App.Path & "\data\virtualshop\vips.ini"
        End Select
        
        VirtualShop(i).Max_Slots = Val(GetVar(filename, "INIT", "MAX_SLOTS"))
        
        
        ReDim VirtualShop(i).Items(1 To VirtualShop(i).Max_Slots)
        
        If VirtualShop(i).Max_Slots >= 1 Then
            For x = 1 To VirtualShop(i).Max_Slots
                VirtualShop(i).Items(x).ItemNum = Val(Trim$(GetVar(filename, "Slot" & x, "ItemNum")))
                VirtualShop(i).Items(x).ItemQuant = Val(Trim$(GetVar(filename, "Slot" & x, "ItemQuant")))
                VirtualShop(i).Items(x).ItemPrice = Val(Trim$(GetVar(filename, "Slot" & x, "ItemPrice")))
                VirtualShop(i).Items(x).CustomDesc = Val(Trim$(GetVar(filename, "Slot" & x, "CustomDesc")))
            Next x
        End If
    Next i
    
End Sub
'
Public Sub ClearVirtualShop()
    Dim i As Long

    For i = 1 To VirtualShopTabsRec.CountTabs - 1
        Call ZeroMemory(ByVal VarPtr(VirtualShop(i)), LenB(VirtualShop(i)))
        VirtualShop(i).Max_Slots = 1
        ReDim VirtualShop(i).Items(1 To VirtualShop(i).Max_Slots)
    Next i
End Sub

'

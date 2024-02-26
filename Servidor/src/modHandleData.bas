Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CLoginInfo) = GetAddress(AddressOf HandleLoginInfo)
    HandleDataSub(CNewCharacter) = GetAddress(AddressOf HandleNewCharacter)
    HandleDataSub(CUseCharacter) = GetAddress(AddressOf HandleUseCharacter)
    HandleDataSub(CDelCharacter) = GetAddress(AddressOf HandleDelCharacter)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(CGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(CPartyMsg) = GetAddress(AddressOf HandlePartyMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CPlayerPokemonMove) = GetAddress(AddressOf HandlePlayerPokemonMove)
    HandleDataSub(CPlayerPokemonDir) = GetAddress(AddressOf HandlePlayerPokemonDir)
    HandleDataSub(CGetItem) = GetAddress(AddressOf HandleGetItem)
    HandleDataSub(CPlayerPokemonState) = GetAddress(AddressOf HandlePlayerPokemonState)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CChangePassword) = GetAddress(AddressOf HandleChangePassword)
    HandleDataSub(CReplaceNewMove) = GetAddress(AddressOf HandleReplaceNewMove)
    HandleDataSub(CEvolvePoke) = GetAddress(AddressOf HandleEvolvePoke)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CSwitchInvSlot) = GetAddress(AddressOf HandleSwitchInvSlot)
    HandleDataSub(CGotData) = GetAddress(AddressOf HandleGotData)
    HandleDataSub(COpenStorage) = GetAddress(AddressOf HandleOpenStorage)
    HandleDataSub(CDepositItemTo) = GetAddress(AddressOf HandleDepositItemTo)
    HandleDataSub(CSwitchStorageSlot) = GetAddress(AddressOf HandleSwitchStorageSlot)
    HandleDataSub(CSwitchStorageItem) = GetAddress(AddressOf HandleSwitchStorageItem)
    HandleDataSub(CWithdrawItemTo) = GetAddress(AddressOf HandleWithdrawItemTo)
    HandleDataSub(CConvo) = GetAddress(AddressOf HandleConvo)
    HandleDataSub(CProcessConvo) = GetAddress(AddressOf HandleProcessConvo)
    HandleDataSub(CDepositPokemon) = GetAddress(AddressOf HandleDepositPokemon)
    HandleDataSub(CWithdrawPokemon) = GetAddress(AddressOf HandleWithdrawPokemon)
    HandleDataSub(CSwitchStoragePokeSlot) = GetAddress(AddressOf HandleSwitchStoragePokeSlot)
    HandleDataSub(CSwitchStoragePoke) = GetAddress(AddressOf HandleSwitchStoragePoke)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CRequest) = GetAddress(AddressOf HandleRequest)
    HandleDataSub(CRequestState) = GetAddress(AddressOf HandleRequestState)
    HandleDataSub(CAddTrade) = GetAddress(AddressOf HandleAddTrade)
    HandleDataSub(CRemoveTrade) = GetAddress(AddressOf HandleRemoveTrade)
    HandleDataSub(CTradeUpdateMoney) = GetAddress(AddressOf HandleTradeUpdateMoney)
    HandleDataSub(CSetTradeState) = GetAddress(AddressOf HandleSetTradeState)
    HandleDataSub(CTradeState) = GetAddress(AddressOf HandleTradeState)
    HandleDataSub(CScanPokedex) = GetAddress(AddressOf HandleScanPokedex)
    HandleDataSub(CMOTD) = GetAddress(AddressOf HandleMOTD)
    HandleDataSub(CCopyMap) = GetAddress(AddressOf HandleCopyMap)
    HandleDataSub(CReleasePokemon) = GetAddress(AddressOf HandleReleasePokemon)
    HandleDataSub(CGiveItemTo) = GetAddress(AddressOf HandleGiveItemTo)
    HandleDataSub(CGivePokemonTo) = GetAddress(AddressOf HandleGivePokemonTo)
    HandleDataSub(CSpawnPokemon) = GetAddress(AddressOf HandleSpawnPokemon)
    HandleDataSub(CSetLanguage) = GetAddress(AddressOf HandleSetLanguage)
    HandleDataSub(CBuyStorageSlot) = GetAddress(AddressOf HandleBuyStorageSlot)
    HandleDataSub(CSellPokeStorageSlot) = GetAddress(AddressOf HandleSellPokeStorageSlot)
    HandleDataSub(CChangeShinyRate) = GetAddress(AddressOf HandleChangeShinyRate)
    HandleDataSub(CRelearnMove) = GetAddress(AddressOf HandleRelearnMove)
    HandleDataSub(CUseRevive) = GetAddress(AddressOf HandleUseRevive)
    HandleDataSub(CAddHeld) = GetAddress(AddressOf HandleAddHeld)
    HandleDataSub(CRemoveHeld) = GetAddress(AddressOf HandleRemoveHeld)
    HandleDataSub(CStealthMode) = GetAddress(AddressOf HandleStealthMode)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CRequestRank) = GetAddress(AddressOf HandleRequestRank)
    HandleDataSub(CHotbarUpdate) = GetAddress(AddressOf HandleHotbarUpdate)
    HandleDataSub(CUseHotbar) = GetAddress(AddressOf HandleUseHotbar)
    HandleDataSub(CCreateParty) = GetAddress(AddressOf HandleCreateParty)
    HandleDataSub(CLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    '//Editors
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CMap) = GetAddress(AddressOf HandleMap)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CRequestNpc) = GetAddress(AddressOf HandleRequestNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditPokemon) = GetAddress(AddressOf HandleRequestEditPokemon)
    HandleDataSub(CRequestPokemon) = GetAddress(AddressOf HandleRequestPokemon)
    HandleDataSub(CSavePokemon) = GetAddress(AddressOf HandleSavePokemon)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CRequestItem) = GetAddress(AddressOf HandleRequestItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditPokemonMove) = GetAddress(AddressOf HandleRequestEditPokemonMove)
    HandleDataSub(CRequestPokemonMove) = GetAddress(AddressOf HandleRequestPokemonMove)
    HandleDataSub(CSavePokemonMove) = GetAddress(AddressOf HandleSavePokemonMove)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CRequestAnimation) = GetAddress(AddressOf HandleRequestAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestEditSpawn) = GetAddress(AddressOf HandleRequestEditSpawn)
    HandleDataSub(CRequestSpawn) = GetAddress(AddressOf HandleRequestSpawn)
    HandleDataSub(CSaveSpawn) = GetAddress(AddressOf HandleSaveSpawn)
    HandleDataSub(CRequestEditConversation) = GetAddress(AddressOf HandleRequestEditConversation)
    HandleDataSub(CRequestConversation) = GetAddress(AddressOf HandleRequestConversation)
    HandleDataSub(CSaveConversation) = GetAddress(AddressOf HandleSaveConversation)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CRequestShop) = GetAddress(AddressOf HandleRequestShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CRequestQuest) = GetAddress(AddressOf HandleRequestQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CMutePlayer) = GetAddress(AddressOf HandleMutePlayer)
    HandleDataSub(CUnmutePlayer) = GetAddress(AddressOf HandleUnmutePlayer)
    HandleDataSub(CFlyToBadge) = GetAddress(AddressOf HandleFlyToBadge)
    HandleDataSub(CRequestCash) = GetAddress(AddressOf HandleRequestCash)
    HandleDataSub(CSetCash) = GetAddress(AddressOf HandleSetCash)
    HandleDataSub(CRequestServerInfo) = GetAddress(AddressOf HandleRequestServerInfo)
    HandleDataSub(CBuyInvSlot) = GetAddress(AddressOf HandleBuyInvSlot)
    HandleDataSub(CRequestVirtualShop) = GetAddress(AddressOf HandleRequestVirtualShop)
    HandleDataSub(CPurchaseVirtualShop) = GetAddress(AddressOf HandlePurchaseVirtualShop)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
End Sub

Public Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    ' Prevent from receiving a empty data
    If ((Not Data) = -1) Or ((Not Data) = 0) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    If MsgType < 1 Or MsgType >= CMSG_Count Then
        buffer.Flush
        Set buffer = Nothing
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, buffer.ReadBytes(buffer.Length), 0, 0
    
    buffer.Flush
    Set buffer = Nothing
End Sub

Public Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long
Dim RemoveLimit As Byte

    RemoveLimit = NO
    If IsPlaying(Index) And TempPlayer(Index).UseChar > 0 Then
        If Player(Index, TempPlayer(Index).UseChar).Access > ACCESS_NONE Then
            RemoveLimit = YES
        End If
    End If
    
    If RemoveLimit = NO Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
        
        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(Index).DataTimer Then
                Exit Sub
            End If
        End If
    End If

    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData buffer(), vbUnicode, DataLength
    
    ' Prevent from receiving a empty data
    If ((Not buffer) = -1) Or ((Not buffer) = 0) Then Exit Sub
    
    TempPlayer(Index).buffer.WriteBytes buffer()
    
    If TempPlayer(Index).buffer.Length >= 4 Then
        pLength = TempPlayer(Index).buffer.ReadLong(False)
        If pLength < 0 Then Exit Sub
    End If
    Do While pLength > 0 And pLength <= TempPlayer(Index).buffer.Length - 4
        If pLength <= TempPlayer(Index).buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).buffer.ReadLong
            HandleData Index, TempPlayer(Index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(Index).buffer.Length >= 4 Then
            pLength = TempPlayer(Index).buffer.ReadLong(False)
            If pLength < 0 Then Exit Sub
        End If
    Loop

    TempPlayer(Index).buffer.Trim
End Sub

Private Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' Prevent from receiving a empty dat

    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Username As String, Password As String, Email As String
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte
    
    ' Prevent from receiving a empty data
    

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Username = Trim$(buffer.ReadString)
        Password = Trim$(buffer.ReadString)
        Email = Trim$(buffer.ReadString)
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(Index).ProcessorID = Trim$(buffer.ReadString)
        
        '//ToDo: Checked Banned ProcessorID
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Username, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Password, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Invalid email", White
                Case LANG_EN: AddAlert Index, "Invalid email", White
                Case LANG_ES: AddAlert Index, "Invalid email", White
            End Select
            Exit Sub
        End If
        
        If AccountExist(Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Username is already in used", White
                Case LANG_EN: AddAlert Index, "Username is already in used", White
                Case LANG_ES: AddAlert Index, "Username is already in used", White
            End Select
            Exit Sub
        End If
        
        AddAccount Username, Password, Email
        Select Case CurLanguage
            Case LANG_PT: AddAlert Index, "Account created", White
            Case LANG_EN: AddAlert Index, "Account created", White
            Case LANG_ES: AddAlert Index, "Account created", White
        End Select
        TextAdd frmServer.txtLog, "Account '" & Username & "' has been created..."
        AddLog "Account '" & Username & "' has been created"
        
        If Not LoadAccount(Index, Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Failed to load account data, Please contact the developer", White
                Case LANG_EN: AddAlert Index, "Failed to load account data, Please contact the developer", White
                Case LANG_ES: AddAlert Index, "Failed to load account data, Please contact the developer", White
            End Select
            Exit Sub
        End If
        
        TextAdd frmServer.txtLog, "Account '" & Username & "' has logged in..."
        AddIPLog "Account '" & Username & "' has logged in on IP " & GetPlayerIP(Index)
        
        '//Update Characters
        'LoadPlayerDatas Index
        SendDataLimit Index
        SendCharacters Index
        '//Send Connect
        SendLoginOk Index
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleLoginInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Username As String, Password As String
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte
Dim i As Long

    ' Prevent from receiving a empty data
    

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Username = Trim$(buffer.ReadString)
        Password = Trim$(buffer.ReadString)
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(Index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Username, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert Index, "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(Password, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert Index, "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        If Not AccountExist(Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Account does not exist", White
                Case LANG_EN: AddAlert Index, "Account does not exist", White
                Case LANG_ES: AddAlert Index, "Account does not exist", White
            End Select
            Exit Sub
        End If
        
        If Not isPasswordOK(Username, Password) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Invalid password", White
                Case LANG_EN: AddAlert Index, "Invalid password", White
                Case LANG_ES: AddAlert Index, "Invalid password", White
            End Select
            Exit Sub
        End If
        
        i = FindAccount(Username)
        If i > 0 Then
            '//Check User IP
            If Len(Trim$(GetPlayerIP(i))) > 0 Then
                If Trim$(GetPlayerIP(i)) = GetPlayerIP(Index) Then
                    '//Disconnect
                    CloseSocket i
                Else
                    Select Case CurLanguage
                        Case LANG_PT: AddAlert Index, "Account is currently connected", White, YES
                        Case LANG_EN: AddAlert Index, "Account is currently connected", White, YES
                        Case LANG_ES: AddAlert Index, "Account is currently connected", White, YES
                    End Select
                    Exit Sub
                End If
            Else
                '//Disconnect
                CloseSocket i
            End If
        End If
        
        If Not LoadAccount(Index, Username) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Failed to load account data, Please contact the developer", White, YES
                Case LANG_EN: AddAlert Index, "Failed to load account data, Please contact the developer", White, YES
                Case LANG_ES: AddAlert Index, "Failed to load account data, Please contact the developer", White, YES
            End Select
            Exit Sub
        End If
        
        frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
        frmServer.lvwInfo.ListItems(Index).SubItems(2) = Username
        
        TextAdd frmServer.txtLog, "Account '" & Username & "' has logged in..."
        AddIPLog "Account '" & Username & "' has logged in on IP " & GetPlayerIP(Index)
        
        SendDataLimit Index
        '//Update Characters
        SendCharacters Index
        '//Send connect
        SendLoginOk Index
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleNewCharacter(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharName As String, Gender As Byte, CharSelected As Byte
Dim Sprite As Long
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharName = Trim$(buffer.ReadString)
        Gender = buffer.ReadByte
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(Index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        If Not CheckNameInput(CharName, False, (NAME_LENGTH - 1)) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_EN: AddAlert Index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                Case LANG_ES: AddAlert Index, "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
            End Select
            Exit Sub
        End If
        
        '//Check if CharName exist
        If CheckSameName(CharName) Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Character name already exist", White
                Case LANG_EN: AddAlert Index, "Character name already exist", White
                Case LANG_ES: AddAlert Index, "Character name already exist", White
            End Select
            Exit Sub
        End If
        
        '//Set specific gender sprite
        If Gender = GENDER_MALE Then
            Sprite = 1
        ElseIf Gender = GENDER_FEMALE Then
            Sprite = 2
        End If
        
        '//Add Character
        AddPlayerData Index, CharSelected, CharName, Sprite
        '//Load data
        LoadPlayerData Index, CharSelected
        LoadPlayerInv Index, CharSelected
        LoadPlayerPokemons Index, CharSelected
        LoadPlayerInvStorage Index, CharSelected
        LoadPlayerPokemonStorage Index, CharSelected
        LoadPlayerPokedex Index, CharSelected
        
        TextAdd frmServer.txtLog, "Character '" & CharName & "' has been created..."
        AddLog "Character '" & CharName & "' has been created from Account '" & Trim$(Account(Index).Username)
        
        Select Case CurLanguage
            Case LANG_PT: AddAlert Index, "New Character Created", White
            Case LANG_EN: AddAlert Index, "New Character Created", White
            Case LANG_ES: AddAlert Index, "New Character Created", White
        End Select
        
        '//Update Characters
        SendCharacters Index
        '//Send connection again
        SendLoginOk Index, YES
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleUseCharacter(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharSelected As Byte
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(Index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        '//Make sure player have character on the selected slot
        If CharSelected > 0 And CharSelected <= MAX_PLAYERCHAR Then
            '//Check if in used
            If Len(Player(Index, CharSelected).Name) <= 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_EN: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_ES: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                End Select
                Exit Sub
            End If
            
            '//Make sure it's not already playing
            If TempPlayer(Index).UseChar > 0 Then
                LeftGame Index
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, "Your character was disconnected, Please try again", White, YES
                    Case LANG_EN: AddAlert Index, "Your character was disconnected, Please try again", White, YES
                    Case LANG_ES: AddAlert Index, "Your character was disconnected, Please try again", White, YES
                End Select
                Exit Sub
            End If
            
            '//Make sure it's not already playing
            If FindPlayer(Trim$(Player(Index, CharSelected).Name)) > 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, "Character is currently connnected", White, YES
                    Case LANG_EN: AddAlert Index, "Character is currently connnected", White, YES
                    Case LANG_ES: AddAlert Index, "Character is currently connnected", White, YES
                End Select
                Exit Sub
            End If
            
            If IsCharacterBanned(Trim$(Player(Index, CharSelected).Name)) Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, "This character is banned", White, YES
                    Case LANG_EN: AddAlert Index, "This character is banned", White, YES
                    Case LANG_ES: AddAlert Index, "This character is banned", White, YES
                End Select
                Exit Sub
            End If
            
            
            '//Set use char
            TempPlayer(Index).UseChar = CharSelected
            JoinGame Index, CurLanguage
        Else
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Invalid Character Slot", White
                Case LANG_EN: AddAlert Index, "Invalid Character Slot", White
                Case LANG_ES: AddAlert Index, "Invalid Character Slot", White
            End Select
            Exit Sub
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleDelCharacter(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim CharSelected As Byte
Dim Major As Long, Minor As Long, Revision As Long
Dim CurLanguage As Byte

    ' Prevent from receiving a empty data
    

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        CharSelected = buffer.ReadByte
        '//Version
        CurLanguage = buffer.ReadByte
        Major = buffer.ReadLong
        Minor = buffer.ReadLong
        Revision = buffer.ReadLong
        TempPlayer(Index).ProcessorID = Trim$(buffer.ReadString)
        
        '//Version
        If Not Major = App.Major Or Not Minor = App.Minor Or Not Revision = App.Revision Then
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_EN: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
                Case LANG_ES: AddAlert Index, "Outdated version of client, Please update your client or download the latest version on site", White
            End Select
            Exit Sub
        End If
        
        '//Make sure player have character on the selected slot
        If CharSelected > 0 And CharSelected <= MAX_PLAYERCHAR Then
            '//Check if in used
            If Len(Player(Index, CharSelected).Name) <= 0 Then
                Select Case CurLanguage
                    Case LANG_PT: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_EN: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                    Case LANG_ES: AddAlert Index, "Character is empty! based on the server, please try to login again. If it didn't work, try to contact the developers", White
                End Select
                Exit Sub
            End If
            
            '//Delete char
            DeletePlayerData Index, CharSelected
            SendCharacters Index
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Character data deleted!", White
                Case LANG_EN: AddAlert Index, "Character data deleted!", White
                Case LANG_ES: AddAlert Index, "Character data deleted!", White
            End Select
        Else
            Select Case CurLanguage
                Case LANG_PT: AddAlert Index, "Invalid Character Slot", White
                Case LANG_EN: AddAlert Index, "Invalid Character Slot", White
                Case LANG_ES: AddAlert Index, "Invalid Character Slot", White
            End Select
            Exit Sub
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim NeedMap As Byte
Dim MapNum As Long
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NeedMap = buffer.ReadByte
    Set buffer = Nothing
    
    If NeedMap = YES Then
        SendMap Index, Player(Index, TempPlayer(Index).UseChar).Map
    End If
    
    '//Send Map Data
    SendMapNpcData Player(Index, TempPlayer(Index).UseChar).Map, Index
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(Index, TempPlayer(Index).UseChar).Map, i, NO, 0, 0, 0, Index
    Next
    
    '//Done loading
    SendJoinMap Index
    
    TempPlayer(Index).GettingMap = False
    
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    If MapNum > 0 Then
        ChangeTempSprite Index, Map(MapNum).SpriteType
    End If
    
    '//Send Weather
    SendWeatherTo Index, MapNum
    
    '//Done Loading
    SendMapDone Index
    
    SendClientTimeTo Index
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
Dim tmpX As Long, tmpY As Long
Dim RndNum As Long

    ' Prevent from receiving a empty data
    

    If Not IsPlaying(Index) Then Exit Sub

    '//Check if can move
    If TempPlayer(Index).GettingMap Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    '//ToFast
    If Player(Index, TempPlayer(Index).UseChar).MoveTmr > GetTickCount Then
        '//ToDo: Create something to prevent hacking
        'SendPlayerXY index, True
        'Exit Sub
    End If
    
    If TempPlayer(Index).StorageType > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).InShop > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).PlayerRequest > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).InDuel > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).InTrade > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).CurConvoNum > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If TempPlayer(Index).InNpcDuel > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If Player(Index, TempPlayer(Index).UseChar).Action > 0 Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    
    '//Desynced
    If Not Player(Index, TempPlayer(Index).UseChar).X = tmpX Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    If Not Player(Index, TempPlayer(Index).UseChar).Y = tmpY Then
        SendPlayerXY Index, True
        Exit Sub
    End If
    
    If Player(Index, TempPlayer(Index).UseChar).IsConfuse = YES Then
        'Dir = Random(0, 3)
        'If Dir < 0 Then Dir = 0
        'If Dir > DIR_RIGHT Then Dir = DIR_RIGHT
        RndNum = Random(1, 10)
        If RndNum = 1 Then
            Player(Index, TempPlayer(Index).UseChar).IsConfuse = 0
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You snapped out of confusion", White
                Case LANG_EN: AddAlert Index, "You snapped out of confusion", White
                Case LANG_ES: AddAlert Index, "You snapped out of confusion", White
            End Select
            SendPlayerStatus Index
        End If
    End If

    Player(Index, TempPlayer(Index).UseChar).MoveTmr = GetTickCount + 200
    Call PlayerMove(Index, Dir)
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
    
    ' Prevent from receiving a empty data
    

    If Not IsPlaying(Index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing
    
    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    
    Player(Index, TempPlayer(Index).UseChar).Dir = Dir
    
    SendPlayerDir Index, True
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long, i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String

    If Not IsPlaying(Index) Then Exit Sub
    
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Muted = YES Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    With Player(Index, TempPlayer(Index).UseChar)
        MapNum = .Map
        
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: UpdateMsg = "[Mapa]" & Trim$(.Name) & ": " & Msg
            Case LANG_EN: UpdateMsg = "[Map]" & Trim$(.Name) & ": " & Msg
            Case LANG_ES: UpdateMsg = "[Mapa]" & Trim$(.Name) & ": " & Msg
        End Select
                
        MsgColor = White
    End With
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If Player(i, TempPlayer(i).UseChar).Map = MapNum Then
                    '//Send Msg
                    SendChatbubble MapNum, Index, TARGET_TYPE_PLAYER, Msg, DarkGrey
                    SendPlayerMsg i, UpdateMsg, MsgColor
                End If
            End If
        End If
    Next
    
    AddLog UpdateMsg
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String

    ' Prevent from receiving a empty data
    

    If Not IsPlaying(Index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Muted = YES Then Exit Sub
    
    With Player(Index, TempPlayer(Index).UseChar)
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: UpdateMsg = "[Todos] " & Trim$(.Name) & ": " & Msg
            Case LANG_EN: UpdateMsg = "[All] " & Trim$(.Name) & ": " & Msg
            Case LANG_ES: UpdateMsg = "[Todos] " & Trim$(.Name) & ": " & Msg
        End Select
        MsgColor = Yellow
    End With
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                If TempPlayer(i).CurLanguage = TempPlayer(Index).CurLanguage Then
                    '//Send Msg
                    SendPlayerMsg i, UpdateMsg, MsgColor
                End If
            End If
        End If
    Next
    
    AddLog UpdateMsg
End Sub

Private Sub HandlePartyMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim Msg As String, MsgColor As Long
    Dim UpdateMsg As String

    ' Prevent from receiving a empty data
    If Not IsPlaying(Index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Set buffer = Nothing

    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Muted = YES Then Exit Sub

    MsgColor = Magenta

    If TempPlayer(Index).InParty > 0 Then

        With TempPlayer(Index)
            For i = 1 To MAX_PARTY
                If .PartyIndex(i) > 0 Then
                    If IsPlaying(.PartyIndex(i)) Then
                        If TempPlayer(.PartyIndex(i)).UseChar > 0 Then
                            Select Case TempPlayer(.PartyIndex(i)).CurLanguage
                            Case LANG_PT: UpdateMsg = "[Grupo] " & Trim$(Player(i, TempPlayer(Index).UseChar).Name) & ": " & Msg
                            Case LANG_EN: UpdateMsg = "[Party] " & Trim$(Player(i, TempPlayer(Index).UseChar).Name) & ": " & Msg
                            Case LANG_ES: UpdateMsg = "[Party] " & Trim$(Player(i, TempPlayer(Index).UseChar).Name) & ": " & Msg
                            End Select

                            '//Send Msg
                            SendPlayerMsg .PartyIndex(i), UpdateMsg, MsgColor

                        End If
                    End If
                End If
            Next
        End With
    Else
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Você precisa estar em grupo!", White
        Case LANG_EN: AddAlert Index, "Você precisa estar em grupo!", White
        Case LANG_ES: AddAlert Index, "Você precisa estar em grupo!", White
        End Select
    End If
    
    AddLog UpdateMsg
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long
Dim Msg As String, MsgColor As Long
Dim UpdateMsg As String

    If Not IsPlaying(Index) Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Muted = YES Then Exit Sub
    
    i = FindPlayer(InputName)
    
    If i > 0 Then
        If Not i = Index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(Index, TempPlayer(Index).UseChar)
                        UpdateMsg = "[" & Trim$(Player(i, TempPlayer(i).UseChar).Name) & "] " & Trim$(.Name) & ": " & Msg
                        MsgColor = BrightCyan
                    End With
    
                    '//Send Msg
                    SendPlayerMsg i, UpdateMsg, MsgColor
                    SendPlayerMsg Index, UpdateMsg, MsgColor
                    
                    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " whisper " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & " [" & Msg & "]"
                Else
                    SendPlayerMsg Index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg Index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg Index, "You cannot message yourself!", BrightRed
        End If
    Else
        SendPlayerMsg Index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    With Player(Index, TempPlayer(Index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
    
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        MapNum = buffer.ReadLong
        Set buffer = Nothing
        
        If MapNum <= 0 Or MapNum > MAX_MAP Then Exit Sub
        
        TextAdd frmServer.txtLog, Trim$(.Name) & " warped to map#" & MapNum
        AddLog Trim$(.Name) & " warped to map#" & MapNum
        
        PlayerWarp Index, MapNum, .X, .Y, .Dir
    End With
End Sub

Private Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim X As Long, Y As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    With Player(Index, TempPlayer(Index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
    
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        X = buffer.ReadLong
        Y = buffer.ReadByte
        Set buffer = Nothing
        
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        If X > Map(.Map).MaxX Then X = Map(.Map).MaxX
        If Y > Map(.Map).MaxY Then Y = Map(.Map).MaxY
        
        '//Set
        .X = X
        .Y = Y
        
        SendPlayerXY Index, True
    End With
End Sub

Private Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    With Player(Index, TempPlayer(Index).UseChar)
        If .Access < ACCESS_MODERATOR Then Exit Sub
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        InputName = buffer.ReadString
        Set buffer = Nothing
    
        i = FindPlayer(InputName)
    
        If i > 0 Then
            If Not i = Index Then
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        PlayerWarp i, .Map, .X, .Y, .Dir
                        
                        AddLog Trim$(Player(i, TempPlayer(i).UseChar).Name) & " warped to [" & .Map & " | " & .X & " | " & .Y & "]"
                    Else
                        SendPlayerMsg Index, "Player is not online!", BrightRed
                    End If
                Else
                    SendPlayerMsg Index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg Index, "You cannot warp yourself to yourself!", BrightRed
            End If
        Else
            SendPlayerMsg Index, "Player is not online!", BrightRed
        End If
    End With
End Sub

Private Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    i = FindPlayer(InputName)
        
    If i > 0 Then
        If Not i = Index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(i, TempPlayer(i).UseChar)
                        PlayerWarp Index, .Map, .X, .Y, .Dir
                        
                        AddLog Trim$(.Name) & " warped to [" & .Map & " | " & .X & " | " & .Y & "]"
                    End With
                Else
                    SendPlayerMsg Index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg Index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg Index, "You cannot warp yourself to yourself!", BrightRed
        End If
    Else
        SendPlayerMsg Index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InputName As String, i As Long
Dim Access As Byte, OldAccess As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Access = buffer.ReadByte
    Set buffer = Nothing
    
    i = FindPlayer(InputName)
        
    If i > 0 Then
        If Not i = Index Then
            If IsPlaying(i) Then
                If TempPlayer(i).UseChar > 0 Then
                    With Player(i, TempPlayer(i).UseChar)
                        If .Access >= Player(Index, TempPlayer(Index).UseChar).Access Then
                            SendPlayerMsg Index, "Player's access is greater than your own!", BrightRed
                        Else
                            OldAccess = .Access
                            .Access = Access
                            SendPlayerData i
                            
                            AddLog Trim$(.Name) & " got his access changed by " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " from " & OldAccess & " to " & Access
                        End If
                    End With
                Else
                    SendPlayerMsg Index, "Player is not online!", BrightRed
                End If
            Else
                SendPlayerMsg Index, "Player is not online!", BrightRed
            End If
        Else
            SendPlayerMsg Index, "You cannot change your own access!", BrightRed
        End If
    Else
        SendPlayerMsg Index, "Player is not online!", BrightRed
    End If
End Sub

Private Sub HandlePlayerPokemonMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
Dim tmpX As Long, tmpY As Long
Dim DuelIndex As Long
Dim RndNum As Byte

    If Not IsPlaying(Index) Then Exit Sub

    '//Check if can move
    If TempPlayer(Index).GettingMap Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub

    '//ToFast
    If PlayerPokemon(Index).MoveTmr > GetTickCount Then
        '//ToDo: Create something to prevent hacking
        'SendPlayerPokemonXY index, True
        'Exit Sub
    End If
    'If PlayerPokemon(Index).QueueMove > 0 Then
    '    SendPlayerPokemonXY Index, True
    '    Exit Sub
    'End If
    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = StatusEnum.Sleep Then
        SendPlayerPokemonXY Index, True
        Exit Sub
    End If
    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = StatusEnum.Frozen Then
        SendPlayerPokemonXY Index, True
        Exit Sub
    End If
    
    '//Desynced
    If Not PlayerPokemon(Index).X = tmpX Then
        SendPlayerPokemonXY Index, True
        Exit Sub
    End If
    If Not PlayerPokemon(Index).Y = tmpY Then
        SendPlayerPokemonXY Index, True
        Exit Sub
    End If
    
    '//Status
    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Status = StatusEnum.Poison Then
        If PlayerPokemon(Index).StatusMove >= 4 Then
            If PlayerPokemon(Index).StatusDamage > 0 Then
                If PlayerPokemon(Index).StatusDamage >= PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp Then
                    '//Dead
                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = 0
                    SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "-" & PlayerPokemon(Index).StatusDamage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, Magenta
                    SendPlayerPokemonVital Index
                    SendPlayerPokemonFaint Index
                    Exit Sub
                Else
                    '//Reduce
                    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp = PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).CurHp - PlayerPokemon(Index).StatusDamage
                    SendActionMsg Player(Index, TempPlayer(Index).UseChar).Map, "-" & PlayerPokemon(Index).StatusDamage, PlayerPokemon(Index).X * 32, PlayerPokemon(Index).Y * 32, Magenta
                    '//Update
                    SendPlayerPokemonVital Index
                End If
                '//Reset
                PlayerPokemon(Index).StatusMove = 0
            Else
                PlayerPokemon(Index).StatusDamage = (PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).MaxHp / 16)
            End If
        Else
            PlayerPokemon(Index).StatusMove = PlayerPokemon(Index).StatusMove + 1
        End If
        '//ToDo: Check if Badly Poisoned
    End If
    
    If PlayerPokemon(Index).IsConfuse = YES Then
        'Dir = Random(0, 3)
        'If Dir < 0 Then Dir = 0
        'If Dir > DIR_RIGHT Then Dir = DIR_RIGHT
        RndNum = Random(1, 10)
        If RndNum = 1 Then
            PlayerPokemon(Index).IsConfuse = 0
            If PlayerPokemon(Index).Num > 0 Then
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, Trim$(Pokemon(PlayerPokemon(Index).Num).Name) & " snapped out of confusion", White
                    Case LANG_EN: AddAlert Index, Trim$(Pokemon(PlayerPokemon(Index).Num).Name) & " snapped out of confusion", White
                    Case LANG_ES: AddAlert Index, Trim$(Pokemon(PlayerPokemon(Index).Num).Name) & " snapped out of confusion", White
                End Select
                SendPlayerPokemonStatus Index
            End If
        End If
    End If
    
    PlayerPokemon(Index).MoveTmr = GetTickCount + 200
    Call PlayerPokemonMove(Index, Dir)
End Sub

Private Sub HandlePlayerPokemonDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing
    
    '//Prevent hacking
    If Dir < 0 Or Dir > DIR_RIGHT Then Exit Sub
    
    PlayerPokemon(Index).Dir = Dir
    
    SendPlayerDir Index, True
End Sub

Private Sub HandleGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ItemNum As Long, ItemVal As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ItemNum = buffer.ReadLong
    ItemVal = buffer.ReadLong
    Set buffer = Nothing
    
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " , Admin Rights: Get Item#" & ItemNum & " x" & ItemVal
    TryGivePlayerItem Index, ItemNum, ItemVal
End Sub

Private Sub HandlePlayerPokemonState(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim State As Byte
Dim PokeSlot As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    State = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If Player(Index, TempPlayer(Index).UseChar).Action > 0 Then Exit Sub
    
    '//Check if exist
    If PokeSlot > 0 And PokeSlot <= MAX_PLAYER_POKEMON Then
        If State = 1 Then
            If PlayerPokemon(Index).Num > 0 Then Exit Sub
            
            If PlayerPokemons(Index).Data(PokeSlot).Num > 0 Then
                '//Make sure it's still alive
                If PlayerPokemons(Index).Data(PokeSlot).CurHp > 0 Then
                    If TempPlayer(Index).DuelReset = YES Then
                        '//Check if in duel, reset timer
                        If TempPlayer(Index).InDuel > 0 Then
                            If IsPlaying(TempPlayer(Index).InDuel) Then
                                If TempPlayer(TempPlayer(Index).InDuel).UseChar > 0 Then
                                    TempPlayer(Index).DuelTime = 3
                                    TempPlayer(Index).DuelTimeTmr = GetTickCount + 1000
                                    TempPlayer(TempPlayer(Index).InDuel).DuelTime = 3
                                    TempPlayer(TempPlayer(Index).InDuel).DuelTimeTmr = GetTickCount + 1000
                                End If
                            End If
                        End If
                        If TempPlayer(Index).InNpcDuel > 0 Then
                            TempPlayer(Index).DuelTime = 3
                            TempPlayer(Index).DuelTimeTmr = GetTickCount + 1000
                        End If
                        TempPlayer(Index).DuelReset = NO
                    End If
                    If PlayerPokemons(Index).Data(PokeSlot).Level <= (Player(Index, TempPlayer(Index).UseChar).Level + 10) Then
                        SpawnPlayerPokemon Index, PokeSlot
                    Else
                        AddAlert Index, "Not enough level", White
                    End If
                End If
            End If
        Else
            If PlayerPokemon(Index).Num = 0 Then Exit Sub
            
            ClearPlayerPokemon Index
        End If
    End If
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte
Dim movesetNum As Byte

    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Not IsPlaying(Index) Then Exit Sub
    
    If PlayerPokemon(Index).Num > 0 Then
        If PlayerPokemon(Index).slot <= 0 Then Exit Sub
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        MoveSlot = buffer.ReadByte
        Set buffer = Nothing
        
        '//ToFast
        If PlayerPokemon(Index).AtkTmr > GetTickCount Then
            Exit Sub
        End If
            
        '//Exit out if there's a Queue Move
        If PlayerPokemon(Index).QueueMove > 0 Then
            Exit Sub
        End If
        
        '//Exit out if in duel timer
        If TempPlayer(Index).DuelTime > 0 Then
            Exit Sub
        End If
        
        '//Select Moveslot
        If MoveSlot > 0 Then
            '//Check if moveslot is not empty
            If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(MoveSlot).Num > 0 Then
                PlayerCastMove Index, PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).Moveset(MoveSlot).Num, MoveSlot, True
            Else
                '//Use Struggle
                PlayerCastMove Index, 1, 0, False
            End If
        Else
            '//Use Struggle
            PlayerCastMove Index, 1, 0, False
        End If
        
        PlayerPokemon(Index).AtkTmr = GetTickCount + 500
        SendAttack Index, Player(Index, TempPlayer(Index).UseChar).Map
    End If
End Sub

Private Sub HandleChangePassword(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim newPass As String, oldPass As String
Dim filename As String

    If IsConnected(Index) And Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        newPass = Trim$(buffer.ReadString)
        oldPass = Trim$(buffer.ReadString)
        
        '//Make sure it's connected
        If Len(Account(Index).Username) > 0 Then
            '//Check if oldpassword match
            If Not Trim$(Account(Index).Password) = oldPass Then
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Incorrect password", White
                    Case LANG_EN: AddAlert Index, "Incorrect password", White
                    Case LANG_ES: AddAlert Index, "Incorrect password", White
                End Select
                Exit Sub
            End If
            
            '//Save
            filename = App.Path & "\data\accounts\" & Trim$(Account(Index).Username) & "\account.ini"
            
            '//Check if file exist
            If Not FileExist(filename) Then
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Error loading account", White
                    Case LANG_EN: AddAlert Index, "Error loading account", White
                    Case LANG_ES: AddAlert Index, "Error loading account", White
                End Select
                Exit Sub
            End If
            
            Call PutVar(filename, "Account", "Password", newPass)
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Password successfully changed", White
                Case LANG_EN: AddAlert Index, "Password successfully changed", White
                Case LANG_ES: AddAlert Index, "Password successfully changed", White
            End Select
            
            TextAdd frmServer.txtLog, "Account: " & Trim$(Account(Index).Username) & " changed it's password"
            AddLog "Account: " & Trim$(Account(Index).Username) & " changed it's password"
        End If
        Set buffer = Nothing
    End If
End Sub

Private Sub HandleReplaceNewMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte
Dim MoveNum As Long
Dim OldMove As Long

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).MoveLearnPokeSlot <= 0 Or TempPlayer(Index).MoveLearnPokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(TempPlayer(Index).MoveLearnPokeSlot).Num <= 0 Then Exit Sub
    If TempPlayer(Index).MoveLearnNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MoveSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Check
    If MoveSlot > 0 And MoveSlot <= MAX_MOVESET Then
        MoveNum = TempPlayer(Index).MoveLearnNum
        '//Replace
        With PlayerPokemons(Index).Data(TempPlayer(Index).MoveLearnPokeSlot)
            OldMove = .Moveset(MoveSlot).Num
            .Moveset(MoveSlot).Num = MoveNum
            .Moveset(MoveSlot).TotalPP = PokemonMove(MoveNum).PP
            .Moveset(MoveSlot).CurPP = .Moveset(MoveSlot).TotalPP
            .Moveset(MoveSlot).CD = 0
            SendPlayerPokemonSlot Index, TempPlayer(Index).MoveLearnPokeSlot
            '//Send Msg
            AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & "'s pokemon, " & Trim$(Pokemon(.Num).Name) & ", learned the move " & Trim$(PokemonMove(MoveNum).Name) & " replacing " & Trim$(PokemonMove(OldMove).Name)
            SendPlayerMsg Index, Trim$(Pokemon(.Num).Name) & " learned the move " & Trim$(PokemonMove(MoveNum).Name), White
        End With
        
        TempPlayer(Index).MoveLearnPokeSlot = 0
        TempPlayer(Index).MoveLearnNum = 0
        TempPlayer(Index).MoveLearnIndex = 0
        
        '//Continue Checking New Move
        CheckNewMove Index, TempPlayer(Index).MoveLearnPokeSlot, TempPlayer(Index).MoveLearnIndex
    Else
        TempPlayer(Index).MoveLearnPokeSlot = 0
        TempPlayer(Index).MoveLearnNum = 0
        TempPlayer(Index).MoveLearnIndex = 0
    End If
End Sub

Private Sub HandleEvolvePoke(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim EvolveSlot As Byte
Dim PokeNum As Long, EvolveNum As Long
Dim statNu As Byte
Dim itemSlot As Byte

    '//Check Error
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    If PlayerPokemon(Index).slot <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    EvolveSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Check
    If EvolveSlot > 0 And EvolveSlot <= MAX_EVOLVE Then
        PokeNum = PlayerPokemon(Index).Num
        EvolveNum = Pokemon(PokeNum).EvolveNum(EvolveSlot)
        If EvolveNum > 0 Then
            '//Check Condition
            With PlayerPokemons(Index).Data(PlayerPokemon(Index).slot)
                If .Level < Pokemon(.Num).EvolveLevel(EvolveSlot) Then
                    Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                        Case LANG_EN: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                        Case LANG_ES: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                    End Select
                    Exit Sub
                End If
                
                Select Case Pokemon(.Num).EvolveCondition(EvolveSlot)
                    Case EVOLVE_CONDT_ITEM
                        If Pokemon(.Num).EvolveConditionData(EvolveSlot) > 0 Then
                            itemSlot = checkItem(Index, Pokemon(.Num).EvolveConditionData(EvolveSlot))
                            If itemSlot > 0 Then
                                '//Take Item
                                PlayerInv(Index).Data(itemSlot).Value = PlayerInv(Index).Data(itemSlot).Value - 1
                                If PlayerInv(Index).Data(itemSlot).Value <= 0 Then
                                    '//Clear Item
                                    PlayerInv(Index).Data(itemSlot).Num = 0
                                    PlayerInv(Index).Data(itemSlot).Value = 0
                                    PlayerInv(Index).Data(itemSlot).TmrCooldown = 0
                                End If
                                SendPlayerInvSlot Index, itemSlot
                            Else
                                Select Case TempPlayer(Index).CurLanguage
                                    Case LANG_PT: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                                    Case LANG_EN: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                                    Case LANG_ES: AddAlert Index, "Your pokemon is not yet ready to evolve", White
                                End Select
                                Exit Sub
                            End If
                        End If
                End Select
                
                '//Change
                .Num = EvolveNum
                '//Calculate new stat
                For statNu = 1 To StatEnum.Stat_Count - 1
                    .Stat(statNu).Value = CalculatePokemonStat(statNu, .Num, .Level, .Stat(statNu).EV, .Stat(statNu).IV, .Nature)
                Next
                
                .MaxHp = .Stat(StatEnum.HP).Value
                
                '//Send Animation
                SendPlayAnimation Player(Index, TempPlayer(Index).UseChar).Map, 76, PlayerPokemon(Index).X, PlayerPokemon(Index).Y ' ToDo: Change to 76
                
                '//Update Map Poke
                PlayerPokemon(Index).Num = .Num
                
                '//Send Update
                SendPlayerPokemonSlot Index, PlayerPokemon(Index).slot
                SendPlayerPokemonData Index, Player(Index, TempPlayer(Index).UseChar).Map
                
                '//Check New Move
                CheckNewMove Index, PlayerPokemon(Index).slot
                
                AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " evolved " & Trim$(Pokemon(PokeNum).Name) & " to " & Trim$(Pokemon(EvolveNum).Name)
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                    Case LANG_EN: AddAlert Index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                    Case LANG_ES: AddAlert Index, "Congratulations! Your " & Trim$(Pokemon(PokeNum).Name) & " has evolved into " & Trim$(Pokemon(EvolveNum).Name), White
                End Select

                '//Add pokedex
                AddPlayerPokedex Index, .Num, YES, YES
                SendPlaySound "evolve.wav", Player(Index, TempPlayer(Index).UseChar).Map
            End With
        End If
    End If
End Sub

Private Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim itemSlot As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    itemSlot = buffer.ReadByte
    Set buffer = Nothing
    
    PlayerUseItem Index, itemSlot
    TempPlayer(Index).MapSwitchTmr = NO
End Sub

Private Sub HandleSwitchInvSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim OldSlot As Byte, NewSlot As Byte
    Dim OldInvData As PlayerInvDataRec, NewInvData As PlayerInvDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing

    If OldSlot <= 0 Then Exit Sub
    If OldSlot > MAX_PLAYER_INV Then Exit Sub
    If NewSlot <= 0 Then Exit Sub
    If NewSlot > MAX_PLAYER_INV Then Exit Sub

    '//Prevenção contra slots blockeados
    If PlayerInv(Index).Data(OldSlot).Locked = YES Or PlayerInv(Index).Data(NewSlot).Locked = YES Then
        Exit Sub
    End If

    '//Store Data
    OldInvData = PlayerInv(Index).Data(OldSlot)
    NewInvData = PlayerInv(Index).Data(NewSlot)

    '//Replace Data
    PlayerInv(Index).Data(OldSlot) = NewInvData
    PlayerInv(Index).Data(NewSlot) = OldInvData

    '//Update
    SendPlayerInvSlot Index, OldSlot
    SendPlayerInvSlot Index, NewSlot
End Sub

Private Sub HandleGotData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InUsed As Byte
    Dim Data1 As Long, Data2 As Long, Data3 As Long
    Dim CatchRate As Single
    Dim CatchValue As Long
    Dim NotTake As Boolean

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InUsed = buffer.ReadByte
    Data1 = buffer.ReadLong
    Data2 = buffer.ReadLong
    Data3 = buffer.ReadLong
    Set buffer = Nothing

    If InUsed = YES Then
        '//Make Sure Usage of item is available
        If Not IsPlaying(Index) Then GoTo Continue
        If TempPlayer(Index).UseChar <= 0 Then GoTo Continue
        If TempPlayer(Index).TmpUseInvSlot <= 0 Then GoTo Continue
        If PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num <= 0 Then GoTo Continue

        '//Check Type
        Select Case Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Type
        Case ItemTypeEnum.PokeBall
            If TempPlayer(Index).TmpCatchPokeNum > 0 Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You are currently capturing a pokemon", White
                Case LANG_EN: AddAlert Index, "You are currently capturing a pokemon", White
                Case LANG_ES: AddAlert Index, "You are currently capturing a pokemon", White
                End Select
                GoTo Continue
            End If

            '//Make sure we still have slot
            If CountFreePokemonSlot(Index) <= 0 Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You don't have any free slot left to capture a pokemon", White
                Case LANG_EN: AddAlert Index, "You don't have any free slot left to capture a pokemon", White
                Case LANG_ES: AddAlert Index, "You don't have any free slot left to capture a pokemon", White
                End Select
                GoTo Continue
            End If

            '//Make sure pokemon is not empty
            If Data1 <= 0 Or Data1 > MAX_GAME_POKEMON Then GoTo Continue
            '//Check if exist on map
            If MapPokemon(Data1).Num <= 0 Then GoTo Continue
            If Not MapPokemon(Data1).Map = Player(Index, TempPlayer(Index).UseChar).Map Then GoTo Continue

            '//Check if inrange
            If MapPokemon(Data1).X < Player(Index, TempPlayer(Index).UseChar).X - 4 Or MapPokemon(Data1).X > Player(Index, TempPlayer(Index).UseChar).X + 4 Or MapPokemon(Data1).Y < Player(Index, TempPlayer(Index).UseChar).Y - 4 Or MapPokemon(Data1).Y > Player(Index, TempPlayer(Index).UseChar).Y + 4 Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Out of range", White
                Case LANG_EN: AddAlert Index, "Out of range", White
                Case LANG_ES: AddAlert Index, "Out of range", White
                End Select
                GoTo Continue
            End If

            '//Check if catchable
            If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then
                If Spawn(Data1).CanCatch = YES Then
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You cannot catch this Pokemon", White
                    Case LANG_EN: AddAlert Index, "You cannot catch this Pokemon", White
                    Case LANG_ES: AddAlert Index, "You cannot catch this Pokemon", White
                    End Select
                    GoTo Continue
                End If
            End If

            '//Make sure no one is trying to catch this pokemon
            If MapPokemon(Data1).InCatch = YES Then GoTo Continue

            If Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data3 = YES Then
                '// Success
                '//Give Pokemon
                If CountFreePokemonSlot(Index) < 5 Then
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                    Case LANG_EN: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                    Case LANG_ES: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                    End Select
                End If

                '//Give Player Pokemon
                If CatchMapPokemonData(Index, Data1, Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data2) Then
                    '//Success
                    '//Clear Map Pokemon
                    TempPlayer(Index).TmpCatchUseBall = Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data2
                    SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).X, MapPokemon(Data1).Y, 2, TempPlayer(Index).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                    ClearMapPokemon Data1

                    TempPlayer(Index).TmpCatchPokeNum = 0
                    TempPlayer(Index).TmpCatchTimer = 0
                    TempPlayer(Index).TmpCatchTries = 0
                    TempPlayer(Index).TmpCatchValue = 0
                    TempPlayer(Index).TmpCatchUseBall = 0
                Else
                    '//Broke
                    MapPokemon(Data1).InCatch = NO
                    MapPokemon(Data1).targetType = TARGET_TYPE_PLAYER
                    MapPokemon(Data1).TargetIndex = Index
                    SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).X, MapPokemon(Data1).Y, 3, TempPlayer(Index).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                    TempPlayer(Index).TmpCatchPokeNum = 0
                    TempPlayer(Index).TmpCatchTimer = 0
                    TempPlayer(Index).TmpCatchTries = 0
                    TempPlayer(Index).TmpCatchValue = 0
                    TempPlayer(Index).TmpCatchUseBall = 0
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Your Pokeball broke", White
                    Case LANG_EN: AddAlert Index, "Your Pokeball broke", White
                    Case LANG_ES: AddAlert Index, "Your Pokeball broke", White
                    End Select
                End If
            Else
                '//Do Catch
                CatchRate = Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data1
                CatchRate = CatchRate / 10

                If MapPokemon(Data1).CurHp > 0 Then
                    '//ToDo: 1 = Status Modifier
                    CatchValue = (((3 * MapPokemon(Data1).MaxHp - 2 * MapPokemon(Data1).CurHp) * Pokemon(MapPokemon(Data1).Num).CatchRate * CatchRate) / (3 * MapPokemon(Data1).MaxHp)) * 1

                    If CatchValue >= 255 Then
                        '// Success
                        '//Give Pokemon
                        If CountFreePokemonSlot(Index) < 5 Then
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                            Case LANG_EN: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                            Case LANG_ES: AddAlert Index, "Warning: You only have few slot left for pokemon", White
                            End Select
                        End If

                        '//Give Player Pokemon
                        If CatchMapPokemonData(Index, Data1, TempPlayer(Index).TmpCatchUseBall) Then
                            '//Success
                            '//Clear Map Pokemon
                            TempPlayer(Index).TmpCatchUseBall = Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data2
                            SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).X, MapPokemon(Data1).Y, 2, TempPlayer(Index).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                            ClearMapPokemon Data1

                            TempPlayer(Index).TmpCatchPokeNum = 0
                            TempPlayer(Index).TmpCatchTimer = 0
                            TempPlayer(Index).TmpCatchTries = 0
                            TempPlayer(Index).TmpCatchValue = 0
                            TempPlayer(Index).TmpCatchUseBall = 0
                        Else
                            '//Broke
                            MapPokemon(Data1).InCatch = NO
                            MapPokemon(Data1).targetType = TARGET_TYPE_PLAYER
                            MapPokemon(Data1).TargetIndex = Index
                            SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).X, MapPokemon(Data1).Y, 3, TempPlayer(Index).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                            TempPlayer(Index).TmpCatchPokeNum = 0
                            TempPlayer(Index).TmpCatchTimer = 0
                            TempPlayer(Index).TmpCatchTries = 0
                            TempPlayer(Index).TmpCatchValue = 0
                            TempPlayer(Index).TmpCatchUseBall = 0
                            Select Case TempPlayer(Index).CurLanguage
                            Case LANG_PT: AddAlert Index, "Your Pokeball broke", White
                            Case LANG_EN: AddAlert Index, "Your Pokeball broke", White
                            Case LANG_ES: AddAlert Index, "Your Pokeball broke", White
                            End Select
                        End If
                    Else
                        TempPlayer(Index).TmpCatchPokeNum = Data1
                        TempPlayer(Index).TmpCatchTimer = GetTickCount + 250
                        TempPlayer(Index).TmpCatchTries = 0
                        TempPlayer(Index).TmpCatchValue = CatchValue
                        TempPlayer(Index).TmpCatchUseBall = Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data2    '//ToDo: Pokeball
                        MapPokemon(Data1).InCatch = YES
                        SendMapPokemonCatchState MapPokemon(Data1).Map, Data1, MapPokemon(Data1).X, MapPokemon(Data1).Y, 0, TempPlayer(Index).TmpCatchUseBall    '// 0 = Init, 1 = Shake, 2 = Success, 3 = Fail
                    End If
                End If
            End If
        Case ItemTypeEnum.Medicine
            '//Revive
            If Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data1 = 4 Then
                If Data1 <= 0 Or Data1 > MAX_PLAYER_POKEMON Then Exit Sub

                If PlayerPokemons(Index).Data(Data1).CurHp <= 0 Then
                    PlayerPokemons(Index).Data(Data1).CurHp = PlayerPokemons(Index).Data(Data1).MaxHp * (Item(PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num).Data2 / 100)
                    SendPlayerPokemonSlot Index, Data1

                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Pokemon was revived", White
                    Case LANG_EN: AddAlert Index, "Pokemon was revived", White
                    Case LANG_ES: AddAlert Index, "Pokemon was revived", White
                    End Select
                    NotTake = False
                Else
                    NotTake = True
                End If
            End If
        End Select

        '//Take Item
        If Not NotTake Then ' -> P/ usar com o revive
            PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Value = PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Value - 1
            If PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Value <= 0 Then
                '//Clear Item
                PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Num = 0
                PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).Value = 0
                PlayerInv(Index).Data(TempPlayer(Index).TmpUseInvSlot).TmrCooldown = 0
            End If
            SendPlayerInvSlot Index, TempPlayer(Index).TmpUseInvSlot
            TempPlayer(Index).TmpUseInvSlot = 0
        End If
    End If

Continue:
    'AddAlert Index, "Invalid Target", White
    '//Clear
    TempPlayer(Index).TmpUseInvSlot = 0
End Sub

Private Sub HandleOpenStorage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageType As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageType = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageType > 0 Then
        If TempPlayer(Index).StorageType > 0 Then Exit Sub
        TempPlayer(Index).StorageType = StorageType
    Else
        TempPlayer(Index).StorageType = StorageType
    End If
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " enter storage"
    SendStorage Index
End Sub

Private Sub HandleDepositItemTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim StorageSlot As Byte, StorageData As Byte, InvSlot As Byte
    Dim gameValue As Long
    Dim MsgFrom As String

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    'StorageData = buffer.ReadByte
    InvSlot = buffer.ReadByte
    gameValue = buffer.ReadLong
    Set buffer = Nothing

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If StorageSlot <= 0 Or StorageSlot > 5 Then Exit Sub
    If InvSlot <= 0 Or InvSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInvStorage(Index).slot(StorageSlot).Unlocked = NO Then Exit Sub
    If gameValue <= 0 Then Exit Sub
    
    '//Verifica se o jogador possui a quantidade
    If Not HasInvItem(Index, GetPlayerInvItemNum(Index, InvSlot)) >= gameValue Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Você não possui " & gameValue, White
        Case LANG_EN: AddAlert Index, "Você não possui " & gameValue, White
        Case LANG_ES: AddAlert Index, "Você não possui " & gameValue, White
        End Select
    
        Exit Sub
    End If
    
    '//Place item to that part
    If TryGiveStorageItem(Index, StorageSlot, PlayerInv(Index).Data(InvSlot).Num, gameValue, PlayerInv(Index).Data(InvSlot).TmrCooldown, MsgFrom) Then
        PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - gameValue
        If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
            PlayerInv(Index).Data(InvSlot).Num = 0
            PlayerInv(Index).Data(InvSlot).Value = 0
            PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
        End If

        '//Update
        SendPlayerInvSlot Index, InvSlot
    End If
End Sub

Private Sub HandleSwitchStorageSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageSlot As Byte, OldSlot As Byte, NewSlot As Byte
Dim OldStorageData As PlayerInvStorageDataRec, NewStorageData As PlayerInvStorageDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageSlot <= 0 Or StorageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If OldSlot <= 0 Or OldSlot > MAX_STORAGE Then Exit Sub
    If NewSlot <= 0 Or NewSlot > MAX_STORAGE Then Exit Sub

    '//Store Data
    OldStorageData = PlayerInvStorage(Index).slot(StorageSlot).Data(OldSlot)
    NewStorageData = PlayerInvStorage(Index).slot(StorageSlot).Data(NewSlot)
    
    '//Replace Data
    PlayerInvStorage(Index).slot(StorageSlot).Data(OldSlot) = NewStorageData
    PlayerInvStorage(Index).slot(StorageSlot).Data(NewSlot) = OldStorageData
    
    '//Update
    SendPlayerInvStorageSlot Index, StorageSlot, OldSlot
    SendPlayerInvStorageSlot Index, StorageSlot, NewSlot
End Sub

Private Sub HandleWithdrawItemTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim StorageSlot As Byte, StorageData As Byte, InvSlot As Byte
    Dim checkSameSlot As Byte
    Dim gameValue As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    'InvSlot = buffer.ReadByte
    gameValue = buffer.ReadLong
    Set buffer = Nothing

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If StorageSlot <= 0 Or StorageSlot > 5 Then Exit Sub
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Sub
    If StorageSlot <= 0 Or StorageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If PlayerInvStorage(Index).slot(StorageSlot).Unlocked = False Then Exit Sub
    If gameValue <= 0 Then Exit Sub

    '//Verifica se o banco possui o item e quantidade
    If Not HasStorageItem(Index, StorageSlot, GetPlayerStorageItemNum(Index, StorageSlot, StorageData)) >= gameValue Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Você não possui " & gameValue, White
        Case LANG_EN: AddAlert Index, "Você não possui " & gameValue, White
        Case LANG_ES: AddAlert Index, "Você não possui " & gameValue, White
        End Select
        
        Exit Sub
    End If

    If TryGivePlayerItem(Index, PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Num, gameValue, PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).TmrCooldown) Then
        PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Value = PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Value - gameValue
        If PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Value <= 0 Then
            PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Num = 0
            PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).Value = 0
            PlayerInvStorage(Index).slot(StorageSlot).Data(StorageData).TmrCooldown = 0
        End If
        '//Update
        SendPlayerInvStorageSlot Index, StorageSlot, StorageData
    End If
End Sub

Private Sub HandleConvo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim cType As Byte
Dim Data1 As Long
Dim NpcNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    cType = buffer.ReadByte
    Data1 = buffer.ReadLong
    Set buffer = Nothing
    
    If cType = 1 Then
        '//Check if npc have convo
        If Data1 > 0 And Data1 <= MAX_MAP_NPC Then
            '//Make sure NPC is not busy
            If MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, Data1).InBattle <= 0 Then
                NpcNum = MapNpc(Player(Index, TempPlayer(Index).UseChar).Map, Data1).Num
                If NpcNum > 0 Then
                    If Npc(NpcNum).Convo > 0 Then
                        TempPlayer(Index).CurConvoNum = Npc(NpcNum).Convo
                        TempPlayer(Index).CurConvoData = 0 '//Always start at 0
                        TempPlayer(Index).CurConvoNpc = NpcNum
                        TempPlayer(Index).CurConvoMapNpc = Data1
                        ProcessConversation Index, TempPlayer(Index).CurConvoNum, TempPlayer(Index).CurConvoData, TempPlayer(Index).CurConvoNpc
                    End If
                End If
            Else
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You cannot talk with this NPC at the moment", White
                    Case LANG_EN: AddAlert Index, "You cannot talk with this NPC at the moment", White
                    Case LANG_ES: AddAlert Index, "You cannot talk with this NPC at the moment", White
                End Select
            End If
        End If
    ElseIf cType = 2 Then
        If Data1 > 0 And Data1 <= MAX_CONVERSATION Then
            TempPlayer(Index).CurConvoNum = Data1
            TempPlayer(Index).CurConvoData = 0 '//Always start at 0
            TempPlayer(Index).CurConvoNpc = 0
            TempPlayer(Index).CurConvoMapNpc = 0
            ProcessConversation Index, TempPlayer(Index).CurConvoNum, TempPlayer(Index).CurConvoData, TempPlayer(Index).CurConvoNpc
        End If
    End If
End Sub

Private Sub HandleProcessConvo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tReply As Byte

    If TempPlayer(Index).CurConvoNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tReply = buffer.ReadByte
    Set buffer = Nothing
    
    ProcessConversation Index, TempPlayer(Index).CurConvoNum, TempPlayer(Index).CurConvoData, TempPlayer(Index).CurConvoNpc, tReply
End Sub

Private Sub HandleDepositPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Byte
Dim StorageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    '//Make sure they don't deposit their last pokemon
    If CountPlayerPokemon(Index) <= 1 Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "You cannot deposit your last pokemon", White
            Case LANG_EN: AddAlert Index, "You cannot deposit your last pokemon", White
            Case LANG_ES: AddAlert Index, "You cannot deposit your last pokemon", White
        End Select
        Exit Sub
    End If
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    
    StorageData = FindFreePokeStorageSlot(Index, StorageSlot)
    '//Check if there's available slot
    If StorageData > 0 Then
        'CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
        Call CopyMemory(ByVal VarPtr(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)), ByVal VarPtr(PlayerPokemons(Index).Data(PokeSlot)), LenB(PlayerPokemons(Index).Data(PokeSlot)))
        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(Index).Data(PokeSlot)), LenB(PlayerPokemons(Index).Data(PokeSlot)))
        '//reupdate order
        UpdatePlayerPokemonOrder Index
        '//update
        SendPlayerPokemons Index
        SendPlayerPokemonStorageSlot Index, StorageSlot, StorageData
    Else
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "There's no available space on this slot", White
            Case LANG_EN: AddAlert Index, "There's no available space on this slot", White
            Case LANG_ES: AddAlert Index, "There's no available space on this slot", White
        End Select
        Exit Sub
    End If
End Sub

Private Sub HandleWithdrawPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Byte
Dim StorageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
    
    '//Make sure they don't deposit their last pokemon
    If CountPlayerPokemon(Index) >= MAX_PLAYER_POKEMON Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "There's no slot available", White
            Case LANG_EN: AddAlert Index, "There's no slot available", White
            Case LANG_ES: AddAlert Index, "There's no slot available", White
        End Select
        Exit Sub
    End If
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Sub
    
    PokeSlot = FindOpenPokeSlot(Index)
    '//Check if there's available slot
    If PokeSlot > 0 Then
        'CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
        Call CopyMemory(ByVal VarPtr(PlayerPokemons(Index).Data(PokeSlot)), ByVal VarPtr(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)))
        Call ZeroMemory(ByVal VarPtr(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)))
        '//reupdate order
        UpdatePlayerPokemonOrder Index
        '//update
        SendPlayerPokemons Index
        SendPlayerPokemonStorageSlot Index, StorageSlot, StorageData
    Else
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "There's no slot available", White
            Case LANG_EN: AddAlert Index, "There's no slot available", White
            Case LANG_ES: AddAlert Index, "There's no slot available", White
        End Select
        Exit Sub
    End If
End Sub

Private Sub HandleReleasePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageSlot As Byte
Dim StorageData As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageSlot <= 0 Or StorageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If StorageData <= 0 Or StorageData > MAX_STORAGE Then Exit Sub
    
    If PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData).Num > 0 Then
        Call ZeroMemory(ByVal VarPtr(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)), LenB(PlayerPokemonStorage(Index).slot(StorageSlot).Data(StorageData)))
        SendPlayerPokemonStorageSlot Index, StorageSlot, StorageData
    End If
End Sub

Private Sub HandleSwitchStoragePokeSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageSlot As Byte, OldSlot As Byte, NewSlot As Byte
Dim OldStorageData As PlayerPokemonStorageDataRec, NewStorageData As PlayerPokemonStorageDataRec

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageSlot <= 0 Or StorageSlot > MAX_STORAGE_SLOT Then Exit Sub
    If OldSlot <= 0 Or OldSlot > MAX_STORAGE Then Exit Sub
    If NewSlot <= 0 Or NewSlot > MAX_STORAGE Then Exit Sub

    '//Store Data
    OldStorageData = PlayerPokemonStorage(Index).slot(StorageSlot).Data(OldSlot)
    NewStorageData = PlayerPokemonStorage(Index).slot(StorageSlot).Data(NewSlot)
    
    '//Replace Data
    PlayerPokemonStorage(Index).slot(StorageSlot).Data(OldSlot) = NewStorageData
    PlayerPokemonStorage(Index).slot(StorageSlot).Data(NewSlot) = OldStorageData
    
    '//Update
    SendPlayerPokemonStorageSlot Index, StorageSlot, OldSlot
    SendPlayerPokemonStorageSlot Index, StorageSlot, NewSlot
End Sub

Private Sub HandleSwitchStoragePoke(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim OldSlot As Byte, OldPokeStorage As Byte, NewPokeStorage As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    OldSlot = buffer.ReadByte
    OldPokeStorage = buffer.ReadByte
    NewPokeStorage = buffer.ReadByte
    Set buffer = Nothing
    
    If NewPokeStorage < 0 Or NewPokeStorage > MAX_STORAGE_SLOT Or NewPokeStorage = OldPokeStorage Then Exit Sub
    If OldPokeStorage < 0 Or OldPokeStorage > MAX_STORAGE_SLOT Then Exit Sub
    If PlayerPokemonStorage(Index).slot(NewPokeStorage).Unlocked = False Then Exit Sub

    SetNewPokeStorageSlotTo Index, OldSlot, OldPokeStorage, NewPokeStorage
End Sub

Private Sub SetNewPokeStorageSlotTo(ByVal Index As Long, ByVal OldPokeSlot, ByVal OldPokeStorage As Byte, ByVal NewPokeStorage As Byte)
    Dim OldStorageData As PlayerPokemonStorageDataRec, NewStorageData As PlayerPokemonStorageDataRec
    Dim i As Byte

    If PlayerPokemonStorage(Index).slot(NewPokeStorage).Unlocked = YES Then
        For i = 1 To MAX_STORAGE
            If PlayerPokemonStorage(Index).slot(NewPokeStorage).Data(i).Num = 0 Then
                '//Store Data
                OldStorageData = PlayerPokemonStorage(Index).slot(OldPokeStorage).Data(OldPokeSlot)
                NewStorageData = PlayerPokemonStorage(Index).slot(NewPokeStorage).Data(i)

                '//Replace Data
                PlayerPokemonStorage(Index).slot(OldPokeStorage).Data(OldPokeSlot) = NewStorageData
                PlayerPokemonStorage(Index).slot(NewPokeStorage).Data(i) = OldStorageData

                '//Update
                SendPlayerPokemonStorageSlot Index, OldPokeStorage, OldPokeSlot
                SendPlayerPokemonStorageSlot Index, NewPokeStorage, i
                
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub HandleSwitchStorageItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim OldSlot As Byte, OldItemStorage As Byte, NewItemStorage As Byte

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    OldSlot = buffer.ReadByte
    OldItemStorage = buffer.ReadByte
    NewItemStorage = buffer.ReadByte
    Set buffer = Nothing
    
    If NewItemStorage < 0 Or NewItemStorage > MAX_STORAGE_SLOT Or NewItemStorage = OldItemStorage Then Exit Sub
    If OldItemStorage < 0 Or OldItemStorage > MAX_STORAGE_SLOT Then Exit Sub
    If PlayerInvStorage(Index).slot(NewItemStorage).Unlocked = False Then Exit Sub

    SetNewItemStorageSlotTo Index, OldSlot, OldItemStorage, NewItemStorage
End Sub

Private Sub SetNewItemStorageSlotTo(ByVal Index As Long, ByVal OldItemSlot, ByVal OldItemStorage As Byte, ByVal NewItemStorage As Byte)
    Dim OldStorageData As PlayerInvStorageDataRec, NewStorageData As PlayerInvStorageDataRec
    Dim i As Byte

    If PlayerInvStorage(Index).slot(NewItemStorage).Unlocked = YES Then
        For i = 1 To MAX_STORAGE
            If PlayerInvStorage(Index).slot(NewItemStorage).Data(i).Num = 0 Then
                '//Store Data
                OldStorageData = PlayerInvStorage(Index).slot(OldItemStorage).Data(OldItemSlot)
                NewStorageData = PlayerInvStorage(Index).slot(NewItemStorage).Data(i)

                '//Replace Data
                PlayerInvStorage(Index).slot(OldItemStorage).Data(OldItemSlot) = NewStorageData
                PlayerInvStorage(Index).slot(NewItemStorage).Data(i) = OldStorageData

                '//Update
                SendPlayerInvStorageSlot Index, OldItemStorage, OldItemSlot
                SendPlayerInvStorageSlot Index, NewItemStorage, i
                
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
    SendOpenShop Index
End Sub

Private Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim ShopSlot As Byte, ShopVal As Long, MoneyValue As Long, IsCash As Byte, ShopValue As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InShop <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ShopSlot = buffer.ReadByte
    ShopVal = buffer.ReadLong
    Set buffer = Nothing

    If ShopSlot <= 0 Or ShopSlot > MAX_SHOP_ITEM Then Exit Sub
    If ShopVal <= 0 Then Exit Sub

    '//Give Item
    With Player(Index, TempPlayer(Index).UseChar)
        ShopValue = (Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Price * ShopVal)
        IsCash = Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).IsCash

        If IsCash = YES Then
            If .Cash >= ShopValue Then
                If TryGivePlayerItem(Index, Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num, ShopVal) Then
                    .Cash = .Cash - ShopValue
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_EN: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_ES: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    End Select
                    
                    Call SendPlayerCash(Index)
                End If
            End If
        Else
            If .Money >= ShopValue Then
                If TryGivePlayerItem(Index, Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num, ShopVal) Then
                    .Money = .Money - ShopValue
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_EN: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    Case LANG_ES: AddAlert Index, "You have successfully bought x" & ShopVal & " " & Trim$(Item(Shop(TempPlayer(Index).InShop).ShopItem(ShopSlot).Num).Name), White
                    End Select
                    
                    Call SendPlayerCash(Index)
                End If
            End If
        End If
    End With
End Sub

Private Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvSlot As Byte, InvVal As Long
    Dim aPrice As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InShop <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvSlot = buffer.ReadByte
    InvVal = buffer.ReadLong
    Set buffer = Nothing

    If InvSlot <= 0 Or InvSlot > MAX_PLAYER_INV Then Exit Sub
    If InvVal < 0 Then Exit Sub

    ' Não pode vender item de cash
    If Item(PlayerInv(Index).Data(InvSlot).Num).IsCash = YES Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Item não permitido para vender.", White
        Case LANG_EN: AddAlert Index, "Item not allowed to sell.", White
        Case LANG_ES: AddAlert Index, "Artículo no permitido para vender.", White
        End Select
        Exit Sub
    End If
    
    ' Não pode vender item Linked
    If Item(PlayerInv(Index).Data(InvSlot).Num).Linked = YES Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Este item não pode ser negociado.", White
        Case LANG_EN: AddAlert Index, "This item cannot be traded.", White
        Case LANG_ES: AddAlert Index, "Este artículo no se puede intercambiar.", White
        End Select
        Exit Sub
    End If

    '//Give Item
    With Player(Index, TempPlayer(Index).UseChar)
        If PlayerInv(Index).Data(InvSlot).Value < InvVal Then
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Invalid amount", White
            Case LANG_EN: AddAlert Index, "Invalid amount", White
            Case LANG_ES: AddAlert Index, "Invalid amount", White
            End Select
        Else
            If PlayerInv(Index).Data(InvSlot).Num > 0 Then
                aPrice = (Item(PlayerInv(Index).Data(InvSlot).Num).Price / 2) * InvVal
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(Index).Data(InvSlot).Num).Name) & " for $" & aPrice, White
                Case LANG_EN: AddAlert Index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(Index).Data(InvSlot).Num).Name) & " for $" & aPrice, White
                Case LANG_ES: AddAlert Index, "You have successfully sold x" & InvVal & " " & Trim$(Item(PlayerInv(Index).Data(InvSlot).Num).Name) & " for $" & aPrice, White
                End Select
                PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - InvVal
                If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
                    PlayerInv(Index).Data(InvSlot).Num = 0
                    PlayerInv(Index).Data(InvSlot).Value = 0
                    PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
                End If
                SendPlayerInvSlot Index, InvSlot
                .Money = .Money + aPrice
                If .Money >= MAX_MONEY Then
                    .Money = MAX_MONEY
                End If
                SendPlayerData Index
            End If
        End If
    End With
End Sub

Private Sub HandleRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim RequestType As Byte
    Dim requestIndex As Long
    Dim stillGotTarget As Boolean

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    stillGotTarget = False
    If TempPlayer(Index).PlayerRequest > 0 Then
        If IsPlaying(TempPlayer(Index).PlayerRequest) Then
            If TempPlayer(TempPlayer(Index).PlayerRequest).UseChar > 0 Then
                If Player(TempPlayer(Index).PlayerRequest, TempPlayer(TempPlayer(Index).PlayerRequest).UseChar).Map = Player(Index, TempPlayer(Index).UseChar).Map Then
                    '//Check if inrange
                    stillGotTarget = True
                End If
            End If
        End If
    End If

    If stillGotTarget Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Cancel your last player request", White
        Case LANG_EN: AddAlert Index, "Cancel your last player request", White
        Case LANG_ES: AddAlert Index, "Cancel your last player request", White
        End Select
    Else
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        RequestType = buffer.ReadByte
        requestIndex = buffer.ReadLong
        Set buffer = Nothing

        If requestIndex <= 0 Or requestIndex > MAX_PLAYER Then Exit Sub
        If Not IsPlaying(requestIndex) Then Exit Sub
        If TempPlayer(requestIndex).UseChar <= 0 Then Exit Sub

        If TempPlayer(requestIndex).PlayerRequest > 0 Or TempPlayer(requestIndex).InDuel > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Player is busy", White
            Case LANG_EN: AddAlert Index, "Player is busy", White
            Case LANG_ES: AddAlert Index, "Player is busy", White
            End Select
            'End If
            TempPlayer(Index).RequestType = 0
            TempPlayer(Index).PlayerRequest = 0
            SendRequest Index
            Exit Sub
        End If

        If Map(Player(requestIndex, TempPlayer(requestIndex).UseChar).Map).Moral = 1 Then
            '//Can't duel
            If RequestType = 1 Then    '//Duel
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You cannot duel here", White
                Case LANG_EN: AddAlert Index, "You cannot duel here", White
                Case LANG_ES: AddAlert Index, "You cannot duel here", White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If
        End If

        If Map(Player(Index, TempPlayer(Index).UseChar).Map).Moral = 1 Then
            '//Can't duel
            If RequestType = 1 Then    '//Duel
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You cannot duel here", White
                Case LANG_EN: AddAlert Index, "You cannot duel here", White
                Case LANG_ES: AddAlert Index, "You cannot duel here", White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If
        End If

        If TempPlayer(requestIndex).StorageType > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Player is busy", White
            Case LANG_EN: AddAlert Index, "Player is busy", White
            Case LANG_ES: AddAlert Index, "Player is busy", White
            End Select
            'End If
            TempPlayer(Index).RequestType = 0
            TempPlayer(Index).PlayerRequest = 0
            SendRequest Index
            Exit Sub
        End If

        If TempPlayer(Index).StorageType > 0 Then
            '//Can't duel
            'If RequestType = 1 Then '//Duel
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "You can't do that right now", White
            Case LANG_EN: AddAlert Index, "You can't do that right now", White
            Case LANG_ES: AddAlert Index, "You can't do that right now", White
            End Select
            'End If
            TempPlayer(Index).RequestType = 0
            TempPlayer(Index).PlayerRequest = 0
            SendRequest Index
            Exit Sub
        End If

        '//Trade
        If RequestType = 2 Then
        
            '//Verifica se o recebinte tem o mínimo de requerimento pra solicitar trade.
            If Player(requestIndex, TempPlayer(requestIndex).UseChar).Level < Options.TradeLvlMin Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "O player não é level maior que " & Options.TradeLvlMin, White
                Case LANG_EN: AddAlert Index, "O player não é level maior que " & Options.TradeLvlMin, White
                Case LANG_ES: AddAlert Index, "O player não é level maior que " & Options.TradeLvlMin, White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If

            '//Verifica se o solicitante tem o mínimo de requerimento pra solicitar trade.
            If Player(Index, TempPlayer(Index).UseChar).Level < Options.TradeLvlMin Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Você precisa ser level maior que " & Options.TradeLvlMin, White
                Case LANG_EN: AddAlert Index, "Você precisa ser level maior que " & Options.TradeLvlMin, White
                Case LANG_ES: AddAlert Index, "Você precisa ser level maior que " & Options.TradeLvlMin, White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If

            If Options.SameIp = YES Then
                If GetPlayerIP(Index) = GetPlayerIP(requestIndex) Then
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Mesmo Ip não permitido", White
                    Case LANG_EN: AddAlert Index, "Mesmo Ip não permitido", White
                    Case LANG_ES: AddAlert Index, "Mesmo Ip não permitido", White
                    End Select
                    TempPlayer(Index).RequestType = 0
                    TempPlayer(Index).PlayerRequest = 0
                    SendRequest Index
                    Exit Sub
                End If
            End If
        End If

        If RequestType = 3 Then
            If TempPlayer(requestIndex).InParty > 0 Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Player already have a party", White
                Case LANG_EN: AddAlert Index, "Player already have a party", White
                Case LANG_ES: AddAlert Index, "Player already have a party", White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If

            If PartyCount(Index) >= MAX_PARTY Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Party is full", White
                Case LANG_EN: AddAlert Index, "Party is full", White
                Case LANG_ES: AddAlert Index, "Party is full", White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If

            If TempPlayer(Index).InParty <= 0 Then
                Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "You are not in a party", White
                Case LANG_EN: AddAlert Index, "You are not in a party", White
                Case LANG_ES: AddAlert Index, "You are not in a party", White
                End Select
                TempPlayer(Index).RequestType = 0
                TempPlayer(Index).PlayerRequest = 0
                SendRequest Index
                Exit Sub
            End If


        End If

        TempPlayer(requestIndex).PlayerRequest = Index
        TempPlayer(requestIndex).RequestType = RequestType
        TempPlayer(Index).PlayerRequest = requestIndex
        TempPlayer(Index).RequestType = RequestType
        '//Send Request to client
        SendRequest requestIndex
    End If

    Exit Sub
exitdata:
    SendRequest Index
End Sub

Private Sub HandleRequestState(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim requestState As Byte
Dim RequestType As Byte
Dim requestIndex As Long
Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    requestState = buffer.ReadByte
    Set buffer = Nothing
    
    requestIndex = TempPlayer(Index).PlayerRequest
    RequestType = TempPlayer(Index).RequestType
    
    Select Case requestState
        Case 0 '//Cancel
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = Index Then
                            TempPlayer(requestIndex).PlayerRequest = 0
                            TempPlayer(requestIndex).RequestType = 0
                            SendRequest requestIndex
                            If RequestType = 1 Then '//Duel
                                Select Case TempPlayer(requestIndex).CurLanguage
                                    Case LANG_PT: AddAlert requestIndex, "Duel request has been cancelled", White
                                    Case LANG_EN: AddAlert requestIndex, "Duel request has been cancelled", White
                                    Case LANG_ES: AddAlert requestIndex, "Duel request has been cancelled", White
                                End Select
                            ElseIf RequestType = 2 Then '//Trade
                                Select Case TempPlayer(requestIndex).CurLanguage
                                    Case LANG_PT: AddAlert requestIndex, "Trade request has been cancelled", White
                                    Case LANG_EN: AddAlert requestIndex, "Trade request has been cancelled", White
                                    Case LANG_ES: AddAlert requestIndex, "Trade request has been cancelled", White
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
            TempPlayer(Index).PlayerRequest = 0
            TempPlayer(Index).RequestType = 0
            SendRequest Index
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Request has been cancelled", White
                Case LANG_EN: AddAlert Index, "Request has been cancelled", White
                Case LANG_ES: AddAlert Index, "Request has been cancelled", White
            End Select
        Case 1 '//Accept
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = Index Then
                            If RequestType = 1 Then '//Duel
                                '//Initiate duel
                                If CountPlayerPokemonAlive(Index) > 0 Then
                                    TempPlayer(Index).InDuel = requestIndex
                                    TempPlayer(Index).DuelTime = 11
                                    TempPlayer(Index).DuelTimeTmr = GetTickCount + 1000
                                    TempPlayer(requestIndex).InDuel = Index
                                    TempPlayer(requestIndex).DuelTime = 11
                                    TempPlayer(requestIndex).DuelTimeTmr = GetTickCount + 1000
                                    Select Case TempPlayer(Index).CurLanguage
                                        Case LANG_PT: AddAlert Index, "Duel invitation accepted", White
                                        Case LANG_EN: AddAlert Index, "Duel invitation accepted", White
                                        Case LANG_ES: AddAlert Index, "Duel invitation accepted", White
                                    End Select
                                    Select Case TempPlayer(requestIndex).CurLanguage
                                        Case LANG_PT: AddAlert requestIndex, "Duel invitation accepted", White
                                        Case LANG_EN: AddAlert requestIndex, "Duel invitation accepted", White
                                        Case LANG_ES: AddAlert requestIndex, "Duel invitation accepted", White
                                    End Select
                                Else
                                    Select Case TempPlayer(Index).CurLanguage
                                        Case LANG_PT: AddAlert Index, "You don't have any active pokemon to use", White
                                        Case LANG_EN: AddAlert Index, "You don't have any active pokemon to use", White
                                        Case LANG_ES: AddAlert Index, "You don't have any active pokemon to use", White
                                    End Select
                                    Select Case TempPlayer(requestIndex).CurLanguage
                                        Case LANG_PT: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                        Case LANG_EN: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                        Case LANG_ES: AddAlert requestIndex, "Player doesn't have any active pokemon to use", White
                                    End Select
                                End If
                            ElseIf RequestType = 2 Then '//Trade
                                
                                TempPlayer(Index).InTrade = requestIndex
                                TempPlayer(requestIndex).InTrade = Index
                                SendOpenTrade Index
                                SendOpenTrade requestIndex
                            ElseIf RequestType = 3 Then '//Party
                                '//Join Party
                                JoinParty requestIndex, Index
                                TempPlayer(Index).RequestType = 0
                                TempPlayer(Index).PlayerRequest = 0
                                SendRequest Index
                                TempPlayer(requestIndex).RequestType = 0
                                TempPlayer(requestIndex).PlayerRequest = 0
                                SendRequest requestIndex
                            End If
                        End If
                    End If
                End If
            End If
        Case Else '//Decline
            If requestIndex > 0 Then
                '//Cancel Request to index
                If IsPlaying(requestIndex) Then
                    If TempPlayer(requestIndex).UseChar > 0 Then
                        If TempPlayer(requestIndex).PlayerRequest = Index Then
                            TempPlayer(requestIndex).PlayerRequest = 0
                            TempPlayer(requestIndex).RequestType = 0
                            SendRequest requestIndex
                            Select Case TempPlayer(requestIndex).CurLanguage
                                Case LANG_PT: AddAlert requestIndex, "Request has been declined", White
                                Case LANG_EN: AddAlert requestIndex, "Request has been declined", White
                                Case LANG_ES: AddAlert requestIndex, "Request has been declined", White
                            End Select
                        End If
                    End If
                End If
            End If
            TempPlayer(Index).PlayerRequest = 0
            TempPlayer(Index).RequestType = 0
            SendRequest Index
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Request has been declined", White
                Case LANG_EN: AddAlert Index, "Request has been declined", White
                Case LANG_ES: AddAlert Index, "Request has been declined", White
            End Select
    End Select
End Sub

Private Sub HandleAddTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TradeType As Byte, TradeSlot As Long, TradeData As Long
    Dim slot As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InTrade <= 0 Then Exit Sub
    If TempPlayer(Index).TradeSet = YES Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TradeType = buffer.ReadByte
    TradeSlot = buffer.ReadLong
    TradeData = buffer.ReadLong
    Set buffer = Nothing

    '//Make sure you can't use the same slot
    For i = 1 To MAX_TRADE
        If TempPlayer(Index).TradeItem(i).TradeSlot = TradeSlot Then
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Trade slot already in used", White
            Case LANG_EN: AddAlert Index, "Trade slot already in used", White
            Case LANG_ES: AddAlert Index, "Trade slot already in used", White
            End Select
            Exit Sub
        End If
    Next

    '//Check available trade slot
    slot = FindOpenTradeSlot(Index)
    If slot > 0 Then
        With TempPlayer(Index).TradeItem(slot)
            Select Case TradeType
            Case 1    '//Item
                If TradeSlot <= 0 Or TradeSlot > MAX_PLAYER_INV Then Exit Sub
                If TradeData <= 0 Then Exit Sub

                ' Não pode negociar item Linked
                If Item(PlayerInv(Index).Data(TradeSlot).Num).Linked = YES Then
                    Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Este item não pode ser negociado.", White
                    Case LANG_EN: AddAlert Index, "This item cannot be traded.", White
                    Case LANG_ES: AddAlert Index, "Este artículo no se puede intercambiar.", White
                    End Select
                    Exit Sub
                End If

                .Num = PlayerInv(Index).Data(TradeSlot).Num
                If TradeData > PlayerInv(Index).Data(TradeSlot).Value Then
                    .Value = PlayerInv(Index).Data(TradeSlot).Value
                Else
                    .Value = TradeData
                End If

                .Level = 0
                For i = 1 To StatEnum.Stat_Count - 1
                    .Stat(i) = 0
                    .StatIV(i) = 0
                    .StatEV(i) = 0
                Next
                .CurHp = 0
                .MaxHp = 0
                .Nature = 0
                .IsShiny = 0
                .Happiness = 0
                .Gender = 0
                .Status = 0
                .CurExp = 0
                .nextExp = 0
                For i = 1 To MAX_MOVESET
                    .Moveset(i).Num = 0
                    .Moveset(i).CurPP = 0
                    .Moveset(i).TotalPP = 0
                Next
                .BallUsed = 0
                .HeldItem = 0
            Case 2    '//Pokemon
                If TradeSlot <= 0 Or TradeSlot > MAX_PLAYER_POKEMON Then Exit Sub

                ' Não pode negociar item de cash
                If PlayerPokemons(Index).Data(TradeSlot).HeldItem > 0 Then
                    ' Não pode negociar item Linked
                    If Item(PlayerPokemons(Index).Data(TradeSlot).HeldItem).Linked = YES Then
                        Select Case TempPlayer(Index).CurLanguage
                        Case LANG_PT: AddAlert Index, "Pokemon usa item que não pode ser negociado.", White
                        Case LANG_EN: AddAlert Index, "This pokemon equip item cannot be traded.", White
                        Case LANG_ES: AddAlert Index, "Este pokemon equip artículo no se puede intercambiar.", White
                        End Select
                        Exit Sub
                    End If
                End If

                .Num = PlayerPokemons(Index).Data(TradeSlot).Num
                .Value = 0

                .Level = PlayerPokemons(Index).Data(TradeSlot).Level
                For i = 1 To StatEnum.Stat_Count - 1
                    .Stat(i) = PlayerPokemons(Index).Data(TradeSlot).Stat(i).Value
                    .StatIV(i) = PlayerPokemons(Index).Data(TradeSlot).Stat(i).IV
                    .StatEV(i) = PlayerPokemons(Index).Data(TradeSlot).Stat(i).EV
                Next
                .CurHp = PlayerPokemons(Index).Data(TradeSlot).CurHp
                .MaxHp = PlayerPokemons(Index).Data(TradeSlot).MaxHp
                .Nature = PlayerPokemons(Index).Data(TradeSlot).Nature
                .IsShiny = PlayerPokemons(Index).Data(TradeSlot).IsShiny
                .Happiness = PlayerPokemons(Index).Data(TradeSlot).Happiness
                .Gender = PlayerPokemons(Index).Data(TradeSlot).Gender
                .Status = PlayerPokemons(Index).Data(TradeSlot).Status
                .CurExp = PlayerPokemons(Index).Data(TradeSlot).CurExp
                If .Num > 0 Then
                    .nextExp = GetPokemonNextExp(.Level, Pokemon(.Num).GrowthRate)
                Else
                    .nextExp = 0
                End If
                For i = 1 To MAX_MOVESET
                    .Moveset(i).Num = PlayerPokemons(Index).Data(TradeSlot).Moveset(i).Num
                    .Moveset(i).CurPP = PlayerPokemons(Index).Data(TradeSlot).Moveset(i).CurPP
                    .Moveset(i).TotalPP = PlayerPokemons(Index).Data(TradeSlot).Moveset(i).TotalPP
                Next
                .BallUsed = PlayerPokemons(Index).Data(TradeSlot).BallUsed
                .HeldItem = PlayerPokemons(Index).Data(TradeSlot).HeldItem
            Case Else
                '//Error
                Exit Sub
            End Select

            .TradeSlot = TradeSlot
            .Type = TradeType
        End With

        '//Update
        SendUpdateTradeItem Index, Index, slot
        SendUpdateTradeItem TempPlayer(Index).InTrade, Index, slot
    End If
End Sub

Private Sub HandleRemoveTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TradeSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InTrade <= 0 Then Exit Sub
    If TempPlayer(Index).TradeSet = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TradeSlot = buffer.ReadLong
    Set buffer = Nothing
    
    If TradeSlot <= 0 Or TradeSlot > MAX_TRADE Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index).TradeItem(TradeSlot)), LenB(TempPlayer(Index).TradeItem(TradeSlot)))
    '//Update
    SendUpdateTradeItem Index, Index, TradeSlot
    SendUpdateTradeItem TempPlayer(Index).InTrade, Index, TradeSlot
End Sub

Private Sub HandleTradeUpdateMoney(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim valMoney As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InTrade <= 0 Then Exit Sub
    If TempPlayer(Index).TradeSet = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    valMoney = buffer.ReadLong
    Set buffer = Nothing
    
    If valMoney > Player(Index, TempPlayer(Index).UseChar).Money Then
        TempPlayer(Index).TradeMoney = Player(Index, TempPlayer(Index).UseChar).Money
    Else
        TempPlayer(Index).TradeMoney = valMoney
    End If
    
    '//Update
    SendTradeUpdateMoney Index, Index
    SendTradeUpdateMoney TempPlayer(Index).InTrade, Index
End Sub

Private Sub HandleSetTradeState(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tSet As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InTrade <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tSet = buffer.ReadByte
    Set buffer = Nothing
    
    TempPlayer(Index).TradeSet = tSet
    If tSet = YES Then
        '//Set
        Select Case TempPlayer(TempPlayer(Index).InTrade).CurLanguage
            Case LANG_PT: AddAlert TempPlayer(Index).InTrade, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " set his trade", White
            Case LANG_EN: AddAlert TempPlayer(Index).InTrade, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " set his trade", White
            Case LANG_ES: AddAlert TempPlayer(Index).InTrade, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " set his trade", White
        End Select
    Else
        '//Cancel
        AddAlert TempPlayer(Index).InTrade, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " cancelled his trade", White
    End If
    SendSetTradeState Index, Index
    SendSetTradeState TempPlayer(Index).InTrade, Index
End Sub

Private Sub HandleTradeState(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tState As Byte
Dim i As Byte
Dim tradeIndex As Long
Dim pokemonCount(1 To 2) As Long, itemCount(1 To 2) As Long
Dim countPoke As Long, countItem As Long
Dim PokeSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If TempPlayer(Index).InTrade <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tState = buffer.ReadByte
    Set buffer = Nothing
    
    tradeIndex = TempPlayer(Index).InTrade
    If tState = 1 Then '//Accept
        TempPlayer(Index).TradeAccept = YES
        Select Case TempPlayer(tradeIndex).CurLanguage
            Case LANG_PT: AddAlert tradeIndex, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
            Case LANG_EN: AddAlert tradeIndex, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
            Case LANG_ES: AddAlert tradeIndex, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
        End Select
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
            Case LANG_EN: AddAlert Index, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
            Case LANG_ES: AddAlert Index, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " accepted the trade", White
        End Select
        '//Check if continue
        If TempPlayer(tradeIndex).TradeAccept = YES Then
            '//Continue trading their items
            pokemonCount(1) = 0
            itemCount(1) = 0
            pokemonCount(2) = 0
            itemCount(2) = 0
            For i = 1 To MAX_TRADE
                If TempPlayer(Index).TradeItem(i).Num > 0 Then
                    If TempPlayer(Index).TradeItem(i).Type = 1 Then '//Item
                        itemCount(1) = itemCount(1) + 1
                    ElseIf TempPlayer(Index).TradeItem(i).Type = 2 Then '//Pokemon
                        pokemonCount(1) = pokemonCount(1) + 1
                    End If
                End If
                If TempPlayer(tradeIndex).TradeItem(i).Num > 0 Then
                    If TempPlayer(tradeIndex).TradeItem(i).Type = 1 Then '//Item
                        itemCount(2) = itemCount(2) + 1
                    ElseIf TempPlayer(tradeIndex).TradeItem(i).Type = 2 Then '//Pokemon
                        pokemonCount(2) = pokemonCount(2) + 1
                    End If
                End If
            Next
            
            '//Check if player can receive item/pokemon
            countPoke = MAX_PLAYER_POKEMON - CountPlayerPokemon(tradeIndex)
            countItem = CountFreeInvSlot(tradeIndex)
            If countItem < itemCount(1) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You don't have enough inventory slot", White
                    Case LANG_EN: AddAlert tradeIndex, "You don't have enough inventory slot", White
                    Case LANG_ES: AddAlert tradeIndex, "You don't have enough inventory slot", White
                End Select
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Trade Error", White
                    Case LANG_EN: AddAlert Index, "Trade Error", White
                    Case LANG_ES: AddAlert Index, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            If countPoke < pokemonCount(1) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                    Case LANG_EN: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                    Case LANG_ES: AddAlert tradeIndex, "You don't have enough pokemon slot", White
                End Select
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Trade Error", White
                    Case LANG_EN: AddAlert Index, "Trade Error", White
                    Case LANG_ES: AddAlert Index, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            If CountPlayerPokemon(tradeIndex) <= pokemonCount(2) Then
                '//Error
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                    Case LANG_EN: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                    Case LANG_ES: AddAlert tradeIndex, "You can't trade all your available pokemon", White
                End Select
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Trade Error", White
                    Case LANG_EN: AddAlert Index, "Trade Error", White
                    Case LANG_ES: AddAlert Index, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            countPoke = MAX_PLAYER_POKEMON - CountPlayerPokemon(Index)
            countItem = CountFreeInvSlot(Index)
            If countItem < itemCount(2) Then
                '//Error
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You don't have enough inventory slot", White
                    Case LANG_EN: AddAlert Index, "You don't have enough inventory slot", White
                    Case LANG_ES: AddAlert Index, "You don't have enough inventory slot", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            If countPoke < pokemonCount(2) Then
                '//Error
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You don't have enough pokemon slot", White
                    Case LANG_EN: AddAlert Index, "You don't have enough pokemon slot", White
                    Case LANG_ES: AddAlert Index, "You don't have enough pokemon slot", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            If CountPlayerPokemon(Index) <= pokemonCount(1) Then
                '//Error
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "You can't trade all your available pokemon", White
                    Case LANG_EN: AddAlert Index, "You can't trade all your available pokemon", White
                    Case LANG_ES: AddAlert Index, "You can't trade all your available pokemon", White
                End Select
                Select Case TempPlayer(tradeIndex).CurLanguage
                    Case LANG_PT: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_EN: AddAlert tradeIndex, "Trade Error", White
                    Case LANG_ES: AddAlert tradeIndex, "Trade Error", White
                End Select
                TempPlayer(Index).TradeAccept = NO
                TempPlayer(tradeIndex).TradeAccept = NO
                TempPlayer(Index).TradeSet = 0
                TempPlayer(tradeIndex).TradeSet = 0
                SendSetTradeState Index, Index
                SendSetTradeState tradeIndex, Index
                Exit Sub
            End If
            
            '//Continue Trading
            For i = 1 To MAX_TRADE
                '//Give Items
                With TempPlayer(Index).TradeItem(i)
                    If .Type = 1 Then '//Item
                        Call TryGivePlayerItem(tradeIndex, PlayerInv(Index).Data(.TradeSlot).Num, .Value)
                    ElseIf .Type = 2 Then '//Pokemon
                        PokeSlot = FindOpenPokeSlot(tradeIndex)
                        If PokeSlot > 0 Then
                            Call CopyMemory(ByVal VarPtr(PlayerPokemons(tradeIndex).Data(PokeSlot)), ByVal VarPtr(PlayerPokemons(Index).Data(.TradeSlot)), LenB(PlayerPokemons(Index).Data(.TradeSlot)))
                            '//Add Pokedex
                            AddPlayerPokedex tradeIndex, PlayerPokemons(tradeIndex).Data(PokeSlot).Num, YES, YES
                            '//reupdate order
                            UpdatePlayerPokemonOrder tradeIndex
                            '//update
                            SendPlayerPokemons tradeIndex
                        End If
                    End If
                End With
                With TempPlayer(tradeIndex).TradeItem(i)
                    If .Type = 1 Then '//Item
                        Call TryGivePlayerItem(Index, PlayerInv(tradeIndex).Data(.TradeSlot).Num, .Value)
                    ElseIf .Type = 2 Then '//Pokemon
                        PokeSlot = FindOpenPokeSlot(Index)
                        If PokeSlot > 0 Then
                            Call CopyMemory(ByVal VarPtr(PlayerPokemons(Index).Data(PokeSlot)), ByVal VarPtr(PlayerPokemons(tradeIndex).Data(.TradeSlot)), LenB(PlayerPokemons(tradeIndex).Data(.TradeSlot)))
                            '//Add Pokedex
                            AddPlayerPokedex Index, PlayerPokemons(Index).Data(PokeSlot).Num, YES, YES
                            '//reupdate order
                            UpdatePlayerPokemonOrder Index
                            '//update
                            SendPlayerPokemons Index
                        End If
                    End If
                End With
                
                '//Take Items
                With TempPlayer(Index).TradeItem(i)
                    If .Type = 1 Then '//Item
                        PlayerInv(Index).Data(.TradeSlot).Value = PlayerInv(Index).Data(.TradeSlot).Value - .Value
                        If PlayerInv(Index).Data(.TradeSlot).Value <= 0 Then
                            PlayerInv(Index).Data(.TradeSlot).Num = 0
                            PlayerInv(Index).Data(.TradeSlot).Value = 0
                            PlayerInv(Index).Data(.TradeSlot).TmrCooldown = 0
                        End If
                        '//Update
                        SendPlayerInvSlot Index, .TradeSlot
                    ElseIf .Type = 2 Then '//Pokemon
                        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(Index).Data(.TradeSlot)), LenB(PlayerPokemons(Index).Data(.TradeSlot)))
                        '//reupdate order
                        UpdatePlayerPokemonOrder Index
                        '//update
                        SendPlayerPokemons Index
                    End If
                End With
                With TempPlayer(tradeIndex).TradeItem(i)
                    If .Type = 1 Then '//Item
                        '//Take item from index
                        PlayerInv(tradeIndex).Data(.TradeSlot).Value = PlayerInv(tradeIndex).Data(.TradeSlot).Value - .Value
                        If PlayerInv(tradeIndex).Data(.TradeSlot).Value <= 0 Then
                            PlayerInv(tradeIndex).Data(.TradeSlot).Num = 0
                            PlayerInv(tradeIndex).Data(.TradeSlot).Value = 0
                            PlayerInv(tradeIndex).Data(.TradeSlot).TmrCooldown = 0
                        End If
                        '//Update
                        SendPlayerInvSlot tradeIndex, .TradeSlot
                    ElseIf .Type = 2 Then '//Pokemon
                        Call ZeroMemory(ByVal VarPtr(PlayerPokemons(tradeIndex).Data(.TradeSlot)), LenB(PlayerPokemons(tradeIndex).Data(.TradeSlot)))
                        '//reupdate order
                        UpdatePlayerPokemonOrder tradeIndex
                        '//update
                        SendPlayerPokemons tradeIndex
                    End If
                End With
            Next
            
            If TempPlayer(Index).TradeMoney > 0 And TempPlayer(Index).TradeMoney <= Player(Index, TempPlayer(Index).UseChar).Money Then
                Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money + TempPlayer(Index).TradeMoney
                If Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money >= MAX_MONEY Then
                    Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = MAX_MONEY
                End If
                Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money - TempPlayer(Index).TradeMoney
                If Player(Index, TempPlayer(Index).UseChar).Money <= 0 Then
                    Player(Index, TempPlayer(Index).UseChar).Money = 0
                End If
                SendPlayerData Index
                SendPlayerData tradeIndex
            End If
            If TempPlayer(tradeIndex).TradeMoney > 0 And TempPlayer(tradeIndex).TradeMoney <= Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money Then
                Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money + TempPlayer(tradeIndex).TradeMoney
                If Player(Index, TempPlayer(Index).UseChar).Money >= MAX_MONEY Then
                    Player(Index, TempPlayer(Index).UseChar).Money = MAX_MONEY
                End If
                Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money - TempPlayer(tradeIndex).TradeMoney
                If Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money <= 0 Then
                    Player(tradeIndex, TempPlayer(tradeIndex).UseChar).Money = 0
                End If
                SendPlayerData Index
                SendPlayerData tradeIndex
            End If
            
            '//Exit Trade
            TempPlayer(Index).InTrade = 0
            TempPlayer(tradeIndex).InTrade = 0
            For i = 1 To MAX_TRADE
                Call ZeroMemory(ByVal VarPtr(TempPlayer(Index).TradeItem(i)), LenB(TempPlayer(Index).TradeItem(i)))
                Call ZeroMemory(ByVal VarPtr(TempPlayer(tradeIndex).TradeItem(i)), LenB(TempPlayer(tradeIndex).TradeItem(i)))
            Next
            TempPlayer(Index).TradeMoney = 0
            TempPlayer(tradeIndex).TradeMoney = 0
            TempPlayer(Index).TradeSet = 0
            TempPlayer(tradeIndex).TradeSet = 0
            TempPlayer(Index).TradeAccept = 0
            TempPlayer(tradeIndex).TradeAccept = 0
            TempPlayer(Index).PlayerRequest = 0
            TempPlayer(tradeIndex).PlayerRequest = 0
            TempPlayer(Index).RequestType = 0
            TempPlayer(tradeIndex).RequestType = 0
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Trade success", White
                Case LANG_EN: AddAlert Index, "Trade success", White
                Case LANG_ES: AddAlert Index, "Trade success", White
            End Select
            Select Case TempPlayer(tradeIndex).CurLanguage
                Case LANG_PT: AddAlert tradeIndex, "Trade success", White
                Case LANG_EN: AddAlert tradeIndex, "Trade success", White
                Case LANG_ES: AddAlert tradeIndex, "Trade success", White
            End Select
            SendCloseTrade tradeIndex
            SendCloseTrade Index
            SendRequest tradeIndex
            SendRequest Index
        End If
    Else '//Decline
        TempPlayer(Index).InTrade = 0
        TempPlayer(tradeIndex).InTrade = 0
        For i = 1 To MAX_TRADE
            Call ZeroMemory(ByVal VarPtr(TempPlayer(Index).TradeItem(i)), LenB(TempPlayer(Index).TradeItem(i)))
            Call ZeroMemory(ByVal VarPtr(TempPlayer(tradeIndex).TradeItem(i)), LenB(TempPlayer(tradeIndex).TradeItem(i)))
        Next
        TempPlayer(Index).TradeMoney = 0
        TempPlayer(tradeIndex).TradeMoney = 0
        TempPlayer(Index).TradeSet = 0
        TempPlayer(tradeIndex).TradeSet = 0
        TempPlayer(Index).TradeAccept = 0
        TempPlayer(tradeIndex).TradeAccept = 0
        TempPlayer(Index).PlayerRequest = 0
        TempPlayer(tradeIndex).PlayerRequest = 0
        TempPlayer(Index).RequestType = 0
        TempPlayer(tradeIndex).RequestType = 0
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "You have decline the trade", White
            Case LANG_EN: AddAlert Index, "You have decline the trade", White
            Case LANG_ES: AddAlert Index, "You have decline the trade", White
        End Select
        Select Case TempPlayer(tradeIndex).CurLanguage
            Case LANG_PT: AddAlert tradeIndex, "The trade was declined", White
            Case LANG_EN: AddAlert tradeIndex, "The trade was declined", White
            Case LANG_ES: AddAlert tradeIndex, "The trade was declined", White
        End Select
        SendCloseTrade tradeIndex
        SendCloseTrade Index
        SendRequest tradeIndex
        SendRequest Index
    End If
End Sub

Private Sub HandleScanPokedex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ScanType As Byte, ScanIndex As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ScanType = buffer.ReadByte
    ScanIndex = buffer.ReadLong
    Set buffer = Nothing
    
    If ScanType = 1 Then '//Map Npc
        If ScanIndex > 0 And ScanIndex <= MAX_GAME_POKEMON Then
            If MapPokemon(ScanIndex).Num > 0 Then
                If MapPokemon(ScanIndex).Map = Player(Index, TempPlayer(Index).UseChar).Map Then
                    '//Add Pokedex
                    AddPlayerPokedex Index, MapPokemon(ScanIndex).Num, , YES
                End If
            End If
        End If
    Else '// Player
        If ScanIndex > 0 And ScanIndex <= MAX_PLAYER Then
            If PlayerPokemon(ScanIndex).Num > 0 Then
                If Player(ScanIndex, TempPlayer(ScanIndex).UseChar).Map = Player(Index, TempPlayer(Index).UseChar).Map Then
                    '//Add Pokedex
                    AddPlayerPokedex Index, PlayerPokemon(ScanIndex).Num, , YES
                End If
            End If
        End If
    End If
End Sub

Private Sub HandleMOTD(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Text As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Text = Trim$(buffer.ReadString)
    Set buffer = Nothing
    
    Options.MOTD = Text
    SaveOption
    
    SendPlayerMsg Index, "MOTD was changed to: " & Text, White
End Sub

Private Sub HandleCopyMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim destinationMap As Long, sourceMap As Long
Dim i As Long
Dim CurRevision As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    destinationMap = buffer.ReadLong
    sourceMap = buffer.ReadLong
    Set buffer = Nothing
    If destinationMap <= 0 Or destinationMap > MAX_MAP Then Exit Sub
    If sourceMap <= 0 Or sourceMap > MAX_MAP Then Exit Sub
    
    CurRevision = Map(destinationMap).Revision
    Call CopyMemory(ByVal VarPtr(Map(destinationMap)), ByVal VarPtr(Map(sourceMap)), LenB(Map(sourceMap)))
    Map(destinationMap).Revision = CurRevision + 1
    
    '//Save the map
    Call SaveMap(destinationMap)
    Call Create_MapCache(destinationMap)
    
    '//Send the clear data first
    Call SendMapNpcData(destinationMap)
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(Index, TempPlayer(Index).UseChar).Map, i, NO, 0, 0, 0, Index
    Next
    '//Map Npc
    Call SpawnMapNpcs(destinationMap)
    
    '//Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                With Player(i, TempPlayer(i).UseChar)
                    If .Map = destinationMap Then
                        Call PlayerWarp(i, destinationMap, .X, .Y, .Dir)
                    End If
                End With
            End If
        End If
    Next
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " updated map#" & Player(Index, TempPlayer(Index).UseChar).Map
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " updated map#" & Player(Index, TempPlayer(Index).UseChar).Map
    Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Copying map complete", White
        Case LANG_EN: AddAlert Index, "Copying map complete", White
        Case LANG_ES: AddAlert Index, "Copying map complete", White
    End Select
End Sub

Private Sub HandleGiveItemTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim ItemNum As Long, ItemVal As Long
    Dim playerName As String
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerName = Trim$(buffer.ReadString)
    ItemNum = buffer.ReadLong
    ItemVal = buffer.ReadLong
    Set buffer = Nothing

    If UCase$(playerName) <> "ALL" Then
        i = FindPlayer(playerName)
    Else
        AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " , Admin Presenteou a todos: Give Item Item#" & ItemNum & " x" & ItemVal

        For i = 1 To Player_HighIndex
            Call TryGivePlayerItem(i, ItemNum, ItemVal)

            Select Case TempPlayer(i).CurLanguage
            Case LANG_PT: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            Case LANG_EN: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            Case LANG_ES: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            End Select
        Next i

        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Todos os jogadores foram presenteados!", White
        Case LANG_EN: AddAlert Index, "Todos os jogadores foram presenteados!", White
        Case LANG_ES: AddAlert Index, "Todos os jogadores foram presenteados!", White
        End Select
        Exit Sub
    End If

    If Not IsPlaying(i) Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Player is offline.", White
        Case LANG_EN: AddAlert Index, "Player is offline.", White
        Case LANG_ES: AddAlert Index, "Player is offline.", White
        End Select
        Exit Sub
    End If
    If TempPlayer(i).UseChar = 0 Then Exit Sub
    If ItemNum <= 0 Or ItemNum > MAX_ITEM Then Exit Sub
    If ItemVal <= 0 Then ItemVal = 1

    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " , Admin Rights: Give Item To " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & ", Item#" & ItemNum & " x" & ItemVal

    If Not TryGivePlayerItem(i, ItemNum, ItemVal) Then
        '//Error msg
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Player Inventory is full", White
        Case LANG_EN: AddAlert Index, "Player Inventory is full", White
        Case LANG_ES: AddAlert Index, "Player Inventory is full", White
        End Select
    Else
        Select Case TempPlayer(i).CurLanguage
        Case LANG_PT: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
        Case LANG_EN: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
        Case LANG_ES: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
        End Select
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Player receive the item", White
        Case LANG_EN: AddAlert Index, "Player receive the item", White
        Case LANG_ES: AddAlert Index, "Player receive the item", White
        End Select
    End If
End Sub

Private Sub HandleGivePokemonTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeNum As Long, Level As Long
Dim playerName As String
Dim i As Long
Dim IsShiny As Byte, IVFull As Byte, TheNature As Integer
Dim PokeBall As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerName = Trim$(buffer.ReadString)
    PokeNum = buffer.ReadLong
    Level = buffer.ReadLong
    IsShiny = buffer.ReadByte
    IVFull = buffer.ReadByte
    TheNature = buffer.ReadInteger
    PokeBall = buffer.ReadByte
    Set buffer = Nothing
    If UCase$(playerName) <> "ALL" Then
        i = FindPlayer(playerName)
    Else
        If PokeNum <= 0 Or PokeNum > MAX_POKEMON Then Exit Sub
        If Level <= 0 Or Level > MAX_LEVEL Then Exit Sub
        
        AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " , Admin Presenteou a todos: Give Poke#" & PokeNum & " Lvl:" & Level

        For i = 1 To Player_HighIndex
            GivePlayerPokemon i, PokeNum, Level, BallEnum.b_Pokeball, IsShiny, IVFull, TheNature
            
            Select Case TempPlayer(i).CurLanguage
            Case LANG_PT: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            Case LANG_EN: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            Case LANG_ES: AddAlert i, "You receive a item gift from " & Trim$(Player(Index, TempPlayer(Index).UseChar).Name), White
            End Select
        Next i
        
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Todos os jogadores foram presenteados!", White
            Case LANG_EN: AddAlert Index, "Todos os jogadores foram presenteados!", White
            Case LANG_ES: AddAlert Index, "Todos os jogadores foram presenteados!", White
            End Select
        
        Exit Sub
    End If
    If Not IsPlaying(i) Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "Player is offline.", White
        Case LANG_EN: AddAlert Index, "Player is offline.", White
        Case LANG_ES: AddAlert Index, "Player is offline.", White
        End Select
        Exit Sub
    End If
    If TempPlayer(i).UseChar = 0 Then Exit Sub
    If PokeNum <= 0 Or PokeNum > MAX_POKEMON Then Exit Sub
    If Level <= 0 Or Level > MAX_LEVEL Then Exit Sub
    
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " , Admin Rights: Give Pokemon To " & Trim$(Player(i, TempPlayer(i).UseChar).Name) & ", Pokemon#" & PokeNum & " Level" & Level
    GivePlayerPokemon i, PokeNum, Level, PokeBall, IsShiny, IVFull, TheNature
End Sub

Private Sub HandleSpawnPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapPokeSlot As Long, IsShiny As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapPokeSlot = buffer.ReadLong
    IsShiny = buffer.ReadByte
    Set buffer = Nothing
    If MapPokeSlot <= 0 Or MapPokeSlot > MAX_GAME_POKEMON Then Exit Sub
    
    ClearMapPokemon MapPokeSlot
    If IsShiny = YES Then
        SpawnMapPokemon MapPokeSlot, True, YES
    Else
        SpawnMapPokemon MapPokeSlot, True
    End If
    TempPlayer(Index).MapSwitchTmr = NO
End Sub

Private Sub HandleSetLanguage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TempPlayer(Index).CurLanguage = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleBuyStorageSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageType As Byte, StorageSlot As Byte
Dim Amount As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageType = buffer.ReadByte
    StorageSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If StorageSlot < 0 Or StorageSlot > MAX_STORAGE_SLOT Then Exit Sub
    Amount = 100000 * (StorageSlot - 2)
    If Player(Index, TempPlayer(Index).UseChar).Money >= Amount Then
        Select Case StorageType
            Case 1 '// Item
                If PlayerInvStorage(Index).slot(StorageSlot).Unlocked = NO Then
                    PlayerInvStorage(Index).slot(StorageSlot).Unlocked = YES
                    Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money - Amount
                    SendPlayerInvStorage Index
                    SendPlayerData Index
                    AddAlert Index, "New Item Storage slot has been unlocked", White
                End If
            Case 2 '// Pokemon
                If PlayerPokemonStorage(Index).slot(StorageSlot).Unlocked = NO Then
                    PlayerPokemonStorage(Index).slot(StorageSlot).Unlocked = YES
                    Player(Index, TempPlayer(Index).UseChar).Money = Player(Index, TempPlayer(Index).UseChar).Money - Amount
                    SendPlayerPokemonStorage Index
                    SendPlayerData Index
                    AddAlert Index, "New Pokemon Storage slot has been unlocked", White
                End If
        End Select
    Else
        AddAlert Index, "Not enough money", White
    End If
End Sub

Private Sub HandleSellPokeStorageSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim StorageSlot As Byte, StorageData As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
  
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    StorageSlot = buffer.ReadByte
    StorageData = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleChangeShinyRate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Rate As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Rate = buffer.ReadLong
    Set buffer = Nothing
    
    SendPlayerMsg Index, "Shiny Rate was changed to: " & Rate, White
End Sub

Private Sub HandleRelearnMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveSlot As Byte, PokeSlot As Byte, PokeNum As Long
Dim MoveNum As Long, oSlot As Byte
Dim X As Byte
Dim InvSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If PlayerPokemon(Index).Num <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MoveSlot = buffer.ReadByte
    PokeSlot = buffer.ReadByte
    PokeNum = buffer.ReadLong
    Set buffer = Nothing
    
    If Not PlayerPokemon(Index).Num = PokeNum Then Exit Sub
    If Not PlayerPokemon(Index).slot = PokeSlot Then Exit Sub
    If MoveSlot <= 0 Or MoveSlot > MAX_POKEMON_MOVESET Then Exit Sub
    
    If PokeNum > 0 Then
        MoveNum = Pokemon(PokeNum).Moveset(MoveSlot).MoveNum
        If MoveNum > 0 Then
            For X = 1 To MAX_MOVESET
                If MoveNum = PlayerPokemons(Index).Data(PokeSlot).Moveset(X).Num Then
                    Exit Sub
                End If
            Next
            If PlayerPokemons(Index).Data(PokeSlot).Level < Pokemon(PokeNum).Moveset(MoveSlot).MoveLevel Then
                Exit Sub
            End If
            InvSlot = FindInvItemSlot(Index, 72)
            '//Check if have the required item
            If InvSlot > 0 Then
                '//Take Item
                PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - 1
                If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
                    '//Clear Item
                    PlayerInv(Index).Data(InvSlot).Num = 0
                    PlayerInv(Index).Data(InvSlot).Value = 0
                    PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
                End If
                SendPlayerInvSlot Index, InvSlot
            
                '//Continue
                oSlot = FindFreeMoveSlot(Index, PokeSlot)
                If oSlot > 0 Then
                    PlayerPokemons(Index).Data(PokeSlot).Moveset(oSlot).Num = MoveNum
                    PlayerPokemons(Index).Data(PokeSlot).Moveset(oSlot).TotalPP = PokemonMove(MoveNum).PP
                    PlayerPokemons(Index).Data(PokeSlot).Moveset(oSlot).CurPP = PlayerPokemons(Index).Data(PokeSlot).Moveset(MoveSlot).TotalPP
                    SendPlayerPokemonSlot Index, PokeSlot
                    '//Send Msg
                    SendPlayerMsg Index, Trim$(Pokemon(PokeNum).Name) & " learned the move " & Trim$(PokemonMove(MoveNum).Name), White
                Else
                    '//Proceed to ask
                    TempPlayer(Index).MoveLearnPokeSlot = PokeSlot
                    TempPlayer(Index).MoveLearnNum = MoveNum
                    TempPlayer(Index).MoveLearnIndex = 0
                    SendNewMove Index
                End If
            Else
                Select Case TempPlayer(Index).CurLanguage
                    Case LANG_PT: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                    Case LANG_EN: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                    Case LANG_ES: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(72).Name), White
                End Select
            End If
        End If
    End If
End Sub

Private Sub HandleUseRevive(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim IsMaxRev As Byte, ReviveItemNum As Long
Dim InvSlot As Long, PokeSlot As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadByte
    IsMaxRev = buffer.ReadByte
    Set buffer = Nothing
    
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).CurHp > 0 Then Exit Sub
    If TempPlayer(Index).InDuel > 0 Then Exit Sub
    If TempPlayer(Index).InNpcDuel > 0 Then Exit Sub
    
    If IsMaxRev = YES Then
        ReviveItemNum = 48
        InvSlot = FindInvItemSlot(Index, ReviveItemNum)
        If InvSlot > 0 Then
            '//Take Item
            PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - 1
            If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
                '//Clear Item
                PlayerInv(Index).Data(InvSlot).Num = 0
                PlayerInv(Index).Data(InvSlot).Value = 0
                PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
            End If
            SendPlayerInvSlot Index, InvSlot
            
            PlayerPokemons(Index).Data(PokeSlot).CurHp = PlayerPokemons(Index).Data(PokeSlot).MaxHp
            SendPlayerPokemonSlot Index, PokeSlot
            
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon was revived", White
                Case LANG_EN: AddAlert Index, "Pokemon was revived", White
                Case LANG_ES: AddAlert Index, "Pokemon was revived", White
            End Select
        Else
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_EN: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_ES: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
            End Select
        End If
    Else
        ReviveItemNum = 62
        InvSlot = FindInvItemSlot(Index, ReviveItemNum)
        If InvSlot > 0 Then
            '//Take Item
            PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - 1
            If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
                '//Clear Item
                PlayerInv(Index).Data(InvSlot).Num = 0
                PlayerInv(Index).Data(InvSlot).Value = 0
                PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
            End If
            SendPlayerInvSlot Index, InvSlot
            
            PlayerPokemons(Index).Data(PokeSlot).CurHp = PlayerPokemons(Index).Data(PokeSlot).MaxHp / 2
            SendPlayerPokemonSlot Index, PokeSlot
            
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Pokemon was revived", White
                Case LANG_EN: AddAlert Index, "Pokemon was revived", White
                Case LANG_ES: AddAlert Index, "Pokemon was revived", White
            End Select
        Else
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_EN: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
                Case LANG_ES: AddAlert Index, "Sorry, you don't have a " & Trim$(Item(ReviveItemNum).Name), White
            End Select
        End If
    End If
End Sub

Private Sub HandleAddHeld(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim InvSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If InvSlot <= 0 Or InvSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInv(Index).Data(InvSlot).Num <= 0 Then Exit Sub
    If PlayerInv(Index).Data(InvSlot).Value < 1 Then Exit Sub
    
    ' Item não pode ser um held para o pokemon
    If Item(PlayerInv(Index).Data(InvSlot).Num).RestrictionData.CanHold = YES Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Não equipável pelo seu pokémon", White
            Case LANG_EN: AddAlert Index, "Not equippable by your pokemon", White
            Case LANG_ES: AddAlert Index, "No equipable por tu pokemon", White
        End Select
        Exit Sub
    End If
    If PlayerPokemon(Index).Num <= 0 Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Spawn your pokemon first", White
            Case LANG_EN: AddAlert Index, "Spawn your pokemon first", White
            Case LANG_ES: AddAlert Index, "Spawn your pokemon first", White
        End Select
        Exit Sub
    End If
    If PlayerPokemon(Index).slot <= 0 Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Spawn your pokemon first", White
            Case LANG_EN: AddAlert Index, "Spawn your pokemon first", White
            Case LANG_ES: AddAlert Index, "Spawn your pokemon first", White
        End Select
        Exit Sub
    End If
    
    ' check previous held item
    If PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).HeldItem > 0 Then
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Your pokemon is currently holding an item", White
            Case LANG_EN: AddAlert Index, "Your pokemon is currently holding an item", White
            Case LANG_ES: AddAlert Index, "Your pokemon is currently holding an item", White
        End Select
        Exit Sub
    End If
    
    ' Complete Process
    ' Give item
    PlayerPokemons(Index).Data(PlayerPokemon(Index).slot).HeldItem = PlayerInv(Index).Data(InvSlot).Num
    SendPlayerPokemonSlot Index, PlayerPokemon(Index).slot
    
    '//Take Item
    PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - 1
    If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
        '//Clear Item
        PlayerInv(Index).Data(InvSlot).Num = 0
        PlayerInv(Index).Data(InvSlot).Value = 0
        PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
    End If
    SendPlayerInvSlot Index, InvSlot
End Sub

Private Sub HandleRemoveHeld(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Long
Dim ItemNum As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If PokeSlot <= 0 Or PokeSlot > MAX_PLAYER_POKEMON Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).Num <= 0 Then Exit Sub
    If PlayerPokemons(Index).Data(PokeSlot).HeldItem <= 0 Then Exit Sub
    
    ItemNum = PlayerPokemons(Index).Data(PokeSlot).HeldItem
    
    If TryGivePlayerItem(Index, ItemNum, 1) Then
        PlayerPokemons(Index).Data(PokeSlot).HeldItem = 0
        SendPlayerPokemonSlot Index, PokeSlot
    End If
End Sub

Private Sub HandleStealthMode(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    If Player(Index, TempPlayer(Index).UseChar).StealthMode = YES Then
        Player(Index, TempPlayer(Index).UseChar).StealthMode = NO
    Else
        Player(Index, TempPlayer(Index).UseChar).StealthMode = YES
    End If
    SendPlayerData Index
End Sub

Private Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendWhosOnline Index
End Sub

Private Sub HandleRequestRank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendRankTo Index
    
    UpdateRank Trim$(Player(Index, TempPlayer(Index).UseChar).Name), Player(Index, TempPlayer(Index).UseChar).Level, Player(Index, TempPlayer(Index).UseChar).CurExp
End Sub

Private Sub HandleHotbarUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim HotbarSlot As Byte, InvSlot As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    HotbarSlot = buffer.ReadByte
    InvSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If HotbarSlot <= 0 Or HotbarSlot > MAX_HOTBAR Then Exit Sub
    
    With Player(Index, TempPlayer(Index).UseChar)
        If InvSlot > 0 And InvSlot <= MAX_PLAYER_INV Then
            If PlayerInv(Index).Data(InvSlot).Num > 0 Then
                .Hotbar(HotbarSlot) = PlayerInv(Index).Data(InvSlot).Num
            Else
                .Hotbar(HotbarSlot) = 0
            End If
        Else
            .Hotbar(HotbarSlot) = 0
        End If
    End With
    SendPlayerData Index
End Sub

Private Sub HandleUseHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim HotbarSlot As Byte, InvSlot As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    HotbarSlot = buffer.ReadByte
    Set buffer = Nothing
    
    If HotbarSlot <= 0 Or HotbarSlot > MAX_HOTBAR Then Exit Sub
 
    With Player(Index, TempPlayer(Index).UseChar)
        If .Hotbar(HotbarSlot) > 0 Then
            InvSlot = checkItem(Index, .Hotbar(HotbarSlot))
            
            If InvSlot > 0 Then
                '//Use Item
                PlayerUseItem Index, InvSlot
            End If
        End If
    End With
End Sub

Private Sub HandleCreateParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    CreateParty Index
End Sub

Private Sub HandleLeaveParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    LeaveParty Index
End Sub

'//Editors
Private Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MAPPER Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit map#" & Player(Index, TempPlayer(Index).UseChar).Map
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit map#" & Player(Index, TempPlayer(Index).UseChar).Map

    Set buffer = New clsBuffer
    buffer.WriteLong SInitMap
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MapNum As Long
Dim X As Long, Y As Long
Dim i As Long, a As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MAPPER Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNum = Player(Index, TempPlayer(Index).UseChar).Map
    Call ClearMap(MapNum)
    
    With Map(MapNum)
        '//General
        .Revision = buffer.ReadLong
        .Name = Trim$(buffer.ReadString)
        .Moral = buffer.ReadByte
        
        '//Size
        .MaxX = buffer.ReadLong
        .MaxY = buffer.ReadLong
        If .MaxX < MAX_MAPX Then .MaxX = MAX_MAPX
        If .MaxY < MAX_MAPY Then .MaxY = MAX_MAPY
        
        '//Redim size
        ReDim Map(MapNum).Tile(0 To .MaxX, 0 To .MaxY)
    End With
    
    '//Tiles
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(X, Y)
                '//Layer
                For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                    For a = MapLayerType.Normal To MapLayerType.Animated
                        .Layer(i, a).Tile = buffer.ReadLong
                        .Layer(i, a).TileX = buffer.ReadLong
                        .Layer(i, a).TileY = buffer.ReadLong
                        '//Map Anim
                        .Layer(i, a).MapAnim = buffer.ReadLong
                    Next
                Next
                '//Tile Data
                .Attribute = buffer.ReadByte
                .Data1 = buffer.ReadLong
                .Data2 = buffer.ReadLong
                .Data3 = buffer.ReadLong
                .Data4 = buffer.ReadLong
            End With
        Next
    Next
    
    With Map(MapNum)
        '//Map Link
        .LinkUp = buffer.ReadLong
        .LinkDown = buffer.ReadLong
        .LinkLeft = buffer.ReadLong
        .LinkRight = buffer.ReadLong
        
        '//Map Data
        .Music = Trim$(buffer.ReadString)
        
        '//Npc
        For i = 1 To MAX_MAP_NPC
            .Npc(i) = buffer.ReadLong
            ClearMapNpc MapNum, i
        Next
        
        '//Moral
        .KillPlayer = buffer.ReadByte
        .IsCave = buffer.ReadByte
        .CaveLight = buffer.ReadByte
        .SpriteType = buffer.ReadByte
        .StartWeather = buffer.ReadByte
        
        .NoCure = buffer.ReadByte
    End With
    Set buffer = Nothing
    
    '//Save the map
    Call SaveMap(MapNum)
    Call Create_MapCache(MapNum)
    
    '//Send the clear data first
    Call SendMapNpcData(MapNum)
    For i = 1 To MAX_MAP_NPC
        SendNpcPokemonData Player(Index, TempPlayer(Index).UseChar).Map, i, NO, 0, 0, 0, Index
    Next
    '//Map Npc
    Call SpawnMapNpcs(MapNum)
    
    '//Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).UseChar > 0 Then
                With Player(i, TempPlayer(i).UseChar)
                    If .Map = MapNum Then
                        Call PlayerWarp(i, MapNum, .X, .Y, .Dir)
                    End If
                End With
            End If
        End If
    Next
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " updated map#" & Player(Index, TempPlayer(Index).UseChar).Map
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " updated map#" & Player(Index, TempPlayer(Index).UseChar).Map
End Sub

Private Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit npc"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit npc"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitNpc
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested npc data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested npc data"
    
    SendNpcs Index
End Sub

Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_NPC Then Exit Sub
    dSize = LenB(Npc(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Npc(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save npc#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save npc#" & xIndex & " data"
    
    Call SendUpdateNpcToAll(xIndex)
    Call SaveNpc(xIndex)
End Sub

Private Sub HandleRequestEditPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Pokemon"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Pokemon"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitPokemon
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Pokemon data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Pokemon data"
    
    SendPokemons Index
End Sub

Private Sub HandleSavePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_POKEMON Then Exit Sub
    dSize = LenB(Pokemon(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Pokemon(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save pokemon#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save pokemon#" & xIndex & " data"
    
    Call SendUpdatePokemonToAll(xIndex)
    Call SavePokemon(xIndex)
End Sub

Private Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Item"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Item"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitItem
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Item data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Item data"
    
    SendItems Index
End Sub

Private Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_ITEM Then Exit Sub
    dSize = LenB(Item(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Item(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save item#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save item#" & xIndex & " data"
    
    Call SendUpdateItemToAll(xIndex)
    Call SaveItem(xIndex)
End Sub

Private Sub HandleRequestEditPokemonMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Pokemon Move"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Pokemon Move"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitPokemonMove
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestPokemonMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Pokemon Move data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Pokemon Move data"
    
    SendPokemonMoves Index
End Sub

Private Sub HandleSavePokemonMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_POKEMON_MOVE Then Exit Sub
    dSize = LenB(PokemonMove(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(PokemonMove(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save pokemonmove#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save pokemonmove#" & xIndex & " data"
    
    Call SendUpdatePokemonMoveToAll(xIndex)
    Call SavePokemonMove(xIndex)
End Sub

Private Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Animation"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Animation"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitAnimation
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Animation data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Animation data"
    
    SendAnimations Index
End Sub

Private Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_ANIMATION Then Exit Sub
    dSize = LenB(Animation(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Animation(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save animation#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save animation#" & xIndex & " data"
    
    Call SendUpdateAnimationToAll(xIndex)
    Call SaveAnimation(xIndex)
End Sub

Private Sub HandleRequestEditSpawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Spawn"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Spawn"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitSpawn
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestSpawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Spawn data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Spawn data"
    
    SendSpawns Index
End Sub

Private Sub HandleSaveSpawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Spawn(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Spawn(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Spawn#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Spawn#" & xIndex & " data"
    
    Call SendUpdateSpawnToAll(xIndex)
    Call SaveSpawn(xIndex)
    
    '//Update Data
    ClearMapPokemon xIndex, True
    MapPokemon(xIndex).PokemonIndex = Spawn(xIndex).PokeNum
    MapPokemon(xIndex).Respawn = GetTickCount
    '//Update HighIndex
    If Spawn(xIndex).PokeNum > 0 Then
        If xIndex > Pokemon_HighIndex Then
            Pokemon_HighIndex = xIndex
            SendPokemonHighIndex
        End If
    End If
End Sub

Private Sub HandleRequestEditConversation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub
    
    Call SendConversations(Index)

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Conversation"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Conversation"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitConversation
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestConversation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Conversation data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Conversation data"
    
    SendConversations Index
End Sub

Private Sub HandleSaveConversation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Conversation(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Conversation(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Conversation#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Conversation#" & xIndex & " data"
    
    Call SendUpdateConversationToAll(xIndex)
    Call SaveConversation(xIndex)
End Sub

Private Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Shop"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Shop"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitShop
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Shop data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Shop data"
    
    SendShops Index
End Sub

Private Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Shop(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Shop(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Shop#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Shop#" & xIndex & " data"
    
    Call SendUpdateShopToAll(xIndex)
    Call SaveShop(xIndex)
End Sub

Private Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Quest"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested to edit Quest"

    Set buffer = New clsBuffer
    buffer.WriteLong SInitQuest
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Private Sub HandleRequestQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Quest data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " requested Quest data"
    
    SendQuests Index
End Sub

Private Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim xIndex As Long
Dim dSize As Long
Dim dData() As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    xIndex = buffer.ReadLong
    If xIndex < 0 Or xIndex > MAX_GAME_POKEMON Then Exit Sub
    dSize = LenB(Quest(xIndex))
    ReDim dData(dSize - 1)
    dData = buffer.ReadBytes(dSize)
    CopyMemory ByVal VarPtr(Quest(xIndex)), ByVal VarPtr(dData(0)), dSize
    Set buffer = Nothing
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Quest#" & xIndex & " data"
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " save Quest#" & xIndex & " data"
    
    Call SendUpdateQuestToAll(xIndex)
    Call SaveQuest(xIndex)
End Sub

Private Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " kick " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " kick " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    CloseSocket TargetIndex
End Sub

Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    ' Banir o IP
    BanIP GetPlayerIP(TargetIndex)
    ' Banir o Character
    BanCharacter Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    CloseSocket TargetIndex
End Sub

Private Sub HandleMutePlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Muted = YES
End Sub

Private Sub HandleUnmutePlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Long
Dim InputName As String

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    '//Check access
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InputName = buffer.ReadString
    Set buffer = Nothing
    
    TargetIndex = FindPlayer(InputName)
    
    If TargetIndex <= 0 Or TargetIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(TargetIndex) Then Exit Sub
    If TempPlayer(TargetIndex).UseChar <= 0 Then Exit Sub
    If Player(Index, TempPlayer(Index).UseChar).Access <= Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Access Then Exit Sub
    
    TextAdd frmServer.txtLog, Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " ban " & Trim$(Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Name)
    
    Player(TargetIndex, TempPlayer(TargetIndex).UseChar).Muted = NO
End Sub

Private Sub HandleFlyToBadge(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim TargetIndex As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TargetIndex = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleRequestCash(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim FindP As Integer, YesNo As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    FindP = FindPlayer(buffer.ReadString)
    YesNo = buffer.ReadByte
    Set buffer = Nothing

    If FindP = 0 Then
        SendPlayerMsg Index, "Jogador Offline!", BrightRed
        Exit Sub
    End If
    
    If YesNo = YES Then
        SendRequestCash Index, FindP
    Else
        SendRequestCash Index, FindP, False
    End If
End Sub

Private Sub HandleSetCash(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim FindP As Integer, YesNo As Byte, Value As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_CREATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    FindP = FindPlayer(buffer.ReadString)
    YesNo = buffer.ReadByte
    Value = buffer.ReadLong
    Set buffer = Nothing

    If FindP = 0 Then
        Select Case TempPlayer(Index).CurLanguage
        Case LANG_PT: AddAlert Index, "O jogador está offline.", BrightRed
        Case LANG_EN: AddAlert Index, "O jogador está offline.", BrightRed
        Case LANG_ES: AddAlert Index, "O jogador está offline.", BrightRed
        End Select
        Exit Sub
    Else

        If YesNo = YES Then
            On Error GoTo TrataErro1
            If Value >= MAX_CASH Or Player(FindP, TempPlayer(FindP).UseChar).Cash + Value >= MAX_CASH Then
                Player(FindP, TempPlayer(FindP).UseChar).Cash = MAX_CASH
            Else
                Player(FindP, TempPlayer(FindP).UseChar).Cash = Player(FindP, TempPlayer(FindP).UseChar).Cash + Value
            End If

            Select Case TempPlayer(FindP).CurLanguage
            Case LANG_PT: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Cash!", White
            Case LANG_EN: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Cash!", White
            Case LANG_ES: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Cash!", White
            End Select
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "O jogador recebeu " & Value & " Cash!", White
            Case LANG_EN: AddAlert Index, "O jogador recebeu " & Value & " Cash!", White
            Case LANG_ES: AddAlert Index, "O jogador recebeu " & Value & " Cash!", White
            End Select

            Call SendRequestCash(Index, FindP, True)
            
            Call SendPlayerCash(FindP)
            Exit Sub
            
                        
TrataErro1:
        Player(FindP, TempPlayer(FindP).UseChar).Cash = MAX_CASH
        Else
            On Error GoTo TrataErro2
            If Value >= MAX_MONEY Or Player(FindP, TempPlayer(FindP).UseChar).Money + Value >= MAX_MONEY Then
                Player(FindP, TempPlayer(FindP).UseChar).Money = MAX_MONEY
            Else
                Player(FindP, TempPlayer(FindP).UseChar).Money = Player(FindP, TempPlayer(FindP).UseChar).Money + Value
            End If
            
            Select Case TempPlayer(FindP).CurLanguage
            Case LANG_PT: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Money!", White
            Case LANG_EN: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Money!", White
            Case LANG_ES: AddAlert FindP, "Congratulations! Your Reiceved " & Value & Space(1) & "Money!", White
            End Select
            Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "O jogador recebeu " & Value & " Money!", White
            Case LANG_EN: AddAlert Index, "O jogador recebeu " & Value & " Money!", White
            Case LANG_ES: AddAlert Index, "O jogador recebeu " & Value & " Money!", White
            End Select
            
            Call SendRequestCash(Index, FindP, False)
            
            Call SendPlayerCash(FindP)
            Exit Sub
            
TrataErro2:
        Player(FindP, TempPlayer(FindP).UseChar).Money = MAX_MONEY
        End If
    End If


End Sub

Private Sub HandleRequestServerInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    If IsPlaying(Index) Then Exit Sub
    'If IsConnected(Index) Then Exit Sub

    Call SendRequestServerInfo(Index)
    
    'Call CloseSocket(Index)
End Sub

Private Sub HandleBuyInvSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvSlot As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvSlot = buffer.ReadByte
    Set buffer = Nothing
    
    Call BuyInvSlot(Index, InvSlot)
End Sub

Private Sub HandleRequestVirtualShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim FindP As Integer, YesNo As Byte

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    SendVirtualShopTo Index
End Sub

Private Sub HandlePurchaseVirtualShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim slot As Long, Indice As Long

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Indice = buffer.ReadLong
    slot = buffer.ReadLong
    Set buffer = Nothing

    '//Proteção pra não estourar o índice geral
    If Indice > VirtualShopTabsRec.CountTabs - 1 Or Indice < VirtualShopTabsRec.Skins Then Exit Sub

    '//Verificação se está em uma matriz valida
    If slot > VirtualShop(Indice).Max_Slots Or slot <= 0 Then Exit Sub

    '//Verificação se existe um item neste slot
    If VirtualShop(Indice).Items(slot).ItemNum <= 0 Or VirtualShop(Indice).Items(slot).ItemNum > MAX_ITEM Then Exit Sub

    '//Verificação se esse slot tem um valor alienado
    If VirtualShop(Indice).Items(slot).ItemPrice <= 0 Or VirtualShop(Indice).Items(slot).ItemPrice > MAX_CASH Then Exit Sub

    '//Verificar se o item não tem uma quantidade especificada na database
    If VirtualShop(Indice).Items(slot).ItemQuant <= 0 Then Exit Sub

    '//Verificação se o jogador possui o valor de cash pra comprar o item
    With Player(Index, TempPlayer(Index).UseChar)
        If .Cash < VirtualShop(Indice).Items(slot).ItemPrice Then Exit Sub

        '//Tentativa de entrega do item, e retirada do cash.
        If TryGivePlayerItem(Index, VirtualShop(Indice).Items(slot).ItemNum, VirtualShop(Indice).Items(slot).ItemQuant) Then
            .Cash = .Cash - VirtualShop(Indice).Items(slot).ItemPrice
            Call SendPlayerCash(Index)
            
            '//Enviar uma mensagem, que tudo ocorreu com sucesso.
            Select Case TempPlayer(Index).CurLanguage
                Case LANG_PT: AddAlert Index, "Parabens, você acaba de receber um item!", White
                Case LANG_EN: AddAlert Index, "Congratulations, You received a item!", White
                Case LANG_ES: AddAlert Index, "Congratulations, You received a item!", White
            End Select
        End If
    End With
End Sub

Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    
    If Player(Index, TempPlayer(Index).UseChar).Access < ACCESS_MODERATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapReport
    For i = 1 To MAX_MAP
        buffer.WriteString Trim$(Map(i).Name)
    Next
    SendDataTo Index, buffer.ToArray()

    buffer.Flush: Set buffer = Nothing
End Sub















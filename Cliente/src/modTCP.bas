Attribute VB_Name = "modTCP"
Option Explicit

' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Public PlayerBuffer As clsBuffer

Public Sub TcpInit()
    Set PlayerBuffer = New clsBuffer
    '//Set the connection settings
    LoadServerList CurServerList
    
    Call ResetServerInfo
    Call RequestServerInfo
End Sub

Public Sub LoadServerList(ByVal ServerSlot As Integer)
    If ServerSlot <= 0 Or ServerSlot > MAX_SERVER_LIST Then Exit Sub
    
    DestroyTCP
    frmMain.Socket.RemoteHost = ServerIP(ServerSlot)
    frmMain.Socket.RemotePort = ServerPort(ServerSlot)
End Sub

Public Sub DestroyTCP()
    '//Close socket
    frmMain.Socket.close
End Sub

Public Function IsPlaying(ByVal Index As Long) As Boolean
    If Len(Trim$(Player(Index).Name)) > 0 Then
        IsPlaying = True
    End If
End Function

'//This function start the connection between server and client
Public Function ConnectToServer() As Boolean
Dim Wait As Long
    
    '//Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.Socket.close
    frmMain.Socket.Connect
    
    '//Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected
End Function

'//Check if Socket is connected
Public Function IsConnected() As Boolean
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Public Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer

    If IsConnected Then
        Set buffer = New clsBuffer
        buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
End Sub

'//Packets
Public Sub CheckPing()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
    '//Start Ping Timer
    PingStart = GetTickCount
End Sub

Public Sub SendNewAccount(ByVal Username As String, ByVal Password As String, ByVal Email As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString Trim$(Username)
    buffer.WriteString Trim$(Password)
    buffer.WriteString Trim$(Email)
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    buffer.WriteLong APP_MAJOR
    buffer.WriteLong APP_MINOR
    buffer.WriteLong APP_REVISION
    buffer.WriteString ProcessorID
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendLoginInfo(ByVal Username As String, ByVal Password As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CLoginInfo
    buffer.WriteString Trim$(Username)
    buffer.WriteString Trim$(Password)
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    buffer.WriteLong APP_MAJOR
    buffer.WriteLong APP_MINOR
    buffer.WriteLong APP_REVISION
    buffer.WriteString ProcessorID
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendNewCharacter(ByVal CharName As String, ByVal Gender As Byte, ByVal CharSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CNewCharacter
    buffer.WriteString Trim$(CharName)
    buffer.WriteByte Gender
    buffer.WriteByte CharSlot
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    buffer.WriteLong APP_MAJOR
    buffer.WriteLong APP_MINOR
    buffer.WriteLong APP_REVISION
    buffer.WriteString ProcessorID
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseCharacter(ByVal CharSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseCharacter
    buffer.WriteByte CharSlot
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    buffer.WriteLong APP_MAJOR
    buffer.WriteLong APP_MINOR
    buffer.WriteLong APP_REVISION
    buffer.WriteString ProcessorID
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDelCharacter(ByVal CharSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDelCharacter
    buffer.WriteByte CharSlot
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    buffer.WriteLong APP_MAJOR
    buffer.WriteLong APP_MINOR
    buffer.WriteLong APP_REVISION
    buffer.WriteString ProcessorID
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteByte Player(MyIndex).Dir
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteByte Player(MyIndex).Dir
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

'//Chat
Public Sub SendMapMsg(ByVal Msg As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CMapMsg
    buffer.WriteString Msg
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendGlobalMsg(ByVal Msg As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGlobalMsg
    buffer.WriteString Msg
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendPartyMsg(ByVal Msg As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPartyMsg
    buffer.WriteString Msg
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendPlayerMsg(ByVal rcName As String, ByVal Msg As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMsg
    buffer.WriteString rcName
    buffer.WriteString Msg
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendWarpTo(ByVal MapNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong MapNum
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWarpToMe(ByVal rcName As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString rcName
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendWarpMeTo(ByVal rcName As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString rcName
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal rcName As String, ByVal inAccess As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString rcName
    buffer.WriteByte inAccess
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonMove()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerPokemonMove
    buffer.WriteByte PlayerPokemon(MyIndex).Dir
    buffer.WriteLong PlayerPokemon(MyIndex).X
    buffer.WriteLong PlayerPokemon(MyIndex).Y
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonDir()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerPokemonDir
    buffer.WriteByte PlayerPokemon(MyIndex).Dir
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGetItem(ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGetItem
    buffer.WriteLong ItemNum
    buffer.WriteLong ItemVal
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerPokemonState(ByVal State As Byte, ByVal PokeSlot As Byte)
Dim buffer As clsBuffer

    If GUI(GuiEnum.GUI_INPUTBOX).Visible Then Exit Sub
    If ChatOn Then Exit Sub
    If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then Exit Sub
    If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then Exit Sub
    If GUI(GuiEnum.GUI_MOVEREPLACE).Visible Then Exit Sub
    If GettingMap Then Exit Sub
    If Player(MyIndex).Action > 0 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerPokemonState
    buffer.WriteByte State
    buffer.WriteByte PokeSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendChangePassword(ByVal NewPassword As String, ByVal OldPassword As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CChangePassword
    buffer.WriteString NewPassword
    buffer.WriteString OldPassword
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendReplaceNewMove(ByVal MoveSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CReplaceNewMove
    buffer.WriteByte MoveSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendEvolvePoke(ByVal EvolveSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CEvolvePoke
    buffer.WriteByte EvolveSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal ItemSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteByte ItemSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSwitchInvSlot(ByVal OldSlot As Byte, ByVal NewSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchInvSlot
    buffer.WriteByte OldSlot
    buffer.WriteByte NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGotData(Optional ByVal InUsed As Byte = 0, Optional ByVal Data1 As Long = 0, Optional ByVal Data2 As Long = 0, Optional ByVal Data3 As Long = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGotData
    buffer.WriteByte InUsed
    buffer.WriteLong Data1
    buffer.WriteLong Data2
    buffer.WriteLong Data3
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendOpenStorage(ByVal StorageType As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong COpenStorage
    buffer.WriteByte StorageType
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDepositItemTo(ByVal StorageSlot As Byte, ByVal StorageData As Byte, ByVal InvSlot As Byte, Optional ByVal wValue As Long = 1)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItemTo
    buffer.WriteByte StorageSlot
    'buffer.WriteByte StorageData
    buffer.WriteByte InvSlot
    buffer.WriteLong wValue
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSwitchStorageSlot(ByVal StorageSlot As Byte, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchStorageSlot
    buffer.WriteByte StorageSlot
    buffer.WriteByte OldSlot
    buffer.WriteByte NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWithdrawItemTo(ByVal StorageSlot As Byte, ByVal StorageData As Byte, ByVal InvSlot As Byte, Optional ByVal wValue As Long = 1)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItemTo
    buffer.WriteByte StorageSlot
    buffer.WriteByte StorageData
    'buffer.WriteByte InvSlot
    buffer.WriteLong wValue
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendConvo(ByVal ctype As Byte, ByVal Data1 As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CConvo
    buffer.WriteByte ctype
    buffer.WriteLong Data1
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendProcessConvo(Optional ByVal tReply As Byte = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CProcessConvo
    buffer.WriteByte tReply
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendDepositPokemon(ByVal StorageSlot As Byte, ByVal PokeSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CDepositPokemon
    buffer.WriteByte StorageSlot
    buffer.WriteByte PokeSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWithdrawPokemon(ByVal StorageSlot As Byte, ByVal StorageData As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawPokemon
    buffer.WriteByte StorageSlot
    buffer.WriteByte StorageData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSwitchStoragePokeSlot(ByVal StorageSlot As Byte, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchStoragePokeSlot
    buffer.WriteByte StorageSlot
    buffer.WriteByte OldSlot
    buffer.WriteByte NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSwitchStoragePoke(ByVal OldPokeSlot As Byte, ByVal PokemonNewStorage As Byte)
Dim buffer As clsBuffer

    If PokemonCurSlot = PokemonNewStorage Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchStoragePoke
    buffer.WriteByte OldPokeSlot        ' last poke slot
    buffer.WriteByte PokemonCurSlot     ' storage
    buffer.WriteByte PokemonNewStorage  ' storage
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSwitchStorageItem(ByVal OldItemSlot As Byte, ByVal ItemNewStorage As Byte)
Dim buffer As clsBuffer

    If InvCurSlot = ItemNewStorage Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchStorageItem
    buffer.WriteByte OldItemSlot        ' last poke slot
    buffer.WriteByte InvCurSlot     ' storage
    buffer.WriteByte ItemNewStorage  ' storage
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBuyItem(ByVal ShopSlot As Byte, Optional ByVal ShopVal As Long = 1)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteByte ShopSlot
    buffer.WriteLong ShopVal
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSellItem(ByVal InvSlot As Byte, Optional ByVal InvVal As Long = 1)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteByte InvSlot
    buffer.WriteLong InvVal
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCloseShop()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCloseShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequest(ByVal requestIndex As Long, ByVal RequestType As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequest
    buffer.WriteByte RequestType
    buffer.WriteLong requestIndex
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestState(ByVal RequestState As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestState
    buffer.WriteByte RequestState
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendAddTrade(ByVal TradeType As Byte, ByVal TradeSlot As Long, Optional ByVal TradeData As Long = 1)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAddTrade
    buffer.WriteByte TradeType
    buffer.WriteLong TradeSlot
    buffer.WriteLong TradeData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRemoveTrade(ByVal TradeSlot As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRemoveTrade
    buffer.WriteLong TradeSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendTradeUpdateMoney(ByVal valMoney As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeUpdateMoney
    buffer.WriteLong valMoney
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetTradeState(ByVal tState As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetTradeState
    buffer.WriteByte tState
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendTradeState(ByVal tState As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeState
    buffer.WriteByte tState
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendScanPokedex(ByVal ScanType As Byte, ByVal ScanIndex As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong CScanPokedex
    buffer.WriteByte ScanType
    buffer.WriteLong ScanIndex
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMOTD(ByVal Text As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CMOTD
    buffer.WriteString Text
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCopyMap(ByVal destinationMap As Long, sourceMap As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCopyMap
    buffer.WriteLong destinationMap
    buffer.WriteLong sourceMap
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendReleasePokemon(ByVal StorageSlot As Byte, ByVal StorageData As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CReleasePokemon
    buffer.WriteByte StorageSlot
    buffer.WriteByte StorageData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGiveItemTo(ByVal PlayerName As String, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGiveItemTo
    buffer.WriteString PlayerName
    buffer.WriteLong ItemNum
    buffer.WriteLong ItemVal
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendGivePokemonTo(ByVal PlayerName As String, ByVal PokeNum As Long, _
                             ByVal Level As Long, Optional ByVal IsShiny As Byte = NO, _
                             Optional IVFull As Byte = NO, Optional Nature As Integer = -1, _
                             Optional PokeBall As Byte = 0)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CGivePokemonTo
    buffer.WriteString PlayerName
    buffer.WriteLong PokeNum
    buffer.WriteLong Level
    buffer.WriteByte IsShiny
    buffer.WriteByte IVFull
    buffer.WriteInteger Nature
    buffer.WriteByte PokeBall
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSpawnPokemon(ByVal MapPokeSlot As Long, ByVal IsShiny As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnPokemon
    buffer.WriteLong MapPokeSlot
    buffer.WriteByte IsShiny
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSetLanguage()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSetLanguage
    '//Send Version
    buffer.WriteByte GameSetting.CurLanguage
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBuyStorageSlot(ByVal StorageType As Byte, ByVal StorageSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBuyStorageSlot
    buffer.WriteByte StorageType
    buffer.WriteByte StorageSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSellPokeStorageSlot(ByVal StorageSlot As Byte, StorageData As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSellPokeStorageSlot
    buffer.WriteByte StorageSlot
    buffer.WriteByte StorageData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendChangeShinyRate(ByVal Rate As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CChangeShinyRate
    buffer.WriteLong Rate
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRelearnMove(ByVal MoveSlot As Byte, ByVal PokeSlot As Byte, ByVal PokeNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRelearnMove
    buffer.WriteByte MoveSlot
    buffer.WriteByte PokeSlot
    buffer.WriteLong PokeNum
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseRevive(ByVal Slot As Byte, ByVal IsMaxRev As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseRevive
    buffer.WriteByte Slot
    buffer.WriteByte IsMaxRev
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SenAddHeld(ByVal InvSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CAddHeld
    buffer.WriteByte InvSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendStealthMode()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CStealthMode
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestRank()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestRank
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub


Public Sub SendHotbarUpdate(ByVal HotbarSlot As Byte, Optional ByVal InvSlot As Byte = 0)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUpdate
    buffer.WriteByte HotbarSlot
    buffer.WriteByte InvSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUseHotbar(ByVal HotbarSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CUseHotbar
    buffer.WriteByte HotbarSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCreateParty()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CCreateParty
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendLeaveParty()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CLeaveParty
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

'//Editors
Public Sub SendRequestEditMap()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_MAPPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMap()
Dim buffer As clsBuffer
Dim X As Long, Y As Long
Dim i As Long, a As Byte

    If Player(MyIndex).Access < ACCESS_MAPPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CMap
    
    With Map
        '//General
        buffer.WriteLong .Revision
        buffer.WriteString Trim$(.Name)
        buffer.WriteByte .Moral
        
        '//Size
        buffer.WriteLong .MaxX
        buffer.WriteLong .MaxY
    End With
    
    '//Tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                '//Layer
                For i = MapLayer.Ground To MapLayer.MapLayer_Count - 1
                    For a = MapLayerType.Normal To MapLayerType.Animated
                        buffer.WriteLong .Layer(i, a).Tile
                        buffer.WriteLong .Layer(i, a).TileX
                        buffer.WriteLong .Layer(i, a).TileY
                        '//Map Anim
                        buffer.WriteLong .Layer(i, a).MapAnim
                    Next
                Next
                '//Tile Data
                buffer.WriteByte .Attribute
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteLong .Data4
            End With
        Next
    Next
    
    With Map
        '//Map Link
        buffer.WriteLong .LinkUp
        buffer.WriteLong .LinkDown
        buffer.WriteLong .LinkLeft
        buffer.WriteLong .LinkRight
        
        '//Map Data
        buffer.WriteString .Music
        
        '//Npc
        For i = 1 To MAX_MAP_NPC
            buffer.WriteLong .Npc(i)
        Next
        
        '//Moral
        buffer.WriteByte .KillPlayer
        buffer.WriteByte .IsCave
        buffer.WriteByte .CaveLight
        buffer.WriteByte .SpriteType
        buffer.WriteByte .StartWeather
        buffer.WriteByte .NoCure
    End With
    
    SendData buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestNpc()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Npc(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Npc(xIndex)), dSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditPokemon()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditPokemon
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestPokemon()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPokemon
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSavePokemon(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Pokemon(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Pokemon(xIndex)), dSize
    buffer.WriteLong CSavePokemon
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestItem()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItem
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte


    Set buffer = New clsBuffer
    dSize = LenB(Item(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Item(xIndex)), dSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditPokemonMove()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditPokemonMove
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestPokemonMove()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPokemonMove
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSavePokemonMove(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(PokemonMove(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(PokemonMove(xIndex)), dSize
    buffer.WriteLong CSavePokemonMove
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestAnimation()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Animation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Animation(xIndex)), dSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditSpawn()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpawn
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestSpawn()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpawn
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveSpawn(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Spawn(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Spawn(xIndex)), dSize
    buffer.WriteLong CSaveSpawn
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditConversation()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditConversation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestConversation()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestConversation
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveConversation(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Conversation(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Conversation(xIndex)), dSize
    buffer.WriteLong CSaveConversation
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestShop()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Shop(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Shop(xIndex)), dSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestEditQuest()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditQuest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRequestQuest()
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestQuest
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendSaveQuest(ByVal xIndex As Long)
Dim buffer As clsBuffer
Dim dSize As Long
Dim dData() As Byte

    Set buffer = New clsBuffer
    dSize = LenB(Quest(xIndex))
    ReDim dData(dSize - 1)
    CopyMemory dData(0), ByVal VarPtr(Quest(xIndex)), dSize
    buffer.WriteLong CSaveQuest
    buffer.WriteLong xIndex
    buffer.WriteBytes dData
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendKickPlayer(ByVal sName As String)
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString sName
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendBanPlayer(ByVal sName As String)
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString sName
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendRemoveHeld(ByVal PokeSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRemoveHeld
    buffer.WriteByte PokeSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMutePlayer(ByVal sName As String)
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CMutePlayer
    buffer.WriteString sName
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendUnmutePlayer(ByVal sName As String)
Dim buffer As clsBuffer

    If Player(MyIndex).Access < ACCESS_MODERATOR Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong CUnmutePlayer
    buffer.WriteString sName
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendFlyToBadge(ByVal BadgeSlot As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CFlyToBadge
    buffer.WriteByte BadgeSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub


Public Sub SendRequestPlayerValue(ByVal Name As String, Optional ByVal IsCash As Boolean = YES)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestCash
    buffer.WriteString Name
    If IsCash Then buffer.WriteByte YES Else buffer.WriteByte NO
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCashValueTo(ByVal Name As String, ByVal value As Long, Optional ByVal IsCash As Boolean = YES)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CSetCash
    buffer.WriteString Name
    If IsCash Then buffer.WriteByte YES Else buffer.WriteByte NO
    buffer.WriteLong value
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub RequestServerInfo()
    Dim buffer As clsBuffer

    ' Resetar antes.
    Call ResetServerInfo
    
    ' Realizar solicitação de informações de jogadores online!
    If ConnectToServer Then
        Set buffer = New clsBuffer
        buffer.WriteLong CRequestServerInfo
        SendData buffer.ToArray()
        Set buffer = Nothing
    End If
End Sub

Public Sub SendBuyInvSlot(ByVal InvSlot As Byte)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CBuyInvSlot
    buffer.WriteByte InvSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub RequestVirtualShop()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CRequestVirtualShop
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PurchaseVirtualShop(ByVal Indice As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CPurchaseVirtualShop
    buffer.WriteLong Indice
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendMapReport()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

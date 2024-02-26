Attribute VB_Name = "ItemController"
Public Sub PlayerUseItem(ByVal Index As Long, ByVal InvSlot As Byte)
    Dim ItemNum As Long
    Dim gothealed As Boolean
    Dim x As Long
    Dim exproll As Long
    Dim Exp As Long
    Dim i As Long, CanLearn As Boolean
    Dim BerriesFunc As Integer, PokeName As String
    Dim TAKE As Boolean

    If Not IsPlaying(Index) Then Exit Sub
    If TempPlayer(Index).UseChar <= 0 Then Exit Sub
    If InvSlot <= 0 Or InvSlot > MAX_PLAYER_INV Then Exit Sub
    If PlayerInv(Index).Data(InvSlot).Num <= 0 Then Exit Sub
    If PlayerInv(Index).Data(InvSlot).Value <= 0 Then Exit Sub
    If TempPlayer(Index).InDuel > 0 Then Exit Sub
    If TempPlayer(Index).InNpcDuel > 0 Then Exit Sub

    ItemNum = PlayerInv(Index).Data(InvSlot).Num

    ' Verificar se o item está em cooldown
    If CheckItemCooldown(Index, InvSlot) = False Then
    
        Select Case TempPlayer(Index).CurLanguage
            Case LANG_PT: AddAlert Index, "Item em cooldown, aguarde: " & GetItemCooldownSegs(Index, InvSlot), White
            Case LANG_EN: AddAlert Index, "Item on cooldown, please wait: " & GetItemCooldownSegs(Index, InvSlot), White
            Case LANG_ES: AddAlert Index, "Artículo en tiempo de reutilización, espera: " & GetItemCooldownSegs(Index, InvSlot), White
        End Select
        
        Call SendPlayerInvSlot(Index, InvSlot)
        Exit Sub
    End If

    Select Case Item(ItemNum).Category
    
        Case ItemCategoryEnum.None
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.PokeBall
            If RequerimentOk() Then
                UsePokeBall
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Medicine
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Protein
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Key
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Skills
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Bracelet
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case ItemCategoryEnum.Gacha
            If RequerimentOk() Then
            End If
            Exit Sub
            
        Case Else
            AlertPT = "Você tentou usar um item inexistente, um log foi registrado."
            AlertEN = "Intentó utilizar un elemento inexistente y se registró un registro."
            AlertES = "You tried to use a non-existent item, a log was recorded."
            
            AlertToPlayer Index, AlertPT, AlertEN, AlertES
            AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & "Tentou usar um item inexiste!"
            Exit Sub
            
    End Select

    '//Set Cooldown
    PlayerInv(Index).Data(InvSlot).TmrCooldown = GetTickCount + Item(ItemNum).Delay

    If TAKE = True Then
        '//Take Item
        PlayerInv(Index).Data(InvSlot).Value = PlayerInv(Index).Data(InvSlot).Value - 1
        If PlayerInv(Index).Data(InvSlot).Value <= 0 Then
            '//Clear Item
            PlayerInv(Index).Data(InvSlot).Num = 0
            PlayerInv(Index).Data(InvSlot).Value = 0
            PlayerInv(Index).Data(InvSlot).TmrCooldown = 0
        End If
    End If


    SendPlayerInvSlot Index, InvSlot

    AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & " use item " & Trim$(Item(ItemNum).Name)
End Sub
Public Function RequerimentsOk() As Boolean
    RequerimentsOk = True
End Function

Public Sub UsePokeBallController(ByVal PlayerIndex As Long, ByVal ItemIndex As Long)
    
    Select Case Map(Player(Index, TempPlayer(Index).UseChar).Map).Moral
        
        Case MapMoral.Danger
        Case MapMoral.Safe
        Case MapMoral.Arena
            AlertPT = "Você não pode usar este item aqui!."
            AlertEN = "You cannot use that item here!"
            AlertES = "¡No puedes usar ese artículo aquí!"
            
            AlertToPlayer Index, AlertPT, AlertEN, AlertES
            Exit Sub
        
        Case MapMoral.Safari
        Case Else
            AlertPT = "Você tentou fazer algo que não deveria! Um log foi gerado!."
            AlertEN = "You tried to do something you shouldn't! A log has been generated!"
            AlertES = "¡Intentaste hacer algo que no debías! ¡Se ha generado un registro!"
            
            AlertToPlayer Index, AlertPT, AlertEN, AlertES
            AddLog Trim$(Player(Index, TempPlayer(Index).UseChar).Name) & "Tentou usar um item inexiste!"
            Exit Sub
        
    End Select
    
    TempPlayer(Index).TmpUseInvSlot = InvSlot
    
    SendGetData Index, ItemCategoryEnum.PokeBall, InvSlot
End Sub

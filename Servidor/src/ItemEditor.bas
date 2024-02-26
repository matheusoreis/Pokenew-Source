Attribute VB_Name = "ItemEditor"
Public EditorItemIndex As Long

' Editor de Itens
Private ItemHasModified(1 To MAX_ITEM) As Boolean

Public Sub ItensEditorInit()
    Dim Index As Long, Chance As Double
    
    
    With frmEditor_Itens
        EditorItemIndex = .listIndex.listIndex + 1
        
        .textName.Text = Trim$(Item(EditorItemIndex).Name)
        
        .labelSprite.Caption = "Sprite: " & Item(EditorItemIndex).SpriteID
        .scrollSprite.Max = MAX_SPRITE_ITENS
        .scrollSprite.Value = Item(EditorItemIndex).SpriteID

        .comboRarity.listIndex = Item(EditorItemIndex).Rarity
        
        .comboCategory.listIndex = Item(EditorItemIndex).Category
        
        .textDescription.Text = Trim$(Item(EditorItemIndex).Description)
        
        .textCooldown.Text = Trim$(Item(EditorItemIndex).CooldownData.Value)
        .comboCooldown.listIndex = Item(EditorItemIndex).CooldownData.Type
        
        .checkRestriction(1).Value = BooleanToByte(Item(EditorItemIndex).RestrictionData.CanStack)
        .checkRestriction(2).Value = BooleanToByte(Item(EditorItemIndex).RestrictionData.CanHold)
        .checkRestriction(3).Value = BooleanToByte(Item(EditorItemIndex).RestrictionData.IsConnected)
        .checkRestriction(4).Value = BooleanToByte(Item(EditorItemIndex).RestrictionData.IsAdminItem)
        
        .labelPokemonLevel.Caption = "Level: " & Item(EditorItemIndex).PokemonRequirementData.RequiredLevel
        .scrollPokemonLevel.Value = Item(EditorItemIndex).PokemonRequirementData.RequiredLevel
        .comboPrimaryType.listIndex = Item(EditorItemIndex).PokemonRequirementData.PrimaryType
        .comboSecondaryType.listIndex = Item(EditorItemIndex).PokemonRequirementData.SecondaryType
        
        .labelPlayerLevel.Caption = "Level: " & Item(EditorItemIndex).PlayerRequirementData.RequiredLevel
        .scrollPlayerLevel.Value = Item(EditorItemIndex).PlayerRequirementData.RequiredLevel
        
        .comboMap.Clear
        .comboMap.AddItem "Nenhum"
        .comboMap.listIndex = 0
        If .comboMap.ListCount >= 0 Then
            For i = 1 To MAX_MAP
                .comboMap.AddItem (Trim$(Map(i).Name))
            Next
        End If
        
        .listMaps.Clear
        For i = 1 To MAX_MAPS_REQUIREMENTS
            If Item(EditorItemIndex).PlayerRequirementData.RequiredMaps(i) > 0 Then
                .listMaps.AddItem i & ": " & Map(Item(EditorItemIndex).PlayerRequirementData.RequiredMaps(i)).Name
            Else
                .listMaps.AddItem i & ": Nada"
            End If
        Next
        
        .comboBadge.listIndex = Item(EditorItemIndex).PlayerRequirementData.RequiredBadge
    
    End With
    
    Call ItensEditorLoadCategory
    Call ItensEditorLoadCategoryKey
    
End Sub

Public Sub ItensEditorSave()
    SaveItem EditorItemIndex
End Sub

Public Sub ItensEditorClear()
End Sub

Public Sub ItensEditorLoadCategory()
    With frmEditor_Itens
        Select Case Item(EditorItemIndex).Category
            Case ItemCategoryEnum.None
            
            Case ItemCategoryEnum.PokeBall
                .scrollChancePokeball.Value = Item(EditorItemIndex).PokeballData.CaptureChance
                .scrollSpritePokeball.Value = Item(EditorItemIndex).PokeballData.SpriteID
                .checkPerfectCapture.Value = BooleanToByte(Item(EditorItemIndex).PokeballData.HasPerfectCapture)
                
            Case ItemCategoryEnum.Medicine
                .comboTypeCure.listIndex = Item(EditorItemIndex).MedicineData.Type
                .scrollCureValue = Item(EditorItemIndex).MedicineData.Value
                .checkLevelUp.Value = BooleanToByte(Item(EditorItemIndex).MedicineData.HasLeveledUp)
                
            Case ItemCategoryEnum.Protein
                .comboProteinType.listIndex = Item(EditorItemIndex).ProteinsData.Type
                .scrollProteinValue.Value = Item(EditorItemIndex).ProteinsData.Value
                
            Case ItemCategoryEnum.Key
                .comboTypeKey.listIndex = Item(EditorItemIndex).KeyData.Type
                'Call ItensEditorLoadCategoryKey
                
                
            Case ItemCategoryEnum.Skills
                .comboMovesSkills.listIndex = Item(EditorItemIndex).SkillsData.Type
                .checkConsumeSkills.Value = Item(EditorItemIndex).SkillsData.CanConsume
                
            Case ItemCategoryEnum.Bracelet
                .comboTypeBracelet.listIndex = Item(EditorItemIndex).BraceletData.Type
                .textValueBracelet.Text = Trim$(Item(EditorItemIndex).BraceletData.Value)
                
            Case ItemCategoryEnum.Gacha
            
                .comboItensGachaBox.Clear
                
                .comboItensGachaBox.AddItem "Sem Itens"
                .comboItensGachaBox.listIndex = 0
                
                If .comboItensGachaBox.ListCount >= 0 Then
                    For i = 1 To MAX_ITEM
                        .comboItensGachaBox.AddItem (Trim$(Item(i).Name))
                    Next
                End If
                
                .listItensGachaBox.Clear
                For i = 1 To MAX_MYSTERY_BOX
                    If Item(EditorItemIndex).GachaData(i).ItemValue > 0 Then
                        .listItensGachaBox.AddItem i & Item(EditorItemIndex).GachaData(i).ItemValue & "x - " & Trim$(Item(Item(EditorItemIndex).GachaData(i).ItemValue).Name) & " " & Item(EditorItemIndex).GachaData(i).ItemChance & "%"
                        Chance = Chance + Item(EditorItemIndex).GachaData(i).ItemValue
                    Else
                        .listItensGachaBox.AddItem i & ": Sem Itens"
                    End If
                Next
                .labelTotalChanceGachaBox = "Chance total: " & Chance & "%"
                .labelMissingChanceGachaBox = "Faltam: " & (100 - Chance) & "%"
                .listItensGachaBox.listIndex = 0
        End Select
    End With
End Sub

Public Sub ItensEditorLoadCategoryKey()
    With frmEditor_Itens
        Select Case Item(EditorItemIndex).KeyData.Type
            Case KeyTypeEnum.None
            
            Case KeyTypeEnum.Sprite
                .scrollSpriteSkin.Value = Item(EditorItemIndex).KeyData.Sprite
                .textBonusExperience.Value = Item(EditorItemIndex).KeyData.ExperienceBonusAmount
                .scrollBonusMoney.Value = Item(EditorItemIndex).KeyData.MoneyBonusAmount
                .checkShiftKey.Value = BooleanToByte(Item(EditorItemIndex).KeyData.IsShiftRunning)
            
            Case KeyTypeEnum.OpenBank
            
            Case KeyTypeEnum.OpenComputer
            
        End Select
    End With
End Sub

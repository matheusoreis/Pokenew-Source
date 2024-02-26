Attribute VB_Name = "Inventory_Window"
Public Sub DrawInventory()
    Dim i As Long
    Dim DrawX As Long, DrawY As Long
    Dim Sprite As Long, Alpha As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Buttons
        'Dim ButtonText As String, DrawText As Boolean
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
            End If
        Next

        '//Items
        For i = 1 To MAX_PLAYER_INV
            If i <> DragInvSlot Then
                If PlayerInv(i).Status.Locked = NO Then
                    If PlayerInv(i).Num > 0 Then
                        Sprite = Item(PlayerInv(i).Num).SpriteID
                        
                        If PlayerInv(i).ItemCooldown > 0 Then
                            Alpha = D3DColorARGB(100, 255, 100, 100)
                        Else
                            Alpha = D3DColorARGB(255, 255, 255, 255)
                        End If

                        DrawX = .X + (7 + ((5 + TILE_X) * (((i - 1) Mod 5))))
                        DrawY = .Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 5)))

                        '//Draw Icon
                        If Sprite > 0 And Sprite <= Count_Item Then
                            RenderTexture Tex_Item(Sprite), DrawX + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), Alpha
                        End If

                        RenderTexture Tex_System(gSystemEnum.UserInterface), DrawX, DrawY, 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(20, 0, 0, 0)

                        '//Count
                        If PlayerInv(i).Value > 1 Then
                            RenderText Font_Default, PlayerInv(i).Value, DrawX + 28 - (GetTextWidth(Font_Default, PlayerInv(i).Value)), DrawY + 14, White
                        End If
                    End If
                Else '= YES
                    '//Renderizando slots bloqueados
                    Alpha = D3DColorARGB(PlayerInv(i).Status.Opacity, 255, 255, 255)
                    
                    Sprite = 532
                    DrawX = .X + (7 + ((5 + TILE_X) * (((i - 1) Mod 5))))
                    DrawY = .Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 5)))
                    RenderTexture Tex_Item(Sprite), DrawX + ((32 / 2) - (GetPicWidth(Tex_Item(Sprite)) / 2)), DrawY + ((32 / 2) - (GetPicHeight(Tex_Item(Sprite)) / 2)), 0, 0, GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), GetPicWidth(Tex_Item(Sprite)), GetPicHeight(Tex_Item(Sprite)), Alpha
                End If
            End If
        Next
    End With
End Sub

' ***************
' ** Inventory **
' ***************
Public Sub InventoryMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INVENTORY
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        If Not SelMenu.Visible And InvUseSlot = 0 Then
            If Buttons = vbRightButton Then
                '//Inv
                i = IsInvItem(CursorX, CursorY)
                If i > 0 Then
                    OpenSelMenu SelMenuType.Inv, i
                End If
            Else
                '//Disable Drag when intrade
                If TradeIndex = 0 Then
                    '//Inv
                    i = IsInvItem(CursorX, CursorY)
                    If i > 0 Then
                        DragInvSlot = i
                        WindowPriority = GuiEnum.GUI_INVENTORY
                    End If
                    
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub InventoryMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpX As Long, tmpY As Long
    Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
        Else
            Exit Sub
        End If

        If DragInvSlot > 0 Or DragStorageSlot > 0 Then
            If WindowPriority = 0 Then
                WindowPriority = GuiEnum.GUI_INVENTORY
                If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then
                    UpdateGuiOrder GUI_INVENTORY
                End If
            End If
        End If

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover

                        IsHovering = True
                        MouseIcon = 1    '//Select
                    End If
                End If
            End If
        Next

        '//Inv
        i = IsInvItem(CursorX, CursorY)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1    '//Select

            If Not InvItemDesc = i Then
                InvItemDesc = i
                InvItemDescTimer = GetTickCount
                InvItemDescShow = False
            End If
        End If

        i = IsInvSlot(CursorX, CursorY)
        If i > 0 Then
            If PlayerInv(i).Status.Opacity = 255 Then
                IsHovering = True
                MouseIcon = 1    '//Select
            End If
        End If


        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY

            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height

            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Public Sub InventoryMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                        Case ButtonEnum.Inventory_Close
                            If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                GuiState GUI_INVENTORY, False
                            End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Replace item
        If TradeIndex = 0 Then
            If DragInvSlot > 0 Then
                i = IsInvSlot(CursorX, CursorY)
                If i > 0 Then
                    If PlayerInv(i).Status.Locked = NO Then
                        SendSwitchInvSlot DragInvSlot, i
                    End If
                End If
            End If


            If DragInvSlot = 0 Then
                i = IsInvSlot(CursorX, CursorY)
                If i > 0 Then
                    If PlayerInv(i).Status.Opacity = 255 Then
                        BuySlotData = i
                        OpenChoiceBox TextUIChoiceBuyInvSlot, CB_BUYINV
                    End If
                End If
            End If

            DragInvSlot = 0
        End If

        '//Replace item
        If DragStorageSlot > 0 Then
            i = IsInvSlot(CursorX, CursorY)
            If i > 0 Then
                '//Check if value is greater than 1
                If PlayerInvStorage(InvCurSlot).data(DragStorageSlot).Value > 1 Then
                    If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                        OpenInputBox TextUIInputAmountHeader, IB_WITHDRAW, DragStorageSlot, i
                    End If
                Else
                    '//Send Withdraw
                    SendWithdrawItemTo InvCurSlot, DragStorageSlot, i
                End If
            End If
        End If
        DragStorageSlot = 0

        '//Check for dragging
        .InDrag = False
    End With
End Sub

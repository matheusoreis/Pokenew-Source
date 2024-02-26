Attribute VB_Name = "LojaVirtual_Window"
Option Explicit

'//Virtual Shop Scrolling
Public Const VirtualShopScrollSize As Byte = 35
Public Const VirtualShopViewLine As Byte = 2    'Quantidade por linha
Public Const VirtualShopViewLines = 4    'Quantidade de linhas
Public Const VirtualShopScrollLength As Byte = 176 - VirtualShopScrollSize
Public Const VirtualShopScrollStartY As Byte = 75 - VirtualShopScrollSize - 5
Public Const VirtualShopScrollEndY As Byte = 216

' Virtual Shop Scrooling
Public VirtualShopScrollCount As Long
Public VirtualShopScrollHold As Boolean
Public VirtualShopScroll As Long
Public VirtualShopScrollY As Long
Public VirtualShopMaxViewLine As Long
Public VirtualShopScrollUp As Boolean
Public VirtualShopScrollDown As Boolean
Public VirtualShopScrollTimer As Long

'//Window
Public VirtalShopSlot_State() As ButtonState

'//Dados
Public Enum VirtualShopTabsRec
    Skins = 1
    Mounts
    Items
    Vips

    CountTabs
End Enum

Public VirtualShop(1 To CountTabs - 1) As VirtualShopDataRec

Private Type VirtualShopRec
    SelectedItem As Boolean
    ItemNum As Long
    ItemQuant As Long
    ItemPrice As Long
    CustomDesc As Byte
End Type
Private Type VirtualShopDataRec
    SelectedWindow As Boolean
    Items() As VirtualShopRec
End Type

Private Function SetState_VirtualShopSlot(ByVal Slot As Long, State As ButtonState)

    VirtalShopSlot_State(Slot) = State

End Function

Public Function ResetState_VirtualShopSlot()
    Dim Y As Long
    
    On Error GoTo Error
    
    For Y = 1 To UBound(VirtalShopSlot_State)
        VirtalShopSlot_State(Y) = ButtonState.StateNormal
    Next Y
    
Error:
    Exit Function
End Function

Private Function GetIndexFromVirtualShop() As Byte
    Dim i As Long
    For i = 1 To CountTabs - 1
        If VirtualShop(i).SelectedWindow = True Then
            GetIndexFromVirtualShop = i
            Exit Function
        End If
    Next i
End Function

Private Function GetSlotSelectedFromVirtualShop() As Long
    Dim i As Long
    Dim VirtualShopIndex As Long

    VirtualShopIndex = GetIndexFromVirtualShop

    If VirtualShopIndex = 0 Then Exit Function

    For i = LBound(VirtualShop(VirtualShopIndex).Items) To UBound(VirtualShop(VirtualShopIndex).Items)
        If VirtualShop(VirtualShopIndex).Items(i).SelectedItem = True Then
            GetSlotSelectedFromVirtualShop = i
        End If
    Next i
End Function

Private Sub SetSlotSelectedFromVirtualShop(ByVal Slot As Long)
    Dim i As Long
    Dim VirtualShopIndex As Long

    VirtualShopIndex = GetIndexFromVirtualShop

    If VirtualShopIndex = 0 Then Exit Sub

    For i = LBound(VirtualShop(VirtualShopIndex).Items) To UBound(VirtualShop(VirtualShopIndex).Items)
        If i = Slot Then
            VirtualShop(VirtualShopIndex).Items(i).SelectedItem = True
        Else
            VirtualShop(VirtualShopIndex).Items(i).SelectedItem = False
        End If
    Next i
End Sub

Public Sub SwitchTabFromVirtualShop(ByVal NewTab As VirtualShopTabsRec)
    Dim i As Byte

    For i = 1 To CountTabs - 1
        If i = NewTab Then
            VirtualShop(i).SelectedWindow = True
            VirtualShopMaxViewLine = (UBound(VirtualShop(i).Items)) + 1
            
            VirtualShopScrollCount = 0
            VirtualShopScrollY = VirtualShopScrollLength
            
            Erase VirtalShopSlot_State
            ReDim VirtalShopSlot_State(1 To UBound(VirtualShop(i).Items))
        Else
            VirtualShop(i).SelectedWindow = False
        End If
    Next i
End Sub


Private Function PlayerHaveCashValue(ByVal Price As Long) As Boolean
    PlayerHaveCashValue = False

    If Player(MyIndex).Cash >= Price Then
        PlayerHaveCashValue = True
    End If
End Function

Public Sub DrawVirtualShop()
    Dim i As Long
    Dim tmpX As Long, tmpY As Long, Width As Integer, Height As Integer
    Dim ColourOpacity As Long
    Dim X As Long, z As Long, Y As Long, tmpS As Integer, XX As Long, YY As Long
    Dim CaptionBuy As String, DescString As String, CaptionContribuitor As String
    Dim VirtualShopIndex As Long, VirtualShopSlot As Long
    Dim ArrayText() As String, ArrayText2() As String

    With GUI(GuiEnum.GUI_VIRTUALSHOP)
        '//Verifica se a janela está visivel.
        If Not .Visible Then Exit Sub

        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height

        '//Define o nome dos botões conforme tradução
        Select Case tmpCurLanguage
        Case LANG_PT: CaptionBuy = "Comprar": CaptionContribuitor = "Ajude o servidor, faça uma doação!!!"
        Case LANG_EN: CaptionBuy = "Purchase": CaptionContribuitor = "Help the server, make a donation!!!"
        Case LANG_ES: CaptionBuy = "Purchase": CaptionContribuitor = "Help the server, make a donation!!!"
        End Select

        '//Obter o índice do VirtualShop selecionado By Tabs
        VirtualShopIndex = GetIndexFromVirtualShop

        '//Obter o índice do item selecionado pelo jogador
        VirtualShopSlot = GetSlotSelectedFromVirtualShop

        ' Desenha o Scroll
        If (VirtualShopMaxViewLine - 1) > (VirtualShopViewLine * VirtualShopViewLines) Then
            RenderTexture Tex_Gui(.Pic), .X + 4, .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY), 328, 310, 19, 35, 19, 35
        End If

        '//Buttons
        For i = ButtonEnum.VirtualShop_Close To ButtonEnum.VirtualShop_ScrollUp
            If CanShowButton(i) Then
                '//Renderiza o botão de compra
                If i = ButtonEnum.VirtualShop_Buy Then
                    If VirtualShopSlot > 0 Then
                        If VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum > 0 _
                           And VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum <= MAX_ITEM Then
                            '//O jogador tem o valor?
                            If PlayerHaveCashValue(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemPrice) = True Then
                                '//Cor normal
                                ColourOpacity = D3DColorARGB(255, 255, 255, 255)
                            Else
                                '//Cor opaca caso não tenha o valor do item
                                ColourOpacity = D3DColorARGB(255, 180, 60, 180)
                            End If

                            '//Renderiza o BackGround button
                            RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height, ColourOpacity
                            '//Renderiza o Texto do button buy
                            RenderText Font_Default, CaptionBuy, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, CaptionBuy) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White, , 255
                        End If
                    End If
                Else
                    '//Close Button
                    RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                End If
            End If
        Next i

        '//Escrever insentivo pra doação na janela da loja virtual
        RenderText Font_Default, CaptionContribuitor, .X + 30, (.Y + .Height - 25), White, , 255

        '//Cor normal
        ColourOpacity = D3DColorARGB(255, 255, 255, 255)
        '//Obter o índice real do slot 1 ao 8 em uma variavel
        For z = (VirtualShopScrollCount * VirtualShopViewLine) To (VirtualShopScrollCount * VirtualShopViewLine) + (VirtualShopViewLine * VirtualShopViewLines) - 1
            X = z + 1
            '//Renderiza o Icone do item/skin/mount
            If VirtualShopIndex > 0 Then
                XX = (.X + 27 + ((4 + 135) * (((((X) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) Mod 2))))
                YY = (.Y + 55 + ((4 + 43) * ((((X) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) \ 2)))

                If X > 0 And X <= (VirtualShopMaxViewLine) Then
                    If X <= UBound(VirtualShop(VirtualShopIndex).Items) Then
                        If VirtualShop(VirtualShopIndex).Items(X).ItemNum > 0 Then

                            '//Background
                            'Call VirtualShop_Background(XX, YY)

                            '//Renderiza o BackGround do item
                            Select Case VirtalShopSlot_State(X)
                            Case ButtonState.StateNormal: RenderTexture Tex_Gui(.Pic), XX, YY, 65, 319, 127, 46, 127, 46
                            Case ButtonState.StateHover: RenderTexture Tex_Gui(.Pic), XX, YY, 65, 365, 127, 46, 127, 46
                            Case ButtonState.StateClick: RenderTexture Tex_Gui(.Pic), XX, YY, 65, 411, 127, 46, 127, 46
                            End Select

                            If Item(VirtualShop(VirtualShopIndex).Items(X).ItemNum).Sprite > 0 Then
                                RenderTexture Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(X).ItemNum).Sprite), XX + 8, YY + 6, 0, 0, 32, 32, GetPicWidth(Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(X).ItemNum).Sprite)), GetPicHeight(Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(X).ItemNum).Sprite)), ColourOpacity
                            End If
                            '//Renderiza o nome do item
                            RenderText Ui_Default, Trim$(Item(VirtualShop(VirtualShopIndex).Items(X).ItemNum).Name), XX + 44, YY + 46 / 2 - 17, White, , 255

                            '//Renderiza a quantidade do item
                            RenderText Ui_Default, KeepTwoDigit(VirtualShop(VirtualShopIndex).Items(X).ItemQuant), XX + 12, YY + 27, White, , 255
                        End If
                    End If
                End If
            End If
        Next z

        '//Labels
        ' ---> Renderiza o cash do jogador na janela da loja
        RenderText Font_Default, Trim$(Player(MyIndex).Cash), .X + 355, .Y + 7, Yellow

        ' --> Renderiza a seleção de qual Tab da loja está
        ' --> Configura o posicionamento do Select e tamanho dele apartir da Tab específica.
        ' == O que for igual pra todos fica fora do for
        tmpY = .Y + 49
        Height = 3
        For i = 1 To CountTabs - 1
            Select Case i
            Case VirtualShopTabsRec.Skins: tmpX = .X + 2: Width = 46
            Case VirtualShopTabsRec.Mounts: tmpX = .X + 49: Width = 63
            Case VirtualShopTabsRec.Items: tmpX = .X + 113: Width = 50
            Case VirtualShopTabsRec.Vips: tmpX = .X + 164: Width = 45
            End Select

            If VirtualShop(i).SelectedWindow = True Then
                RenderTexture Tex_Gui(.Pic), tmpX, tmpY, 0, 320, Width, Height, 46, 2, ColourOpacity
            End If
        Next i

        '//Seleção do item
        If VirtualShopSlot > 0 Then
            If VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum > 0 Then
                If Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Sprite > 0 Then

                    '--> Renderiza o icone do item na descrição
                    RenderTexture Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Sprite), .X + 362, (.Y + 43), 0, 0, 32, 32, GetPicWidth(Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Sprite)), GetPicHeight(Tex_Item(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Sprite))
                    '//Renderiza o nome do item
                    RenderText Font_Default, Trim$(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Name), .X + 372 - (GetTextWidth(Ui_Default, Trim$(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Name)) / 2), (.Y + 85), White, , 255
                    '//Renderiza o preço do item
                    RenderText Font_Default, "Price: " & KeepTwoDigit(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemPrice), .X + 342, (.Y + 215), Yellow, , 255
                    '//Obtem a descrição do item
                    DescString = Trim$(Item(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum).Desc)

                    '//DESCRIÇÕES PERSONALIZADAS
                    '--> Item Normal
                    If VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).CustomDesc = NO Then
                        '//Wrap the text
                        WordWrap_Array Font_Default, DescString, 150, ArrayText

                        '//Loop to all items
                        '//Reset
                        tmpY = 0
                        For i = LBound(ArrayText) To UBound(ArrayText)
                            '//Set Location
                            '//Keep it centered
                            X = (.X + 277) + ((182 * 0.5) - (GetTextWidth(Ui_Default, Trim$(ArrayText(i))) * 0.5))
                            Y = (.Y + 115) + tmpY

                            '//Render the text
                            RenderText Ui_Default, Trim$(ArrayText(i)), X, Y, White

                            '//Increase the location for each line
                            tmpY = tmpY + 16
                        Next i

                        '//Item com descrição diferenciada
                    ElseIf VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).CustomDesc = YES Then
                        ' -->Utilizando um [+] antes da escrita, vai torná-lo um benefício abrangente.
                        ' -->Utilizando um [-] antes da escrita, vai torná-lo um benefício não abrangente.
                        ' -->Não utilizando nada antes da escrita, vai torná-lo uma descrição normal apenas.
                        ' -->Por favor, separe cada benefício com uma "," ao final de cada.
                        ' -->Dê preferencia primeiramente aos benefícios abrangentes, e após os não abrangentes e após a descrição, se houver uma.

                        '//Reset
                        tmpY = 0

                        '//Set Matriz
                        ArrayText = Split(DescString, ",")

                        For i = LBound(ArrayText) To UBound(ArrayText)

                            ' -->Verifica se contém um benefício abrangente
                            If InStr(1, ArrayText(i), "[+]") > 0 Then
                                '//Keep it centered
                                X = (.X + 230) + ((182 * 0.5))
                                Y = (.Y) + 110 + tmpY
                                '--> Renderiza o icone de beneficio abrangente na descrição
                                RenderTexture Tex_Gui(.Pic), X - 15, Y, 46, 404, 15, 15, 18, 18
                                RenderTexture Tex_Gui(.Pic), X - 15, Y, 46, 404 + (17 * 2), 15, 15, 18, 18
                                '//Render the text
                                RenderText Ui_Default, Trim$(Mid(ArrayText(i), 4)), X, Y, White
                            ElseIf InStr(1, ArrayText(i), "[-]") > 0 Then    ' -->Verifica se contém um benefício não abrangente
                                '//Keep it centered
                                X = (.X + 230) + ((182 * 0.5))
                                Y = (.Y) + 110 + tmpY
                                '--> Renderiza o icone de beneficio não abrangente na descrição
                                RenderTexture Tex_Gui(.Pic), X - 15, Y, 46, 404 + (17 * 1), 15, 15, 18, 18
                                '//Render the text
                                RenderText Ui_Default, Trim$(Mid(ArrayText(i), 4)), X, Y, White
                            Else  ' -->Descrição normal

                                If Len(ArrayText(i)) > 0 Then
                                    '--> Reset
                                    tmpS = tmpY

                                    '//Wrap the text
                                    WordWrap_Array Font_Default, ArrayText(i), 150, ArrayText2

                                    '//Render the text
                                    For z = LBound(ArrayText2) To UBound(ArrayText2)
                                        '//Keep it centered
                                        X = (.X + 277) + ((182 * 0.5) - (GetTextWidth(Ui_Default, Trim$(ArrayText2(z))) * 0.5))
                                        Y = (.Y + 115) + tmpS
                                        RenderText Ui_Default, Trim$(ArrayText2(z)), X, Y, White
                                        tmpS = tmpS + 16
                                    Next z
                                End If
                            End If

                            tmpY = tmpY + 16
                        Next i

                    End If
                End If
            End If

        End If

    End With
End Sub

Public Sub VirtualShopMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim tmpX As Integer, tmpY As Integer, Width As Integer, Height As Integer
    Dim f As Long, VirtualShopIndex As Long, VirtualShopSlot As Long, z As Long

    With GUI(GuiEnum.GUI_VIRTUALSHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        '//Set to top most
        UpdateGuiOrder GUI_VIRTUALSHOP

        '//Obter o índice do VirtualShop selecionado By Tabs
        VirtualShopIndex = GetIndexFromVirtualShop

        '//Obter o índice do item selecionado pelo jogador
        VirtualShopSlot = GetSlotSelectedFromVirtualShop

        '//Loop through all items
        For i = ButtonEnum.VirtualShop_Close To ButtonEnum.VirtualShop_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If

                    Select Case i
                    Case ButtonEnum.VirtualShop_ScrollUp
                        If (VirtualShopMaxViewLine - 1) > (VirtualShopViewLine * VirtualShopViewLines) Then
                            If VirtualShopScrollCount > 0 Then
                                VirtualShopScrollUp = True
                                VirtualShopScrollDown = False
                                VirtualShopScrollCount = VirtualShopScrollCount - 1
                                VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) \ (VirtualShopMaxViewLine \ VirtualShopViewLines)
                                VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)
                                VirtualShopScrollTimer = GetTickCount
                            End If
                        End If
                    Case ButtonEnum.VirtualShop_ScrollDown
                        If (VirtualShopMaxViewLine - 1) > (VirtualShopViewLine * VirtualShopViewLines) Then
                            If VirtualShopScrollCount < (VirtualShopMaxViewLine \ VirtualShopViewLines) Then
                                VirtualShopScrollUp = False
                                VirtualShopScrollDown = True
                                VirtualShopScrollCount = VirtualShopScrollCount + 1
                                VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) \ (VirtualShopMaxViewLine \ VirtualShopViewLines)
                                VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)
                                VirtualShopScrollTimer = GetTickCount
                            End If
                        End If

                    Case ButtonEnum.VirtualShop_Buy
                        If PlayerHaveCashValue(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemPrice) = True Then
                            PurchaseVirtualShop VirtualShopIndex, VirtualShopSlot
                            SetSlotSelectedFromVirtualShop 0
                        End If
                    End Select
                End If
            End If
        Next

        '//Slots
        '//Obter o índice real do slot 1 ao 8 em uma variavel
        For z = (VirtualShopScrollCount * VirtualShopViewLine) To (VirtualShopScrollCount * VirtualShopViewLine) + (VirtualShopViewLine * VirtualShopViewLines) - 1
            f = z + 1
            '//Renderiza o Icone do item/skin/mount
            If VirtualShopIndex > 0 Then
                If f > 0 And f <= (VirtualShopMaxViewLine) Then
                    If f <= UBound(VirtualShop(VirtualShopIndex).Items) Then
                        tmpX = (.X + 27 + ((4 + 135) * (((((f) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) Mod 2))))
                        tmpY = (.Y + 55 + ((4 + 43) * ((((f) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) \ 2)))
                        If CursorX >= tmpX And CursorX <= tmpX + 127 _
                           And CursorY >= tmpY And CursorY <= tmpY + 46 Then
                            '//Adiciona o State
                            SetState_VirtualShopSlot f, StateClick
                            '//Adicionar o slot selecionado a descrição da janela.
                            Call SetSlotSelectedFromVirtualShop(f)
                        End If
                    End If
                End If
            End If
        Next z

        '//Troca de Tabs Loja Virtual
        tmpY = .Y + 32
        Height = 18
        For i = 1 To CountTabs - 1
            Select Case i
            Case VirtualShopTabsRec.Skins: tmpX = .X + 2: Width = 46
            Case VirtualShopTabsRec.Mounts: tmpX = .X + 49: Width = 63
            Case VirtualShopTabsRec.Items: tmpX = .X + 113: Width = 50
            Case VirtualShopTabsRec.Vips: tmpX = .X + 164: Width = 45
            End Select

            If CursorX >= tmpX And CursorX <= tmpX + Width _
               And CursorY >= tmpY And CursorY <= tmpY + Height Then
                SwitchTabFromVirtualShop i
                Call SetSlotSelectedFromVirtualShop(0)
            End If
        Next i

        '//Check for scroll
        If CursorX >= .X + 4 And CursorX <= .X + 4 + 19 And CursorY >= .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) And CursorY <= .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) + VirtualShopScrollSize Then
            VirtualShopScrollHold = True
        End If

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Public Sub VirtualShopMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpX As Long, tmpY As Long, Width As Long, Height As Long
    Dim i As Long, VirtualShopSlot As Long, VirtualShopIndex As Long, z As Long

    With GUI(GuiEnum.GUI_VIRTUALSHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_VIRTUALSHOP Then Exit Sub

        IsHovering = False

        '//Obter o índice do VirtualShop selecionado By Tabs
        VirtualShopIndex = GetIndexFromVirtualShop

        '//Obter o índice do item selecionado pelo jogador
        VirtualShopSlot = GetSlotSelectedFromVirtualShop

        '//Loop through all items
        For i = ButtonEnum.VirtualShop_Close To ButtonEnum.VirtualShop_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    '//Renderiza o botão de compra
                    If i = ButtonEnum.VirtualShop_Buy Then
                        If VirtualShopSlot > 0 Then
                            If VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum > 0 And VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemNum <= MAX_ITEM Then
                                '//O jogador tem o valor?
                                If PlayerHaveCashValue(VirtualShop(VirtualShopIndex).Items(VirtualShopSlot).ItemPrice) = True Then
                                    If Button(i).State = ButtonState.StateNormal Then
                                        Button(i).State = ButtonState.StateHover
                                        IsHovering = True
                                        MouseIcon = 1    '//Select
                                    End If
                                End If
                            End If
                        End If
                    Else

                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover

                            IsHovering = True
                            MouseIcon = 1    '//Select
                        End If
                    End If
                End If
            End If
        Next

        '//Slots
        '//Obter o índice real do slot 1 ao 6 em uma variavel
        For z = (VirtualShopScrollCount * VirtualShopViewLine) To (VirtualShopScrollCount * VirtualShopViewLine) + (VirtualShopViewLine * VirtualShopViewLines) - 1
            X = z + 1
            '//Renderiza o Icone do item/skin/mount
            If VirtualShopIndex > 0 Then
                tmpX = (.X + 27 + ((4 + 135) * (((((X) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) Mod 2))))
                tmpY = (.Y + 55 + ((4 + 43) * ((((X) - (VirtualShopScrollCount * VirtualShopViewLine)) - 1) \ 2)))

                If X <= UBound(VirtualShop(VirtualShopIndex).Items) Then
                    If VirtualShop(VirtualShopIndex).Items(X).ItemNum > 0 Then
                        If CursorX >= tmpX And CursorX <= tmpX + 127 _
                           And CursorY >= tmpY And CursorY <= tmpY + 46 Then
                            SetState_VirtualShopSlot X, StateHover

                            IsHovering = True
                            MouseIcon = 1    '//Select
                        End If
                    End If
                End If
            End If
        Next z

        '//Troca de Tabs Loja Virtual
        tmpY = .Y + 32
        Height = 18
        For i = 1 To CountTabs - 1
            Select Case i
            Case VirtualShopTabsRec.Skins: tmpX = .X + 2: Width = 46
            Case VirtualShopTabsRec.Mounts: tmpX = .X + 49: Width = 63
            Case VirtualShopTabsRec.Items: tmpX = .X + 113: Width = 50
            Case VirtualShopTabsRec.Vips: tmpX = .X + 164: Width = 45
            End Select

            If i <> VirtualShopIndex Then
                If CursorX >= tmpX And CursorX <= tmpX + Width _
                   And CursorY >= tmpY And CursorY <= tmpY + Height Then
                    IsHovering = True
                    MouseIcon = 1    '//Select
                End If
            End If
        Next i

        '//Scroll moving
        
        Debug.Print (VirtualShopMaxViewLine \ VirtualShopViewLine)

        If (VirtualShopMaxViewLine \ VirtualShopViewLine) > 4 Then
            If CursorX >= .X + 4 And CursorX <= .X + 4 + 19 And CursorY >= .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) And CursorY <= .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) + VirtualShopScrollSize Then
                IsHovering = True
                MouseIcon = 1    '//Select
            End If
            If VirtualShopScrollHold Then
                '//Upward
                If CursorY < .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) + (VirtualShopScrollSize / 2) Then
                    If VirtualShopScrollY < VirtualShopScrollEndY - VirtualShopScrollSize Then
                        VirtualShopScrollY = (CursorY - (.Y + VirtualShopScrollStartY + (VirtualShopScrollEndY - VirtualShopScrollSize)) - (VirtualShopScrollSize / 2)) * -1
                        If VirtualShopScrollY >= VirtualShopScrollEndY - VirtualShopScrollSize Then VirtualShopScrollY = VirtualShopScrollEndY - VirtualShopScrollSize

                        VirtualShopScrollCount = VirtualShopScrollCount - 1

                        If VirtualShopScrollCount < 0 Then
                            VirtualShopScrollCount = 0
                        End If

                        VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) / (VirtualShopMaxViewLine \ VirtualShopViewLines)
                        VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)

                    End If
                End If
                '//Downward
                If CursorY > .Y + VirtualShopScrollStartY + ((VirtualShopScrollEndY - VirtualShopScrollSize) - VirtualShopScrollY) + VirtualShopScrollSize - (VirtualShopScrollSize / 2) Then
                    If VirtualShopScrollY > 0 Then
                        VirtualShopScrollY = (CursorY - (.Y + VirtualShopScrollStartY + (VirtualShopScrollEndY - VirtualShopScrollSize)) - VirtualShopScrollSize + (VirtualShopScrollSize / 2)) * -1
                        If VirtualShopScrollY <= 0 Then VirtualShopScrollY = 0

                        VirtualShopScrollCount = VirtualShopScrollCount + 1

                        'If VirtualShopScrollCount >= (VirtualShopMaxViewLine \ 4) Then
                        '    VirtualShopScrollCount = (VirtualShopMaxViewLine \ 4) - 1
                        'End If

                        VirtualShopScrollY = (VirtualShopScrollCount * VirtualShopScrollLength) / (VirtualShopMaxViewLine \ VirtualShopViewLines)
                        VirtualShopScrollY = (VirtualShopScrollLength - VirtualShopScrollY)

                    End If
                End If
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

Public Sub VirtualShopMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long

    With GUI(GuiEnum.GUI_VIRTUALSHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub

        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_VIRTUALSHOP Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.VirtualShop_Close To ButtonEnum.VirtualShop_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                        Case ButtonEnum.VirtualShop_Close
                            If GUI(GuiEnum.GUI_VIRTUALSHOP).Visible Then
                                GuiState GUI_VIRTUALSHOP, False
                            End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Virtual Shop Scroll
        VirtualShopScrollHold = False

        '//Check for dragging
        .InDrag = False
    End With
End Sub

Public Sub ClearVirtualShop()
    Dim i As Long

    For i = 1 To CountTabs - 1
        Call ZeroMemory(ByVal VarPtr(VirtualShop(i)), LenB(VirtualShop(i)))
        Erase VirtalShopSlot_State
    Next i
End Sub

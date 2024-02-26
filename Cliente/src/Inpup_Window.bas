Attribute VB_Name = "Inpup_Window"
' **************
' ** InputBox **
' **************

Public Sub DrawInputBox()
Dim i As Long

    If GettingMap Then Exit Sub

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)
        
        '//Render the window
        RenderTexture Tex_Gui(.Pic), .X, .Y, .StartX, .StartY, .Width, .Height, .Width, .Height
                
        '//Buttons
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                RenderTexture Tex_Gui(.Pic), .X + Button(i).X, .Y + Button(i).Y, Button(i).StartX(Button(i).State), Button(i).StartY(Button(i).State), Button(i).Width, Button(i).Height, Button(i).Width, Button(i).Height
                
                Select Case i
                    Case InputBox_Okay
                        RenderText Font_Default, TextUIInputConfirm, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIInputConfirm) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White
                    Case InputBox_Cancel
                        RenderText Font_Default, TextUIInputCancel, .X + Button(i).X + (Button(i).Width / 2) - (GetTextWidth(Font_Default, TextUIInputCancel) / 2) - 3, (.Y + Button(i).Y) + Button(i).Height / 2 - 11, White
                End Select
                
            End If
        Next

        '//Render Text
        RenderArrayText Font_Default, InputBoxHeader, .X + 10, .Y + 36, 250, White, , True
        
        '//Text
        RenderArrayText Font_Default, UpdateChatText(Font_Default, InputBoxText & TextLine, 210), .X + 8, .Y + 64, 252, Dark, , True
    End With
End Sub

Public Sub InputBoxKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then Exit Sub
    
    Select Case InputBoxType
        Case IB_NEWPASSWORD, IB_PASSWORDCONFIRM, IB_OLDPASSWORD
            If (isNameLegal(KeyAscii, True) And Len(InputBoxText) < (InputBoxLen - 1)) Or KeyAscii = vbKeyBack Then InputBoxText = InputText(InputBoxText, KeyAscii)
        Case IB_WITHDRAW
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInvStorage(InvCurSlot).data(InputBoxData1).value Then
                    InputBoxText = PlayerInvStorage(InvCurSlot).data(InputBoxData1).value
                End If
            End If
        Case IB_DEPOSIT
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInv(InputBoxData1).value Then
                    InputBoxText = PlayerInv(InputBoxData1).value
                End If
            End If
        Case IB_BUYITEM
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If (Shop(ShopNum).ShopItem(InputBoxData1).Price * Val(InputBoxText)) > Player(MyIndex).Money Then
                    InputBoxText = Round(Player(MyIndex).Money / Shop(ShopNum).ShopItem(InputBoxData1).Price, 0)
                End If
            End If
        Case IB_SELLITEM, IB_ADDTRADE
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInv(InputBoxData1).value Then
                    InputBoxText = PlayerInv(InputBoxData1).value
                End If
            End If
    End Select
End Sub

Public Sub InputBoxMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INPUTBOX
        
        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub InputBoxMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Textbox
        If CursorX >= .X + 22 And CursorX <= .X + 22 + 223 And CursorY >= .Y + 34 And CursorY <= .Y + 34 + 19 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
    End With
End Sub

Public Sub InputBoxMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.InputBox_Okay
                                Select Case InputBoxType
                                    Case IB_NEWPASSWORD
                                        NewPassword = InputBoxText
                                        InputBoxHeader = "Confirm your password"
                                        InputBoxType = IB_PASSWORDCONFIRM
                                        InputBoxText = vbNullString
                                    Case IB_PASSWORDCONFIRM
                                        If NewPassword = InputBoxText Then
                                            InputBoxHeader = "Enter your old password"
                                            InputBoxType = IB_OLDPASSWORD
                                            InputBoxText = vbNullString
                                        Else
                                            AddAlert "Password doesn't match", White
                                        End If
                                    Case IB_OLDPASSWORD
                                        OldPassword = InputBoxText
                                        '//Send Change Pass Data
                                        SendChangePassword NewPassword, OldPassword
                                        CloseInputBox
                                    Case IB_WITHDRAW
                                        If IsNumeric(InputBoxText) Then
                                            SendWithdrawItemTo InvCurSlot, InputBoxData1, InputBoxData2, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_DEPOSIT
                                        If IsNumeric(InputBoxText) Then
                                            SendDepositItemTo InvCurSlot, InputBoxData2, InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_BUYITEM
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Buy Item
                                            SendBuyItem InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_SELLITEM
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Sell Item
                                            SendSellItem InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_ADDTRADE
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Add Trade Item
                                            SendAddTrade 1, InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                End Select
                            Case ButtonEnum.InputBox_Cancel
                                Select Case InputBoxType
                                    Case Else
                                        CloseInputBox
                                End Select
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub OpenInputBox(ByVal cText As String, ByVal ctype As Byte, Optional ByVal Data1 As Long = 0, Optional ByVal Data2 As Long = 0)
    If GameState = GameStateEnum.InMenu Then
        If Not MenuState = MenuStateEnum.StateNormal Then Exit Sub
    End If
    
    GuiState GUI_INPUTBOX, True
    InputBoxHeader = cText
    InputBoxType = ctype
    
    Select Case ctype
        Case IB_NEWPASSWORD, IB_PASSWORDCONFIRM, IB_OLDPASSWORD
            InputBoxLen = NAME_LENGTH
            InputBoxText = vbNullString
    End Select
    InputBoxData1 = Data1
    InputBoxData2 = Data2
End Sub

Public Sub CloseInputBox()
    GuiState GUI_INPUTBOX, False
    InputBoxHeader = vbNullString
    InputBoxText = vbNullString
    InputBoxType = 0
    
    '//Password
    NewPassword = vbNullString
    OldPassword = vbNullString
End Sub

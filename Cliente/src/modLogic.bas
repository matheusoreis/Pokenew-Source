Attribute VB_Name = "modLogic"
Option Explicit

'//This function change a single digit number to two digit (often used on time)
Public Function KeepTwoDigit(ByVal Val As Long) As String
    If Val > 9 Then
        KeepTwoDigit = Val
    Else
        KeepTwoDigit = "0" & Val
    End If
End Function

'//This remove the duplicate on String Arrays
Public Sub strArrRemoveDuplicate(ByRef StringArray() As String)
Dim LowBound As Long, UpBound As Long
Dim TempArray() As String, Cur As Long
Dim a As Long, B As Long
        
    '//check for empty array
    If (Not StringArray) = True Then Exit Sub

    '//we need these often
    LowBound = LBound(StringArray)
    UpBound = UBound(StringArray)

    '//reserve check buffer
    ReDim TempArray(LowBound To UpBound)
        
    '//set first item
    Cur = LowBound
    TempArray(Cur) = StringArray(LowBound)
        
    '//loop through all items
    For a = LowBound + 1 To UpBound
        '//make a comparison against all items
        For B = LowBound To Cur
            '//if is a duplicate, exit array
            If LenB(TempArray(B)) = LenB(StringArray(a)) Then
                If InStrB(1, StringArray(a), TempArray(B), vbBinaryCompare) = 1 Then Exit For
            End If
        Next B
        '//check if the loop was exited: add new item to check buffer if not
        If B > Cur Then Cur = B: TempArray(Cur) = StringArray(a)
    Next a
        
    '//fix size
    ReDim Preserve TempArray(LowBound To Cur)
    '//copy
    StringArray = TempArray
End Sub

'//This remove all null or empty on string arrays
Public Sub strArrRemoveNull(ByRef StringArray() As String)
Dim LowBound As Long, UpBound As Long
Dim i As Long, Cur As Long
Dim TempArray() As String

    '//check for empty array
    If (Not StringArray) = True Then Exit Sub
    
    '//we need these often
    LowBound = LBound(StringArray)
    UpBound = UBound(StringArray)
    
    '//reserve check buffer
    ReDim TempArray(LowBound To UpBound)
    
    '//set first item
    Cur = LowBound
    
    '//loop through all items
    For i = LowBound To UpBound
        '//Check if string have length
        If Len(Trim$(StringArray(i))) > 0 Then
            '//Add it
            TempArray(Cur) = Trim$(StringArray(i))
            Cur = Cur + 1
        End If
    Next i
    
    '//fix size
    Cur = Cur - 1
    ReDim Preserve TempArray(LowBound To Cur)
    '//copy
    StringArray = TempArray
End Sub

'//This remove the count on the array
Public Sub byteArrRemoveData(ByRef ValArray() As Byte, ByVal dataNum As Byte)
Dim LowBound As Byte, UpBound As Byte
Dim i As Byte
Dim TempArray() As Byte
Dim sCount As Long

    '//check for empty array
    If (Not ValArray) = True Then Exit Sub

    '//we need these often
    LowBound = LBound(ValArray)
    UpBound = UBound(ValArray)
    
    '//Make sure it's above 1 or else empty it
    If UpBound <= 1 Then
        Erase ValArray
        Exit Sub
    End If
    
    '//reserve check buffer
    ReDim TempArray(LowBound To UpBound - 1)
    
    sCount = LowBound
    For i = LowBound To UpBound
        If Not i = dataNum Then
            TempArray(sCount) = ValArray(i)
            sCount = sCount + 1
        End If
    Next
        
    '//loop through all items
    'If dataNum = UpBound Then
    '    For i = LowBound To UpBound - 1
    '        TempArray(i) = ValArray(i)
    '    Next
    'Else
    '    For i = dataNum To UpBound - 1
    '        TempArray(i) = ValArray(i + 1)
    '    Next
    'End If

    '//copy
    ValArray = TempArray
End Sub

'//This remove the count on the array
Public Sub longArrRemoveData(ByRef ValArray() As Long, ByVal dataNum As Long)
Dim LowBound As Long, UpBound As Long
Dim i As Long
Dim TempArray() As Long

    '//check for empty array
    If (Not ValArray) = True Then Exit Sub

    '//we need these often
    LowBound = LBound(ValArray)
    UpBound = UBound(ValArray)
    
    '//Make sure it's above 1 or else empty it
    If UpBound <= 1 Then
        Erase ValArray
        Exit Sub
    End If
    
    '//reserve check buffer
    ReDim TempArray(LowBound To UpBound - 1)
    
    '//loop through all items
    If dataNum = UpBound Then
        For i = LowBound To UpBound - 1
            TempArray(i) = ValArray(i)
        Next
    Else
        For i = dataNum To UpBound - 1
            TempArray(i) = ValArray(i + 1)
        Next
    End If

    '//copy
    ValArray = TempArray
End Sub

'//This look for the data on the array
Public Function findDataArray(ByVal dataNum As Byte, ByRef dataArray() As Byte) As Byte
Dim i As Byte

    '//check for empty array
    If (Not dataArray) = True Then Exit Function

    findDataArray = 0

    '//Loop through all items
    For i = LBound(dataArray) To UBound(dataArray)
        If dataArray(i) = dataNum Then
            findDataArray = i
            Exit Function
        End If
    Next
End Function

'//Fade
Public Sub InitFade(ByVal WaitTimer As Long, ByVal fState As FadeStateEnum, Optional ByVal fType As Byte = 0)
    FadeState = fState
    
    If FadeState = FadeStateEnum.FadeIn Then
        FadeAlpha = 0
    ElseIf FadeState = FadeStateEnum.FadeOut Then
        FadeAlpha = 255
    End If
    
    FadeType = fType
    FadeWait = GetTickCount + WaitTimer
    
    Fade = True
End Sub

Public Sub FadeLogic()
Dim FadeComplete As Boolean

    FadeComplete = False

    '//Check if we can fade
    If Fade Then
        If GetTickCount > FadeWait Then
            '//Select which type of fading will we do
            Select Case FadeState
                Case FadeStateEnum.FadeIn  '//FadeIn
                    If FadeAlpha < 255 Then
                        FadeAlpha = FadeAlpha + 15
                        If FadeAlpha >= 255 Then
                            FadeAlpha = 255
                            FadeComplete = True
                        End If
                    End If
                Case FadeStateEnum.FadeOut '//FadeOut
                    '//Make sure that fadealpha is greater than zero to avoid error
                    If FadeAlpha > 0 Then
                        FadeAlpha = FadeAlpha - 15
                        If FadeAlpha <= 0 Then
                            FadeAlpha = 0
                            FadeComplete = True
                        End If
                    End If
            End Select
            
            '//If fade is complete, go to event process
            If FadeComplete Then
                '//Make sure that we complete the fade
                Fade = False
                FadeWait = 0
            
                Select Case FadeType
                    Case 1 ' Fade Out (Event: Change Title Screen), Fade In
                        MenuState = MenuStateEnum.StateTitleScreen
                        InitFade 0, FadeOut, 2
                    Case 2 ' Fade In
                        InitFade 2500, FadeIn, 3
                    Case 3 ' Fade Out (Event: Change Menu Screen)
                        MenuState = MenuStateEnum.StateNormal
                        CanShowCursor = True
                        InitCursorTimer = True
                        InitFade 0, FadeOut
                        
                        '//Play Menu Music
                        If Trim$(GameSetting.MenuMusic) <> "None." Then
                            If CurMusic <> Trim$(GameSetting.MenuMusic) Then
                                PlayMusic Trim$(GameSetting.MenuMusic), False, True
                            End If
                        Else
                            StopMusic True
                        End If
                    Case 4 ' Entering Game
                        InitGameState InGame
                        InitFade 0, FadeOut
                    Case 5 ' exit game
                        UnloadMain
                    Case 6 ' Reset to Menu
                        ResetMenu
                        InitFade 0, FadeOut
                    Case Else '//Do nothing
                End Select
            End If
        End If
    End If
End Sub

Public Sub AddAlert(ByVal Text As String, ByVal Color As Long)
Dim LowBound As Long, UpBound As Long
Dim ArrayText() As String
Dim MaxWidth As Long, MaxHeight As Long
Dim i As Long
    
    '//Wrap the text
    WordWrap_Array Font_Default, Text, ALERT_STRING_LENGTH, ArrayText
    
    '//we use this to get the size
    LowBound = LBound(ArrayText)
    UpBound = UBound(ArrayText)
    
    '//Check if it wrap
    If UpBound > LowBound Then
        '//Set the size
        MaxWidth = GetTextWidth(Font_Default, ArrayText(LowBound))
        For i = LowBound + 1 To UpBound
            If MaxWidth < GetTextWidth(Font_Default, ArrayText(i)) Then
                MaxWidth = GetTextWidth(Font_Default, ArrayText(i))
            End If
        Next
        MaxHeight = 16 * UpBound
    Else
        '//Set the size
        MaxWidth = GetTextWidth(Font_Default, Text)
        MaxHeight = 16
    End If
    
    '//Add the padding
    MaxHeight = MaxHeight + 26

    '//Loop through all items
    For i = MAX_ALERT To 2 Step -1
        With AlertWindow(i)
            '//Move all data
            .IsUsed = AlertWindow(i - 1).IsUsed
            .Text = AlertWindow(i - 1).Text
            .Color = AlertWindow(i - 1).Color
            .Width = AlertWindow(i - 1).Width
            .Height = AlertWindow(i - 1).Height
            .AlertTimer = AlertWindow(i - 1).AlertTimer
            .SetYPos = AlertWindow(i - 1).SetYPos - MaxHeight
            .CurYPos = AlertWindow(i - 1).CurYPos
        End With
    Next
    
    '//Add Data
    With AlertWindow(1)
        .IsUsed = True
        .Text = Text
        .Color = Color
        
        '//set size
        .Width = MaxWidth
        .Height = MaxHeight
        
        '//set position
        .SetYPos = Screen_Height - 60 - MaxHeight
        .CurYPos = Screen_Height - 60 - MaxHeight
        
        '//start timer
        .AlertTimer = GetTickCount + ALERT_TIMER
    End With
End Sub

Public Sub RemoveAlert(ByVal AlertIndex As Byte)
Dim i As Long
Dim MaxHeight As Long

    '//Check for error
    If AlertIndex <= 0 Or AlertIndex > MAX_ALERT Then Exit Sub

    MaxHeight = AlertWindow(AlertIndex).Height
    
    '//Add the padding
    MaxHeight = MaxHeight + 26

    '//Update all items
    If AlertIndex < MAX_ALERT Then
        For i = MAX_ALERT - 1 To AlertIndex Step -1
            With AlertWindow(i)
                '//Move all data
                .IsUsed = AlertWindow(i + 1).IsUsed
                .Text = AlertWindow(i + 1).Text
                .Color = AlertWindow(i + 1).Color
                .Width = AlertWindow(i + 1).Width
                .Height = AlertWindow(i + 1).Height
                .AlertTimer = AlertWindow(i + 1).AlertTimer
                .SetYPos = AlertWindow(i + 1).SetYPos + MaxHeight
            End With
        Next
    Else
        With AlertWindow(AlertIndex)
            '//Move all data
            .IsUsed = False
            .Text = vbNullString
            .Color = 0
            .Width = 0
            .Height = 0
            .AlertTimer = 0
            .SetYPos = 0
            .CurYPos = 0
        End With
    End If
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function ConvertMapX(ByVal X As Long) As Long
    ConvertMapX = X - (TileView.Left * TILE_X) - Camera.Left
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ConvertMapY = Y - (TileView.top * TILE_Y) - Camera.top
End Function

'//Player Movement
Private Function IsTryingToMove() As Boolean
    If UpKey Or DownKey Or LeftKey Or RightKey Then
        IsTryingToMove = True
    End If
End Function

Private Function CheckDirection(ByVal direction As Byte) As Boolean
Dim X As Long, Y As Long
Dim i As Long

    CheckDirection = False
 
    If PlayerPokemon(MyIndex).Num > 0 Then
        Select Case direction
            Case DIR_UP
                X = PlayerPokemon(MyIndex).X
                Y = PlayerPokemon(MyIndex).Y - 1
            Case DIR_DOWN
                X = PlayerPokemon(MyIndex).X
                Y = PlayerPokemon(MyIndex).Y + 1
            Case DIR_LEFT
                X = PlayerPokemon(MyIndex).X - 1
                Y = PlayerPokemon(MyIndex).Y
            Case DIR_RIGHT
                X = PlayerPokemon(MyIndex).X + 1
                Y = PlayerPokemon(MyIndex).Y
        End Select
    Else
        Select Case direction
            Case DIR_UP
                X = Player(MyIndex).X
                Y = Player(MyIndex).Y - 1
            Case DIR_DOWN
                X = Player(MyIndex).X
                Y = Player(MyIndex).Y + 1
            Case DIR_LEFT
                X = Player(MyIndex).X - 1
                Y = Player(MyIndex).Y
            Case DIR_RIGHT
                X = Player(MyIndex).X + 1
                Y = Player(MyIndex).Y
        End Select
    End If

    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map.Tile(X, Y).Attribute = MapAttribute.Blocked Then
        CheckDirection = True
        Exit Function
    End If
    If Map.Tile(X, Y).Attribute = MapAttribute.ConvoTile Then
        CheckDirection = True
        Exit Function
    End If
    If Map.Tile(X, Y).Attribute = MapAttribute.BothStorage Or Map.Tile(X, Y).Attribute = MapAttribute.InvStorage Or Map.Tile(X, Y).Attribute = MapAttribute.PokemonStorage Then
        CheckDirection = True
        Exit Function
    End If
    
    If Map.Tile(X, Y).Attribute = MapAttribute.FishSpot Then
        CheckDirection = True
        Exit Function
    End If
    
    '//Check Npc
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X And MapNpc(i).Y = Y Then
                CheckDirection = True
                Exit Function
            End If
        End If
        If MapNpcPokemon(i).Num > 0 Then
            If MapNpcPokemon(i).X = X And MapNpcPokemon(i).Y = Y Then
                CheckDirection = True
                Exit Function
            End If
        End If
    Next
    
    '//Check Pokemon
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = Player(MyIndex).Map Then
                If MapPokemon(i).X = X And MapPokemon(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function CanMove(Optional ByVal DirInput As Long = -1, Optional ByVal ForceMove As Boolean = False) As Boolean
Dim X As Long, Y As Long, Dir As Byte
Dim oldDir As Byte
Dim dX As Long, dY As Long
Dim setDir As Byte
Dim rndNum As Long

    CanMove = True

    '//Make sure they aren't trying to move when they are already moving
    If PlayerPokemon(MyIndex).Num > 0 Then
        If PlayerPokemon(MyIndex).Moving <> 0 Then
            CanMove = False
            Exit Function
        End If
        
        If PlayerPokemon(MyIndex).Init = YES Then
            CanMove = False
            Exit Function
        End If
        
        If PlayerPokemon(MyIndex).Status = StatusEnum.Sleep Then
            CanMove = False
            Exit Function
        End If
        If PlayerPokemon(MyIndex).Status = StatusEnum.Frozen Then
            CanMove = False
            Exit Function
        End If
    Else
        If Player(MyIndex).Moving <> 0 Then
            CanMove = False
            Exit Function
        End If
        
        If InNpcDuel > 0 Then
            CanMove = False
            Exit Function
        End If
        
        If Not ForceMove Then
            If Player(MyIndex).Action > 0 Then
                CanMove = False
                Exit Function
            End If
        End If
    End If
    
    If GettingMap Then
        CanMove = False
        Exit Function
    End If
    
    If ChatOn Then
        CanMove = False
        Exit Function
    End If
    
    If StorageType > 0 Then
        CanMove = False
        Exit Function
    End If
    If SelMenu.Visible Then
        CanMove = False
        Exit Function
    End If
    If ShopNum > 0 Then
        CanMove = False
        Exit Function
    End If
    If PlayerPokemon(MyIndex).Num <= 0 Then
        If PlayerRequest > 0 Then
            CanMove = False
            Exit Function
        End If
    End If
    If ConvoNum > 0 Then
        CanMove = False
        Exit Function
    End If
    
    '//Input data
    If PlayerPokemon(MyIndex).Num > 0 Then
        X = PlayerPokemon(MyIndex).X
        Y = PlayerPokemon(MyIndex).Y
        Dir = PlayerPokemon(MyIndex).Dir
    Else
        X = Player(MyIndex).X
        Y = Player(MyIndex).Y
        Dir = Player(MyIndex).Dir
    End If
    
    If UpKey Then setDir = DIR_UP
    If DownKey Then setDir = DIR_DOWN
    If LeftKey Then setDir = DIR_LEFT
    If RightKey Then setDir = DIR_RIGHT
    If PlayerPokemon(MyIndex).Num > 0 Then
        If PlayerPokemon(MyIndex).IsConfused = YES Then
            rndNum = Rand(0, 3)
            If rndNum <= 0 Then rndNum = 0
            If rndNum >= 3 Then rndNum = 3
            setDir = rndNum
        End If
    Else
        If Player(MyIndex).IsConfuse = YES Then
            rndNum = Rand(0, 3)
            If rndNum <= 0 Then rndNum = 0
            If rndNum >= 3 Then rndNum = 3
            setDir = rndNum
        End If
    End If
    If DirInput >= 0 Then
        setDir = DirInput
    End If
    
    If setDir = DIR_UP Then
        If PlayerPokemon(MyIndex).Num > 0 Then
            PlayerPokemon(MyIndex).Dir = DIR_UP
        Else
            Player(MyIndex).Dir = DIR_UP
        End If
        
        '//Check to see if they are trying to go out of bounds
        If Y > 0 Then
            If CheckDirection(DIR_UP) Then
                If Dir <> DIR_UP Then
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        SendPlayerPokemonDir
                    Else
                        SendPlayerDir
                    End If
                End If
                
                If PlayerPokemon(MyIndex).Num <= 0 Then
                    If Player(MyIndex).Action > 0 Then
                        Player(MyIndex).Action = 0
                    End If
                End If
                
                CanMove = False
                Exit Function
            End If
        Else
            If PlayerPokemon(MyIndex).Num <= 0 Then
                If Player(MyIndex).Action > 0 Then
                    Player(MyIndex).Action = 0
                End If
                
                If Map.LinkUp > 0 Then
                    If Editor = 0 Then
                        SendPlayerMove
                        CanMove = False
                        Exit Function
                    End If
                End If
            End If
            
            CanMove = False
            Exit Function
        End If
        
        '//Check Distance
        If PlayerPokemon(MyIndex).Num > 0 Then
            dX = PlayerPokemon(MyIndex).X - Player(MyIndex).X
            dY = (PlayerPokemon(MyIndex).Y - 1) - Player(MyIndex).Y
            
            '//Make sure we get a positive value
            If dX < 0 Then dX = dX * -1
            If dY < 0 Then dY = dY * -1
            
            If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                If Dir <> DIR_UP Then
                    SendPlayerPokemonDir
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If

    If setDir = DIR_DOWN Then
        If PlayerPokemon(MyIndex).Num > 0 Then
            PlayerPokemon(MyIndex).Dir = DIR_DOWN
        Else
            Player(MyIndex).Dir = DIR_DOWN
        End If
        
        '//Check to see if they are trying to go out of bounds
        If Y < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                If Dir <> DIR_DOWN Then
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        SendPlayerPokemonDir
                    Else
                        SendPlayerDir
                    End If
                End If
                
                If PlayerPokemon(MyIndex).Num <= 0 Then
                    If Player(MyIndex).Action > 0 Then
                        Player(MyIndex).Action = 0
                    End If
                End If
                
                CanMove = False
                Exit Function
            End If
        Else
            If PlayerPokemon(MyIndex).Num <= 0 Then
                If Player(MyIndex).Action > 0 Then
                    Player(MyIndex).Action = 0
                End If
            
                If Map.LinkDown > 0 Then
                    If Editor = 0 Then
                        SendPlayerMove
                        CanMove = False
                        Exit Function
                    End If
                End If
            End If
            CanMove = False
            Exit Function
        End If
        
        '//Check Distance
        If PlayerPokemon(MyIndex).Num > 0 Then
            dX = PlayerPokemon(MyIndex).X - Player(MyIndex).X
            dY = (PlayerPokemon(MyIndex).Y + 1) - Player(MyIndex).Y
            
            '//Make sure we get a positive value
            If dX < 0 Then dX = dX * -1
            If dY < 0 Then dY = dY * -1
            
            If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                If Dir <> DIR_DOWN Then
                    SendPlayerPokemonDir
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If
    
    If setDir = DIR_LEFT Then
        If PlayerPokemon(MyIndex).Num > 0 Then
            PlayerPokemon(MyIndex).Dir = DIR_LEFT
        Else
            Player(MyIndex).Dir = DIR_LEFT
        End If
        
        '//Check to see if they are trying to go out of bounds
        If X > 0 Then
            If CheckDirection(DIR_LEFT) Then
                If Dir <> DIR_LEFT Then
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        SendPlayerPokemonDir
                    Else
                        SendPlayerDir
                    End If
                End If
                
                If PlayerPokemon(MyIndex).Num <= 0 Then
                    If Player(MyIndex).Action > 0 Then
                        Player(MyIndex).Action = 0
                    End If
                End If
                
                CanMove = False
                Exit Function
            End If
        Else
            If PlayerPokemon(MyIndex).Num <= 0 Then
                If Player(MyIndex).Action > 0 Then
                    Player(MyIndex).Action = 0
                End If
                
                If Map.LinkLeft > 0 Then
                    If Editor = 0 Then
                        SendPlayerMove
                        CanMove = False
                        Exit Function
                    End If
                End If
            End If
            CanMove = False
            Exit Function
        End If
        
        '//Check Distance
        If PlayerPokemon(MyIndex).Num > 0 Then
            dX = (PlayerPokemon(MyIndex).X - 1) - Player(MyIndex).X
            dY = PlayerPokemon(MyIndex).Y - Player(MyIndex).Y
            
            '//Make sure we get a positive value
            If dX < 0 Then dX = dX * -1
            If dY < 0 Then dY = dY * -1
            
            If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                If Dir <> DIR_LEFT Then
                    SendPlayerPokemonDir
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If
    
    If setDir = DIR_RIGHT Then
        If PlayerPokemon(MyIndex).Num > 0 Then
            PlayerPokemon(MyIndex).Dir = DIR_RIGHT
        Else
            Player(MyIndex).Dir = DIR_RIGHT
        End If
        
        '//Check to see if they are trying to go out of bounds
        If X < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                If Dir <> DIR_RIGHT Then
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        SendPlayerPokemonDir
                    Else
                        SendPlayerDir
                    End If
                End If
                
                If PlayerPokemon(MyIndex).Num <= 0 Then
                    If Player(MyIndex).Action > 0 Then
                        Player(MyIndex).Action = 0
                    End If
                End If
                
                CanMove = False
                Exit Function
            End If
        Else
            If PlayerPokemon(MyIndex).Num <= 0 Then
                If Player(MyIndex).Action > 0 Then
                    Player(MyIndex).Action = 0
                End If
                
                If Map.LinkRight > 0 Then
                    If Editor = 0 Then
                        SendPlayerMove
                        CanMove = False
                        Exit Function
                    End If
                End If
            End If
            CanMove = False
            Exit Function
        End If
        
        '//Check Distance
        If PlayerPokemon(MyIndex).Num > 0 Then
            dX = (PlayerPokemon(MyIndex).X + 1) - Player(MyIndex).X
            dY = PlayerPokemon(MyIndex).Y - Player(MyIndex).Y
            
            '//Make sure we get a positive value
            If dX < 0 Then dX = dX * -1
            If dY < 0 Then dY = dY * -1
            
            If Not (dX <= MAX_DISTANCE And dY <= MAX_DISTANCE) Then
                If Dir <> DIR_RIGHT Then
                    SendPlayerPokemonDir
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If
End Function

Public Sub ForcePlayerMove(ByVal Dir As Byte)
    If CanMove(Dir, True) Then
        If PlayerPokemon(MyIndex).Num <= 0 Then
            Player(MyIndex).Moving = YES
    
            Select Case Player(MyIndex).Dir
                Case DIR_UP
                    SendPlayerMove
                    Player(MyIndex).yOffset = TILE_Y
                    Player(MyIndex).Y = Player(MyIndex).Y - 1
                Case DIR_DOWN
                    SendPlayerMove
                    Player(MyIndex).yOffset = TILE_Y * -1
                    Player(MyIndex).Y = Player(MyIndex).Y + 1
                Case DIR_LEFT
                    SendPlayerMove
                    Player(MyIndex).xOffset = TILE_X
                    Player(MyIndex).X = Player(MyIndex).X - 1
                Case DIR_RIGHT
                    SendPlayerMove
                    Player(MyIndex).xOffset = TILE_X * -1
                    Player(MyIndex).X = Player(MyIndex).X + 1
            End Select
            
            Select Case Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Attribute
                Case MapAttribute.Warp
                    GettingMap = True
                Case MapAttribute.Slide
                    Player(MyIndex).Action = ACTION_SLIDE
                    Player(MyIndex).ActionTmr = GetTickCount + 100
                Case MapAttribute.WarpCheckpoint
                    GettingMap = True
            End Select
        End If
    End If
End Sub

Public Sub CheckMovement()
Dim rndNum As Byte

    '//Check if movement key are being pressed
    If IsTryingToMove Then
        If CanMove Then
            If PlayerPokemon(MyIndex).Num > 0 Then
                PlayerPokemon(MyIndex).Moving = YES
    
                Select Case PlayerPokemon(MyIndex).Dir
                    Case DIR_UP
                        SendPlayerPokemonMove
                        PlayerPokemon(MyIndex).yOffset = TILE_Y
                        PlayerPokemon(MyIndex).Y = PlayerPokemon(MyIndex).Y - 1
                    Case DIR_DOWN
                        SendPlayerPokemonMove
                        PlayerPokemon(MyIndex).yOffset = TILE_Y * -1
                        PlayerPokemon(MyIndex).Y = PlayerPokemon(MyIndex).Y + 1
                    Case DIR_LEFT
                        SendPlayerPokemonMove
                        PlayerPokemon(MyIndex).xOffset = TILE_X
                        PlayerPokemon(MyIndex).X = PlayerPokemon(MyIndex).X - 1
                    Case DIR_RIGHT
                        SendPlayerPokemonMove
                        PlayerPokemon(MyIndex).xOffset = TILE_X * -1
                        PlayerPokemon(MyIndex).X = PlayerPokemon(MyIndex).X + 1
                End Select
            Else
                Player(MyIndex).Moving = YES
    
                Select Case Player(MyIndex).Dir
                    Case DIR_UP
                        SendPlayerMove
                        Player(MyIndex).yOffset = TILE_Y
                        Player(MyIndex).Y = Player(MyIndex).Y - 1
                    Case DIR_DOWN
                        SendPlayerMove
                        Player(MyIndex).yOffset = TILE_Y * -1
                        Player(MyIndex).Y = Player(MyIndex).Y + 1
                    Case DIR_LEFT
                        SendPlayerMove
                        Player(MyIndex).xOffset = TILE_X
                        Player(MyIndex).X = Player(MyIndex).X - 1
                    Case DIR_RIGHT
                        SendPlayerMove
                        Player(MyIndex).xOffset = TILE_X * -1
                        Player(MyIndex).X = Player(MyIndex).X + 1
                End Select
            
                Select Case Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Attribute
                    Case MapAttribute.Warp
                        GettingMap = True
                    Case MapAttribute.Slide
                        Player(MyIndex).Action = ACTION_SLIDE
                        Player(MyIndex).ActionTmr = GetTickCount + 100
                    Case MapAttribute.WarpCheckpoint
                        GettingMap = True
                End Select
            End If
        End If
    End If
End Sub

Public Function IsTryingToSwitchAttack() As Boolean
    IsTryingToSwitchAttack = False
    If UpMoveKey Or DownMoveKey Or LeftMoveKey Or RightMoveKey Then
        IsTryingToSwitchAttack = True
    End If
End Function

Public Sub CheckAttack()
Dim buffer As clsBuffer
Dim AttackSpeed As Long
    
    If ChatOn Then Exit Sub

    '//Check if is trying to attack
    If PlayerPokemon(MyIndex).Num > 0 Then
        If IsTryingToSwitchAttack Or AtkKey Then
            If AtkKey Then
                SetAttackMove = 0
            End If
            If UpMoveKey Then
                SetAttackMove = 1
            End If
            If DownMoveKey Then
                SetAttackMove = 2
            End If
            If LeftMoveKey Then
                SetAttackMove = 3
            End If
            If RightMoveKey Then
                SetAttackMove = 4
            End If
        
            '//ToDo: Attack Speed
            AttackSpeed = 1000
            
            If PlayerPokemon(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
                If PlayerPokemon(MyIndex).Attacking = 0 Then
                    With PlayerPokemon(MyIndex)
                        .Attacking = 1
                        .AttackTimer = GetTickCount
                                
                        .IdleTimer = GetTickCount
                        .IdleAnim = 0
                        .IdleFrameTmr = GetTickCount
                    End With
            
                    '//Check Key Press
                    Set buffer = New clsBuffer
                    buffer.WriteLong CAttack
                    buffer.WriteByte SetAttackMove
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                End If
            End If
        Else
            SetAttackMove = 0
        End If
    End If
End Sub

Public Function isInBounds() As Boolean
    isInBounds = False
    '//Check if pointed tileset is within the game area
    If curTileX >= 0 And curTileX <= Map.MaxX And curTileY >= 0 And curTileY <= Map.MaxY Then isInBounds = True
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    IsValidMapPoint = False
    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
End Function

Public Function GetKeyCodeName(ByVal KeyCode As Integer) As String
    Select Case KeyCode
        Case 32
            GetKeyCodeName = "Space"
        Case 65 To 90, 48 To 57
            GetKeyCodeName = ChrW$(KeyCode)
        Case 16
            GetKeyCodeName = "Shift"
        Case 112 To 123
            GetKeyCodeName = "F" & KeyCode - 111
        Case 17
            GetKeyCodeName = "Ctrl"
        Case 192
            GetKeyCodeName = "~"
        Case 13
            GetKeyCodeName = "Enter"
        Case 37
            GetKeyCodeName = "Left"
        Case 38
            GetKeyCodeName = "Up"
        Case 39
            GetKeyCodeName = "Right"
        Case 40
            GetKeyCodeName = "Down"
        Case 35
            GetKeyCodeName = "End"
        Case 33
            GetKeyCodeName = "Page Up"
        Case 34
            GetKeyCodeName = "Page Down"
        Case 36
            GetKeyCodeName = "Home"
        Case 45
            GetKeyCodeName = "Insert"
        Case 46
            GetKeyCodeName = "Delete"
        Case Else
            GetKeyCodeName = "Invalid"
    End Select
End Function

Public Function InvalidInput(ByVal KeyCode As Integer) As Boolean
    Select Case KeyCode
        Case 13, 16 To 17, 32 To 40, 45 To 46, 48 To 57, 65 To 90, 112 To 123, 192
            InvalidInput = False
        Case Else
            InvalidInput = True
    End Select
End Function

Public Function CheckSameKey(ByVal KeyCode As Integer) As Boolean
Dim i As Long

    For i = 1 To ControlEnum.Control_Count - 1
        If KeyCode = TmpKey(i) Then
            CheckSameKey = True
            Exit Function
        End If
    Next
End Function

'//Chatbubble
Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal Colour As Long, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1)
Dim i As Long, Index As Long

    '//set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > 255 Then chatBubbleIndex = 1
    
    '//default to new bubble
    Index = chatBubbleIndex
    
    '//loop through and see if that player/npc already has a chat bubble
    For i = 1 To 255
        If chatBubble(i).targetType = targetType Then
            If chatBubble(i).target = target Then
                '//reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                '//we use this one now, yes?
                Index = i
                Exit For
            End If
        End If
    Next
    
    '//set the bubble up
    With chatBubble(Index)
        .Msg = Msg
        .Colour = Colour
        .target = target
        .targetType = targetType
        .X = X
        .Y = Y

        .timer = GetTickCount
        .active = True
    End With
End Sub

' *************
' ** SelMenu **
' *************
Public Sub OpenSelMenu(ByVal menuType As Byte, Optional ByVal Data1 As Long = 0)
    Dim i As Long
    Dim LeftSpawn As Boolean

    '//Reset datas
    ClearSelMenu

    With SelMenu
        '//General
        .Type = menuType

        '//Select data
        Select Case menuType
        Case SelMenuType.Inv
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Use"
            AddSelMenuText "Add Held"
            'AddSelMenuText "Remove"
            If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
                AddSelMenuText "Deposit"
            ElseIf GUI(GuiEnum.GUI_SHOP).Visible Then
                AddSelMenuText "Sell"
            ElseIf GUI(GuiEnum.GUI_TRADE).Visible Then
                AddSelMenuText "Add Trade"
            End If

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.SpawnPokes
            '//Remember slot
            .Data1 = SelPoke

            '//Add text
            AddSelMenuText "Spawn"
            AddSelMenuText "Summary"
            AddSelMenuText "Remove Held"
            If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                AddSelMenuText "Deposit"
            ElseIf GUI(GuiEnum.GUI_TRADE).Visible Then
                AddSelMenuText "Add Trade"
            End If

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.PlayerPokes
            '//Remember slot
            .Data1 = SelPoke

            '//Add text
            AddSelMenuText "Call Back"
            AddSelMenuText "Summary"

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.Evolve
            '//Check Error
            If MyIndex <= 0 Then
                ClearSelMenu
                Exit Sub
            End If
            If PlayerPokemon(MyIndex).Num <= 0 Then
                ClearSelMenu
                Exit Sub
            End If
            If PlayerPokemon(MyIndex).Slot <= 0 Then
                ClearSelMenu
                Exit Sub
            End If

            '//Check Evolve
            For i = 1 To MAX_EVOLVE
                If Pokemon(PlayerPokemon(MyIndex).Num).evolveNum(i) > 0 Then
                    AddSelMenuText Trim$(Pokemon(Pokemon(PlayerPokemon(MyIndex).Num).evolveNum(i)).Name)
                End If
            Next

            '//Set visible
            .Visible = True
            LeftSpawn = False
        Case SelMenuType.Storage
            '//Add text
            AddSelMenuText "Item Storage"
            AddSelMenuText "Pokemon Storage"

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.NPCChat
            '//Add data
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Talk"

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.InvStorage
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Withdraw"

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.PokeStorage
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Summary"
            AddSelMenuText "Withdraw"
            AddSelMenuText "Release"

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.PlayerMenu
            '//Remember Index
            .Data1 = Data1

            '//Add text
            If Data1 = MyIndex Then
                AddSelMenuText "Create Party"
                AddSelMenuText "Leave Party"
                If PlayerRequest > 0 Then
                    AddSelMenuText "Cancel Request"
                End If
            Else
                AddSelMenuText Trim$(Player(Data1).Name)
                AddSelMenuText "Duel"
                AddSelMenuText "Trade"
                AddSelMenuText "Whisper"
                AddSelMenuText "Invite"
            End If

            '//Set visible
            .Visible = True
            LeftSpawn = True
        Case SelMenuType.TradeItem
            '//Remember slot
            .Data1 = Data1

            If CheckingTrade = 1 Then    ' Their
                If YourTrade.TradeSet = NO Then
                    AddSelMenuText "Remove"
                    If YourTrade.data(Data1).TradeType = 2 Then    ' Type Poke
                        AddSelMenuText "Summary"
                    End If
                End If
            ElseIf CheckingTrade = 2 Then    ' Your
                If TheirTrade.TradeSet = NO Then
                    If TheirTrade.data(Data1).TradeType = 2 Then    ' Type Poke
                        AddSelMenuText "Summary"
                    End If
                End If
            End If

            '//Set visible
            .Visible = True
            LeftSpawn = False
        Case SelMenuType.PokedexMapPokemon
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Scan"

            '//Set visible
            .Visible = True
            LeftSpawn = False
        Case SelMenuType.PokedexPlayerPokemon
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Scan"

            '//Set visible
            .Visible = True
            LeftSpawn = False
        Case SelMenuType.ConvoTileCheck
            '//Remember slot
            .Data1 = Data1

            '//Add text
            AddSelMenuText "Check"

            '//Set visible
            .Visible = True
            LeftSpawn = False
        Case SelMenuType.RevivePokes
            '//Remember slot
            .Data1 = SelPoke

            '//Add text
            AddSelMenuText "Summary"
            If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                AddSelMenuText "Deposit"
            ElseIf GUI(GuiEnum.GUI_TRADE).Visible Then
                AddSelMenuText "Add Trade"
            End If

            '//Set visible
            .Visible = True
            LeftSpawn = True
        End Select

        '//Set Position
        If .Visible Then
            If LeftSpawn Then
                .X = CursorX - .MaxWidth
            Else
                .X = CursorX
            End If
            .Y = CursorY
        End If
    End With
End Sub

Public Sub AddSelMenuText(ByVal seltext As String)
Dim sText() As String
Dim i As Long

    With SelMenu
        '//Create a temporary holder of text
        sText = .Text
        
        '//add the count
        .MaxText = .MaxText + 1
        
        '//add the temporary text to new array
        ReDim .Text(1 To .MaxText)
        For i = 1 To .MaxText - 1
            .Text(i) = sText(i)
        Next
        
        '//Add text
        .Text(.MaxText) = seltext
        
        '//Set Max Width
        If GetTextWidth(Font_Default, .Text(.MaxText)) > .MaxWidth Then
            .MaxWidth = GetTextWidth(Font_Default, .Text(.MaxText))
        End If
    End With
End Sub

Public Function SelMenuLogic(ByVal Button As Integer) As Boolean
    Dim i As Byte
    '//Select button input
    If Button = vbLeftButton Then
        Select Case SelMenu.Type
        Case SelMenuType.Inv
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Use
                    SendUseItem SelMenu.Data1
                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Add Held
                    SenAddHeld SelMenu.Data1

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function

                Case 3    '//Remove
                    If ShopNum > 0 Then
                        If GUI(GuiEnum.GUI_SHOP).Visible Then
                            '//Sell Selected Item
                            If Item(PlayerInv(SelMenu.Data1).Num).RestrictionData.CanStack = YES Then
                                OpenInputBox TextUIInputAmountHeader, IB_SELLITEM, SelMenu.Data1
                            Else
                                '//Sell Item
                                SendSellItem SelMenu.Data1
                            End If
                        End If
                    Else    ' Is ItemStorage
                        '//Check
                        If StorageType = 1 Then    '//Inv
                            If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
                                '//Deposit Selected Item
                                If Item(PlayerInv(SelMenu.Data1).Num).RestrictionData.CanStack = YES Then
                                    OpenInputBox TextUIInputAmountHeader, IB_DEPOSIT, SelMenu.Data1, 0
                                Else
                                    SendDepositItemTo InvCurSlot, 0, SelMenu.Data1
                                End If
                            End If
                        ElseIf TradeIndex > 0 Then    '//Trade
                            If GUI(GuiEnum.GUI_TRADE).Visible Then
                                '//Add Trade Selected Item
                                If Item(PlayerInv(SelMenu.Data1).Num).RestrictionData.CanStack = YES Then
                                    OpenInputBox TextUIInputAmountHeader, IB_ADDTRADE, SelMenu.Data1
                                Else
                                    '//Add Trade Item
                                    SendAddTrade 1, SelMenu.Data1
                                End If
                            End If
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 4    '//Deposit

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.SpawnPokes
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Spawn
                    If PlayerPokemons(SelPoke).CurHP > 0 Then
                        SendPlayerPokemonState 1, SelPoke
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Summary
                    SummaryType = 1
                    SummarySlot = SelPoke
                    SummaryData = 0
                    If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                        GuiState GUI_POKEMONSUMMARY, True
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 3    '//Remove Held
                    SendRemoveHeld SelPoke
                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 4    '//Deposit
                    '//Check
                    If StorageType = 2 Then    '//Pokemon
                        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                            '//Deposit Selected Pokemon
                            SendDepositPokemon PokemonCurSlot, SelPoke
                        End If
                    ElseIf TradeIndex > 0 Then    '//Add Trade
                        If GUI(GuiEnum.GUI_TRADE).Visible Then
                            '//Add Trade Pokemon
                            SendAddTrade 2, SelPoke
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.PlayerPokes
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Call Back
                    SendPlayerPokemonState 0, SelPoke

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Summary
                    SummaryType = 1
                    SummarySlot = SelPoke
                    SummaryData = 0
                    If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                        GuiState GUI_POKEMONSUMMARY, True
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.Evolve
            EvolveSelect = SelMenu.CurPick
            If EvolveSelect > 0 Then
                '//MsgBox
                OpenChoiceBox TextUIChoiceEvolve, CB_EVOLVE

                '//Clear
                ClearSelMenu
                SelMenuLogic = True
                Exit Function
            End If
        Case SelMenuType.Storage
            Select Case SelMenu.CurPick
            Case 1    '//Item Storage
                SendOpenStorage 1

                '//Clear
                ClearSelMenu
                SelMenuLogic = True
                Exit Function
            Case 2    '//Pokemon Storage
                SendOpenStorage 2

                '//Clear
                ClearSelMenu
                SelMenuLogic = True
                Exit Function
            End Select
        Case SelMenuType.NPCChat
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Talk
                    SendConvo 1, SelMenu.Data1

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.InvStorage
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Withdraw
                    '//Check
                    If StorageType = 1 Then    '//Inv
                        If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
                            '//Deposit Selected Item
                            If Item(PlayerInvStorage(InvCurSlot).data(SelMenu.Data1).Num).RestrictionData.CanStack = YES Then
                                OpenInputBox TextUIInputAmountHeader, IB_WITHDRAW, SelMenu.Data1, 0
                            Else
                                SendWithdrawItemTo InvCurSlot, SelMenu.Data1, 0
                            End If
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.PokeStorage
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Summary
                    SummaryType = 2
                    SummarySlot = SelMenu.Data1
                    SummaryData = PokemonCurSlot
                    If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                        GuiState GUI_POKEMONSUMMARY, True
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Withdraw
                    '//Check
                    If StorageType = 2 Then    '//Pokemon
                        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                            '//Deposit Selected Item

                            Dim hasSelected As Boolean
                            For i = 1 To MAX_STORAGE
                                If IsPokemonSelected(i) Then
                                    SendWithdrawPokemon PokemonCurSlot, i
                                    ClearPokemonSelected i
                                    hasSelected = True
                                End If
                            Next i

                            If Not hasSelected Then
                                SendWithdrawPokemon PokemonCurSlot, SelMenu.Data1
                            End If
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 3    '//Release
                    If StorageType = 2 Then    '//Pokemon
                        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                            If PlayerPokemonStorage(PokemonCurSlot).data(SelMenu.Data1).Num > 0 Then
                                ReleaseStorageData = SelMenu.Data1
                                ReleaseStorageSlot = PokemonCurSlot
                                OpenChoiceBox TextUIChoiceRelease, CB_RELEASE
                            End If
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.PlayerMenu
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Player Info / Create Party
                    If SelMenu.Data1 = MyIndex Then
                        '//Create Party
                        If InParty <= 0 Then
                            SendCreateParty
                        End If
                    Else
                        '//Player Info
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Duel / Leave Party
                    If SelMenu.Data1 = MyIndex Then
                        '//Leave Party
                        If InParty > 0 Then
                            SendLeaveParty
                        End If
                    Else
                        PlayerRequest = SelMenu.Data1
                        RequestType = 1
                        SendRequest SelMenu.Data1, 1    '//1 = Duel
                        AddAlert "Request sent", White
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 3    '//Trade /  Cancel Request
                    If SelMenu.Data1 = MyIndex Then
                        SendRequestState 0
                    Else
                        PlayerRequest = SelMenu.Data1
                        RequestType = 2
                        SendRequest SelMenu.Data1, 2    '//2 = Trade
                        AddAlert "Request sent", White
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 4    '//Whisper
                    ChatTab = Trim$(Player(SelMenu.Data1).Name)
                    Language
                    ChatOn = True
                    ChatMinimize = False
                    EditTab = False
                    MyChat = vbNullString

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 5    '//Invite
                    If InParty > 0 Then
                        PlayerRequest = SelMenu.Data1
                        RequestType = 3
                        SendRequest SelMenu.Data1, 3    '//2 = Trade
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.TradeItem
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Remove
                    If CheckingTrade = 1 Then    ' their trade
                        SendRemoveTrade SelMenu.Data1
                    ElseIf CheckingTrade = 2 Then    ' your trade
                        SummaryType = 4
                        SummarySlot = SelMenu.Data1
                        SummaryData = 0
                        If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                            GuiState GUI_POKEMONSUMMARY, True
                        End If
                    End If
                Case 2
                    If CheckingTrade = 1 Then    ' their trade
                        SummaryType = 3
                        SummarySlot = SelMenu.Data1
                        SummaryData = 0
                        If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                            GuiState GUI_POKEMONSUMMARY, True
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.PokedexMapPokemon
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Scan
                    SendScanPokedex 1, SelMenu.Data1

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.PokedexPlayerPokemon
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Scan
                    SendScanPokedex 2, SelMenu.Data1

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.ConvoTileCheck
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Scan
                    SendConvo 2, SelMenu.Data1

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case SelMenuType.RevivePokes
            If SelMenu.Data1 > 0 Then
                Select Case SelMenu.CurPick
                Case 1    '//Summary
                    SummaryType = 1
                    SummarySlot = SelPoke
                    SummaryData = 0
                    If (GUI(GUI_POKEMONSUMMARY).Visible = False) Then
                        GuiState GUI_POKEMONSUMMARY, True
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                Case 2    '//Deposit
                    '//Check
                    If StorageType = 2 Then    '//Pokemon
                        If GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                            '//Deposit Selected Pokemon
                            SendDepositPokemon PokemonCurSlot, SelPoke
                        End If
                    ElseIf TradeIndex > 0 Then    '//Add Trade
                        If GUI(GuiEnum.GUI_TRADE).Visible Then
                            '//Add Trade Pokemon
                            SendAddTrade 2, SelPoke
                        End If
                    End If

                    '//Clear
                    ClearSelMenu
                    SelMenuLogic = True
                    Exit Function
                End Select
            End If
        Case Else
            '//Clear
            ClearSelMenu
            SelMenuLogic = False
            Exit Function
        End Select
    End If

    '//Clear
    ClearSelMenu
    SelMenuLogic = False
End Function

'//Action Msg
Public Sub CreateActionMsg(ByVal Msg As String, ByVal Color As Long, ByVal X As Long, ByVal Y As Long)
Dim i As Long

    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= 255 Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Msg = Msg
        .Color = Color
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
        .Alpha = 255
    End With

    ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
    ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    
    '//find the new high index
    For i = 255 To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    '//make sure we don't overflow
    If Action_HighIndex > 255 Then Action_HighIndex = 255
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ActionMsg(Index).Msg = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    
    '//find the new high index
    For i = 255 To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    '//make sure we don't overflow
    If Action_HighIndex > 255 Then Action_HighIndex = 255
End Sub

'//Animation
Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim Layer As Long
Dim FrameCount As Long

    '//if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATION Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            '//if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            '//check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).timer(Layer) + looptime <= GetTickCount Then
                '//check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    '//if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
End Sub

'//player Pokemon
Public Function CanPlayerPokemonEvolve() As Boolean
Dim evolveNum As Long

    If MyIndex <= 0 Or MyIndex > MAX_PLAYER Then Exit Function
    If PlayerPokemon(MyIndex).Num <= 0 Then Exit Function
    If PlayerPokemon(MyIndex).Slot <= 0 Then Exit Function
    
    With PlayerPokemons(PlayerPokemon(MyIndex).Slot)
        For evolveNum = 1 To MAX_EVOLVE
            If Pokemon(.Num).evolveNum(evolveNum) > 0 Then
                '//Check Condition
                If Pokemon(.Num).EvolveLevel(evolveNum) <= .Level Then
                    '//ToDo: Condition
                    CanPlayerPokemonEvolve = True
                    Exit Function
                End If
            End If
        Next
    End With
    
    CanPlayerPokemonEvolve = False
End Function

'//Weather
Public Sub InitWeather(ByVal WeatherType As WeatherEnum)
Dim i As Long

    With Weather
        .Type = WeatherType
        
        Select Case WeatherType
            Case WeatherEnum.Rain
                .MaxDrop = 100
                ReDim Weather.Drop(1 To .MaxDrop)
                For i = 1 To .MaxDrop
                    .Drop(i).X = Rand(0, (Screen_Width * 2))
                    .Drop(i).Y = Rand((-1 * Screen_Height), -32)
                    .Drop(i).SpeedY = 6
                    .Drop(i).Pic = 1
                    .Drop(i).PicType = Rand(0, 3)
                    If .Drop(i).PicType < 0 Then .Drop(i).PicType = 0
                    If .Drop(i).PicType > 3 Then .Drop(i).PicType = 3
                Next
                Weather.InitDrop = True
            Case WeatherEnum.Snow
                .MaxDrop = 255
                ReDim Weather.Drop(1 To .MaxDrop)
                For i = 1 To .MaxDrop
                    .Drop(i).X = Rand(0, Screen_Width)
                    .Drop(i).Y = Rand((-1 * Screen_Height), -32)
                    .Drop(i).SpeedY = Rand(1, 4)
                    .Drop(i).Pic = 2
                    .Drop(i).PicType = Rand(0, 3)
                    If .Drop(i).PicType < 0 Then .Drop(i).PicType = 0
                    If .Drop(i).PicType > 3 Then .Drop(i).PicType = 3
                Next
                Weather.InitDrop = True
            Case WeatherEnum.SandStorm
                .MaxDrop = 50
                ReDim Weather.Drop(1 To .MaxDrop)
                For i = 1 To .MaxDrop
                    .Drop(i).X = Rand((-1 * Screen_Width), -32)
                    .Drop(i).Y = Rand(0, Screen_Height)
                    .Drop(i).SpeedY = Rand(6, 9)
                    .Drop(i).Pic = 3
                    .Drop(i).PicType = Rand(0, 3)
                    If .Drop(i).PicType < 0 Then .Drop(i).PicType = 0
                    If .Drop(i).PicType > 3 Then .Drop(i).PicType = 3
                Next
                Weather.InitDrop = True
            Case WeatherEnum.Hail
                .MaxDrop = 150
                ReDim Weather.Drop(1 To .MaxDrop)
                For i = 1 To .MaxDrop
                    .Drop(i).X = Rand(0, (Screen_Width * 2))
                    .Drop(i).Y = Rand((-1 * Screen_Height), -32)
                    .Drop(i).SpeedY = 6
                    .Drop(i).Pic = 4
                    .Drop(i).PicType = 0
                    If .Drop(i).PicType < 0 Then .Drop(i).PicType = 0
                    If .Drop(i).PicType > 3 Then .Drop(i).PicType = 3
                Next
                Weather.InitDrop = True
            Case WeatherEnum.Sunny
                .MaxDrop = 0
                ReDim Weather.Drop(0)
                Weather.InitDrop = True
        End Select
    End With
End Sub

Public Function IsInvItem(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsInvItem = 0

    For i = 1 To MAX_PLAYER_INV
        If PlayerInv(i).Num > 0 Then
            DrawX = GUI(GuiEnum.GUI_INVENTORY).X + (7 + ((5 + TILE_X) * (((i - 1) Mod 5))))
            DrawY = GUI(GuiEnum.GUI_INVENTORY).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 5)))
            
            If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
                IsInvItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsInvSlot(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsInvSlot = 0

    For i = 1 To MAX_PLAYER_INV
        DrawX = GUI(GuiEnum.GUI_INVENTORY).X + (7 + ((5 + TILE_X) * (((i - 1) Mod 5))))
        DrawY = GUI(GuiEnum.GUI_INVENTORY).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 5)))
            
        If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
            IsInvSlot = i
            Exit Function
        End If
    Next
End Function

Public Function IsInvStorageItem(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsInvStorageItem = 0

    For i = 1 To MAX_STORAGE
        If PlayerInvStorage(InvCurSlot).data(i).Num > 0 Then
            DrawX = GUI(GuiEnum.GUI_INVSTORAGE).X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
            DrawY = GUI(GuiEnum.GUI_INVSTORAGE).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))
                
            If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
                IsInvStorageItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsStorage_Poke() As Byte
    Dim i As Integer, Count As Byte
    
    IsStorage_Poke = 0
    For i = ButtonEnum.PokemonStorage_Slot1 To ButtonEnum.PokemonStorage_Slot5
        Count = Count + 1
        If Button(i).State = ButtonState.StateHover Then
            IsStorage_Poke = Count
            Exit Function
        End If
    Next i
End Function

Public Function IsStorage_Item() As Byte
    Dim i As Integer, Count As Byte
    
    IsStorage_Item = 0
    For i = ButtonEnum.InvStorage_Slot1 To ButtonEnum.InvStorage_Slot5
        Count = Count + 1
        If Button(i).State = ButtonState.StateHover Then
            IsStorage_Item = Count
            Exit Function
        End If
    Next i
End Function

Public Function IsInvStorageSlot(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsInvStorageSlot = 0

    For i = 1 To MAX_STORAGE
        DrawX = GUI(GuiEnum.GUI_INVSTORAGE).X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
        DrawY = GUI(GuiEnum.GUI_INVSTORAGE).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))
            
        If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
            IsInvStorageSlot = i
            Exit Function
        End If
    Next
End Function

Public Function IsPokeStorage(ByVal X As Long, ByVal Y As Long) As Long
    Dim DrawX As Long, DrawY As Long
    Dim i As Long

    IsPokeStorage = 0

    For i = 1 To MAX_STORAGE
        If PlayerPokemonStorage(PokemonCurSlot).data(i).Num > 0 Then
            DrawX = GUI(GuiEnum.GUI_POKEMONSTORAGE).X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
            DrawY = GUI(GuiEnum.GUI_POKEMONSTORAGE).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))

            If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
                IsPokeStorage = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsPokeStorageSlot(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsPokeStorageSlot = 0

    For i = 1 To MAX_STORAGE
        DrawX = GUI(GuiEnum.GUI_POKEMONSTORAGE).X + (98 + ((5 + TILE_X) * (((i - 1) Mod 7))))
        DrawY = GUI(GuiEnum.GUI_POKEMONSTORAGE).Y + (37 + ((5 + TILE_Y) * ((i - 1) \ 7)))
            
        If X >= DrawX And X <= DrawX + TILE_X And Y >= DrawY And Y <= DrawY + TILE_Y Then
            IsPokeStorageSlot = i
            Exit Function
        End If
    Next
End Function

Public Function IsShopItem(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsShopItem = 0
    
    For i = ShopAddY To ShopAddY + 8
        If i > 0 And i <= MAX_SHOP_ITEM Then
            DrawX = GUI(GuiEnum.GUI_SHOP).X + (31 + ((4 + 127) * (((((i + 1) - ShopAddY) - 1) Mod 3))))
            DrawY = GUI(GuiEnum.GUI_SHOP).Y + (42 + ((4 + 78) * ((((i + 1) - ShopAddY) - 1) \ 3)))
                
            If X >= DrawX And X <= DrawX + 127 And Y >= DrawY And Y <= DrawY + 78 Then
                IsShopItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsTradeYourItem(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsTradeYourItem = 0
    
    For i = 1 To MAX_TRADE
        DrawX = GUI(GuiEnum.GUI_TRADE).X + (12 + ((3 + 44) * ((i - 1) Mod 4)))
        DrawY = GUI(GuiEnum.GUI_TRADE).Y + (71 + ((3 + 46) * ((i - 1) \ 4)))
                
        If X >= DrawX And X <= DrawX + 44 And Y >= DrawY And Y <= DrawY + 46 Then
            If YourTrade.data(i).TradeType > 0 Then
                IsTradeYourItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsTradeTheirItem(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long

    IsTradeTheirItem = 0
    
    For i = 1 To MAX_TRADE
        DrawX = GUI(GuiEnum.GUI_TRADE).X + (222 + ((3 + 44) * ((i - 1) Mod 4)))
        DrawY = GUI(GuiEnum.GUI_TRADE).Y + (71 + ((3 + 46) * ((i - 1) \ 4)))

        If X >= DrawX And X <= DrawX + 44 And Y >= DrawY And Y <= DrawY + 46 Then
            If TheirTrade.data(i).TradeType > 0 Then
                IsTradeTheirItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsPokedexSlot(ByVal X As Long, ByVal Y As Long) As Long
Dim DrawX As Long, DrawY As Long
Dim i As Long
Dim pokeDexIndex As Long

    IsPokedexSlot = -1

    For i = (PokedexViewCount * 8) To (PokedexViewCount * 8) + 31
        If i >= 0 And i <= PokedexHighIndex Then
            pokeDexIndex = i + 1
            DrawX = GUI(GuiEnum.GUI_POKEDEX).X + (31 + ((4 + 44) * (((((i + 1) - (PokedexViewCount * 8)) - 1) Mod 8))))
            DrawY = GUI(GuiEnum.GUI_POKEDEX).Y + (42 + ((4 + 46) * ((((i + 1) - (PokedexViewCount * 8)) - 1) \ 8)))
            
            If X >= DrawX And X <= DrawX + 44 And Y >= DrawY And Y <= DrawY + 46 Then
                IsPokedexSlot = i
                Exit Function
            End If
        End If
    Next
End Function

'Public Function IsRankingSlot(ByVal X As Long, ByVal Y As Long) As Long
'Dim DrawX As Long, DrawY As Long
'Dim i As Long
'Dim RankingIndex As Long

'    IsRankingSlot = -1
'RenderTexture Tex_Gui(.Pic), .X + 30, .Y + 41 + (31 * (i - 1)), 28, 328, 212, 28, 212, 28
'    For i = (RankingCount * 8) To (RankingViewCount * 8) + 31
'        If i >= 0 And i <= RankingHighIndex Then
'            RankingIndex = i + 1
'
'            DrawX = GUI(GuiEnum.GUI_RANK).X + 30
'            DrawY = GUI(GuiEnum.GUI_RANK).Y + 41 + (31 * (i - 1))
'
'            If X >= DrawX And X <= DrawX + 212 And Y >= DrawY And Y <= DrawY + 28 Then
'                IsRankingSlot = i
'                Exit Function
'            End If
'        End If
'    Next
'End Function

'//Stat
Public Function GetStatBuff(ByVal Stat As Long, ByVal StatBuff As Long) As Long
    On Error GoTo errorHandler

    '//Select Buff Stage
    Select Case StatBuff
        Case -6: GetStatBuff = Stat * 0.25
        Case -5: GetStatBuff = Stat * 0.285
        Case -4: GetStatBuff = Stat * 0.33
        Case -3: GetStatBuff = Stat * 0.4
        Case -2: GetStatBuff = Stat * 0.5
        Case -1: GetStatBuff = Stat * 0.66
        Case 0: GetStatBuff = Stat * 1
        Case 1: GetStatBuff = Stat * 1.5
        Case 2: GetStatBuff = Stat * 2
        Case 3: GetStatBuff = Stat * 2.5
        Case 4: GetStatBuff = Stat * 3
        Case 5: GetStatBuff = Stat * 3.5
        Case 6: GetStatBuff = Stat * 4
        Case Else:
    End Select
    
    Exit Function
errorHandler:
    GetStatBuff = Stat
End Function

Public Function CalculateSpeed(ByVal Spd As Long) As Long
Dim RangePercent As Long
    
    On Error GoTo errorHandler
    
    RangePercent = ((Spd / 100) / (255 / 100)) * 100
    CalculateSpeed = Round(((12 - 4) * (RangePercent / 100)) + 4, 0)
    
    Exit Function
errorHandler:
    CalculateSpeed = 4
End Function

Public Function CheckNatureString(ByVal natureNum As Integer) As String
    Select Case natureNum
        Case PokemonNature.None: CheckNatureString = "None"
        '//Neutral
        Case PokemonNature.NatureHardy: CheckNatureString = "Hardy"
        Case PokemonNature.NatureDocile: CheckNatureString = "Docile"
        Case PokemonNature.NatureSerious: CheckNatureString = "Serious"
        Case PokemonNature.NatureBashful: CheckNatureString = "Bashful"
        Case PokemonNature.NatureQuirky: CheckNatureString = "Quirky"
        '//Others
        Case PokemonNature.NatureLonely: CheckNatureString = "Lonely"
        Case PokemonNature.NatureBrave: CheckNatureString = "Brave"
        Case PokemonNature.NatureAdamant: CheckNatureString = "Adamant"
        Case PokemonNature.NatureNaughty: CheckNatureString = "Naughty"
        Case PokemonNature.NatureBold: CheckNatureString = "Bold"
        Case PokemonNature.NatureRelaxed: CheckNatureString = "Relaxed"
        Case PokemonNature.NatureImpish: CheckNatureString = "Impish"
        Case PokemonNature.NatureLax: CheckNatureString = "Lax"
        Case PokemonNature.NatureTimid: CheckNatureString = "Timid"
        Case PokemonNature.NatureHasty: CheckNatureString = "Hasty"
        Case PokemonNature.NatureJolly: CheckNatureString = "Jolly"
        Case PokemonNature.NatureNaive: CheckNatureString = "Naive"
        Case PokemonNature.NatureModest: CheckNatureString = "Modest"
        Case PokemonNature.NatureMild: CheckNatureString = "Mild"
        Case PokemonNature.NatureQuiet: CheckNatureString = "Quiet"
        Case PokemonNature.NatureRash: CheckNatureString = "Rash"
        Case PokemonNature.NatureCalm: CheckNatureString = "Calm"
        Case PokemonNature.NatureGentle: CheckNatureString = "Gentle"
        Case PokemonNature.NatureSassy: CheckNatureString = "Sassy"
        Case PokemonNature.NatureCareful: CheckNatureString = "Careful"
        Case Else: CheckNatureString = "Error"
    End Select
End Function

Public Function CheckPokeBallString(ByVal ballNum As Byte) As String
    Select Case ballNum
        Case BallEnum.b_Pokeball: CheckPokeBallString = "Poke Ball"
        Case BallEnum.b_Greatball: CheckPokeBallString = "Great Ball"
        Case BallEnum.b_Ultraball: CheckPokeBallString = "Ultra Ball"
        Case BallEnum.b_Masterball: CheckPokeBallString = "Master Ball"
        Case BallEnum.b_Primerball: CheckPokeBallString = "Primer Ball"
        Case BallEnum.b_CherishBall: CheckPokeBallString = "Cherish Ball"
        Case BallEnum.b_LuxuryBall: CheckPokeBallString = "Luxury Ball"
        Case BallEnum.b_FriendBall: CheckPokeBallString = "Friend Ball"
        Case BallEnum.b_NetBall: CheckPokeBallString = "Net Ball"
        Case BallEnum.b_DiveBall: CheckPokeBallString = "Dive Ball"
        Case BallEnum.b_RepeatBall: CheckPokeBallString = "Repeat Ball"
        Case BallEnum.b_TimerBall: CheckPokeBallString = "Timer Ball"
        Case BallEnum.b_SafariBall: CheckPokeBallString = "Safari Ball"
        Case BallEnum.b_QuickBall: CheckPokeBallString = "Quick Ball"
        Case BallEnum.b_DuskBall: CheckPokeBallString = "Dusk Ball"
        Case BallEnum.b_LoveBall: CheckPokeBallString = "Love Ball"
    End Select
End Function

Public Function GetLevelNextExp(ByVal Level As Long) As Long
    'GetLevelNextExp = (1.2 * ((Level + 1) ^ 3)) - (15 * ((Level + 1) ^ 2)) + (100 * (Level + 1)) - 140
    GetLevelNextExp = ((Level + 5) ^ 3) * (((((Level + 5) + 1) / 3) + 24) / 50)
End Function

Public Function GetPlayerHP(ByVal Level As Long) As Long
    GetPlayerHP = ((250 * Level) / 100) + ((10 + Level) / 2)
End Function

Public Function GetProcessorID() As String
Dim myWMI As Object, myObj As Object, Itm

    Set myWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set myObj = myWMI.ExecQuery("SELECT * FROM " & "Win32_Processor")
    
    For Each Itm In myObj
        GetProcessorID = Itm.ProcessorID
        Exit Function
    Next
End Function

Public Function RandomNumBetween(ByVal LowerLimit As Long, ByVal UpperLimit As Long) As Long
  RandomNumBetween = Rnd * (UpperLimit - LowerLimit) + LowerLimit
End Function

Function SecondsToHMS(ByRef Segundos As Long) As String
    Dim HR As Long, MS As Long, SS As Long, MM As Long, DD As Long, MES As Long, YY As Long
    Dim Total As Long, Count As Long

    If Segundos = 0 Then Exit Function
    YY = (Segundos \ 31104000)
    MES = (Segundos \ 2592000)
    DD = (Segundos \ 86400)
    HR = (Segundos \ 3600)
    MM = (Segundos \ 60)
    SS = Segundos
    'ms = (Segundos * 10)

    ' Pega o total de segundos pra trabalharmos melhor na variavel!
    Total = Segundos

    If HR > 0 Then
        '// Horas
        Do While (Total >= 3600)
            Total = Total - 3600
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = Count & "H "
            Count = 0
        End If
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "M "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "S "
            Count = 0
        End If
    ElseIf MM > 0 Then
        '// Minutos
        Do While (Total >= 60)
            Total = Total - 60
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "M "
            Count = 0
        End If
        '// Segundos
        Do While (Total > 0)
            Total = Total - 1
            Count = Count + 1
        Loop
        If Count > 0 Then
            SecondsToHMS = SecondsToHMS & Count & "S "
            Count = 0
        End If
    ElseIf SS > 0 Then
        ' Joga na função esse segundo.
        SecondsToHMS = SS & "S "
        Total = Total - SS
    End If
End Function

Public Sub ResetServerInfo()
    Dim i As Byte

    For i = 1 To MAX_SERVER_LIST
        ServerInfo(i).Status = "Offline"
        ServerInfo(i).Player = 0
        ServerInfo(i).Colour = BrightRed
    Next i
End Sub

Function GetColStr(Colour As Integer)
    If Colour < 10 Then
        GetColStr = "0" & Colour
    Else
        GetColStr = Colour
    End If
End Function

Function GetMapNameColour() As Byte
    Select Case Map.Moral
        Case MAP_MORAL_DANGER: GetMapNameColour = White
        Case MAP_MORAL_SAFE: GetMapNameColour = White
        Case MAP_MORAL_ARENA: GetMapNameColour = White
        Case MAP_MORAL_SAFARI: GetMapNameColour = White
        Case MAP_MORAL_PVP: GetMapNameColour = Yellow
        Case Else: GetMapNameColour = White
    End Select
End Function

Function GetWeekDay() As String
    Select Case GameWeek
    Case 1
        '//Language
        'Public Const LANG_PT As Byte = 0
        'Public Const LANG_EN As Byte = 1
        'Public Const LANG_ES As Byte = 2
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Dom"
        Case LANG_EN: GetWeekDay = "Sun"
        Case LANG_ES: GetWeekDay = "Dom"
        End Select
    Case 2
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Seg"
        Case LANG_EN: GetWeekDay = "Sun"
        Case LANG_ES: GetWeekDay = "Seg"
        End Select
    Case 3
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Ter"
        Case LANG_EN: GetWeekDay = "Tue"
        Case LANG_ES: GetWeekDay = "Ter"
        End Select
    Case 4
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Qua"
        Case LANG_EN: GetWeekDay = "Wed"
        Case LANG_ES: GetWeekDay = "Qua"
        End Select
    Case 5
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Qui"
        Case LANG_EN: GetWeekDay = "Thu"
        Case LANG_ES: GetWeekDay = "Qui"
        End Select
    Case 6
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Sex"
        Case LANG_EN: GetWeekDay = "Fri"
        Case LANG_ES: GetWeekDay = "Sex"
        End Select
    Case 7
        Select Case GameSetting.CurLanguage
        Case LANG_PT: GetWeekDay = "Sab"
        Case LANG_EN: GetWeekDay = "Sat"
        Case LANG_ES: GetWeekDay = "Sab"
        End Select
    End Select
End Function

Public Function CheckPassivaMount() As Boolean
    If Player(MyIndex).TempSpritePassiva > 0 Then CheckPassivaMount = True
End Function



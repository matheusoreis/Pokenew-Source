Attribute VB_Name = "modText"
Option Explicit

'//Custom Font Stuffs
Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Public Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
End Type

'//Available Font
Public Font_Default As CustomFont
Public Ui_Default As CustomFont

Private Const Font_Path As String = "\fonts\"

' *************
' ** Chatbox **
' *************
'//Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

'//Text buffer
Public Type ChatTextBuffer
    Text As String
    Color As Long
End Type

'//Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

'//GUI consts
Public Const MaxChatLine As Long = 9
Public Const MaxChatWidth As Long = 300

Public Const chatScrollX As Long = 2
Public Const chatScrollTop As Long = 24
Public Const chatScrollW As Long = 24
Public Const chatScrollH As Long = 35
Public Const chatScrollL As Long = 56
Public Const ChatOffsetX As Long = 28
Public Const ChatOffsetY As Long = 35
Public Const ChatWidth As Long = 327

Public ChatHold As Boolean
Public chatScrollY As Long
Public ChatScroll As Long
Public ChatScrollUp As Boolean
Public ChatScrollDown As Boolean
Public ChatScrollTimer As Long
Public totalChatLines As Long

'//Initialise Custom Font Texture and Settings
Public Sub InitFont()
Dim FilePath As String
Dim FontName As String
Dim SizeX As Long, SizeY As Long

    'On Error GoTo errHandler
    
    '//Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '//Font Default
    '//Check name of font
    FilePath = App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Font_Path
    FontName = "texdefault"
    SizeX = 256
    SizeY = 256
    Set Font_Default.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath & FontName & GFX_EXT, SizeX, SizeY, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Font_Default.TextureSize.X = SizeX
    Font_Default.TextureSize.y = SizeY
    '//Load Font Settings
    LoadFontHeader Font_Default, FilePath & FontName & DATA_EXT
    
    '//Ui Font
    '//Check name of font
    FilePath = App.Path & Texture_Path & Trim$(GameSetting.ThemePath) & Font_Path
    FontName = "rockwell_15"
    SizeX = 256
    SizeY = 256
    Set Ui_Default.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath & FontName & GFX_EXT, SizeX, SizeY, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Ui_Default.TextureSize.X = SizeX
    Ui_Default.TextureSize.y = SizeY
    '//Load Font Settings
    LoadFontHeader Ui_Default, FilePath & FontName & DATA_EXT
    
    Exit Sub
errHandler:
    MsgBox "Error loading font. Exiting...", vbCritical
    UnloadMain
End Sub

'//Load Font Setting on file
Private Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    On Error GoTo errHandler

    '//Load the header information
    FileNum = FreeFile
    Open FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    '//Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    '//Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        '//tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        '//Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0) '//Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).y = 0
            .Vertex(0).z = 0
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).z = 0
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
    
    Exit Sub
errHandler:
    MsgBox "Error loading font. Exiting...", vbCritical
    UnloadMain
End Sub

'//Default Color Setting by number
Public Function dx8Colour(ByVal colourNum As Long, Optional ByVal Alpha As Byte = 255) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 ' Dark
            dx8Colour = D3DColorARGB(Alpha, 75, 75, 75)
    End Select
End Function

'//This render the given text on screen
Public Sub RenderText(ByRef theFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal y As Long, ByVal Color As Long, Optional ByVal isColorSet As Boolean = True, Optional ByVal Alpha As Long = 255, Optional ByVal Effect As Boolean)
    Dim TempVA(0 To 3) As TLVERTEX, TempVAS(0 To 3) As TLVERTEX
    Dim TempColor As Long, ResetColor As Byte
    Dim v2 As D3DVECTOR2, v3 As D3DVECTOR2
    Dim u As Single, v As Single
    Dim i As Long, j As Long
    Dim TempStr() As String
    Dim Count As Integer
    Dim Ascii() As Byte
    Dim Row As Integer
    Dim KeyPhrase As Byte
    Dim srcRect As RECT
    Dim yOffset As Single
    Dim ignoreChar As Long

    '//set the color
    If isColorSet Then Color = dx8Colour(Color, Alpha)

    '//Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub

    '//Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)

    '//Set the temp color (or else the first character has no color)
    TempColor = Color

    'Fazer a Sombra
    If Effect = True Then RenderText theFont, Text, X + 1, y + 1, Black

    '//Set the texture
    D3DDevice.SetTexture 0, theFont.Texture
    CurrentTexture = -1

    '//Set Default Position
    X = X + 2
    y = y + 2

    '//Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * theFont.CharHeight
            Count = 0
            '//Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)

            '//Loop through the characters
            For j = 1 To Len(TempStr(i))
                ' check for colour change
                If Mid$(TempStr(i), j, 1) = ColourChar Then
                    Color = Val(Mid$(TempStr(i), j + 1, 2))
                    ' make sure the colour exists
                    If Color = -1 Then
                        TempColor = ResetColor
                    Else
                        TempColor = dx8Colour(Color, Alpha)
                    End If
                    ignoreChar = 3
                End If

                ' check if we're ignoring this character
                If ignoreChar > 0 Then
                    ignoreChar = ignoreChar - 1
                Else
                    '//Copy from the cached vertex array to the temp vertex array
                    CopyMemory TempVA(0), theFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4

                    '//Set up the verticies
                    TempVA(0).X = X + Count
                    TempVA(0).y = y + yOffset
                    TempVA(1).X = TempVA(1).X + X + Count
                    TempVA(1).y = TempVA(0).y
                    TempVA(2).X = TempVA(0).X
                    TempVA(2).y = TempVA(2).y + TempVA(0).y
                    TempVA(3).X = TempVA(1).X
                    TempVA(3).y = TempVA(2).y

                    '//Set the colors
                    TempVA(0).Color = TempColor
                    TempVA(1).Color = TempColor
                    TempVA(2).Color = TempColor
                    TempVA(3).Color = TempColor

                    '//Draw the verticies
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size

                    '//Shift over the the position to render the next character
                    Count = Count + theFont.HeaderInfo.CharWidth(Ascii(j - 1))

                    '//Check to reset the color
                    If ResetColor Then
                        ResetColor = 0
                        TempColor = Color
                    End If
                End If
            Next j
        End If
    Next i
End Sub

'//this check the actual size of the whole string
Public Function GetTextWidth(ByRef theFont As CustomFont, ByVal Text As String) As Integer
Dim LoopI As Integer

    '//Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    '//Loop through the text
    For LoopI = 1 To Len(Text)
        GetTextWidth = GetTextWidth + theFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
    Next LoopI
End Function

'//This wrap the text and place it on a array for each line
Public Sub WordWrap_Array(ByRef theFont As CustomFont, ByVal Text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, Size As Long, lastSpace As Long, B As Long

    On Error Resume Next
    

    '//Too small of text
    If Len(Text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = Text
        Exit Sub
    End If
    
    '//default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For i = 1 To Len(Text)
        '//if it's a space, store it
        Select Case Mid$(Text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        '//Add up the size
        Size = Size + theFont.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
        
        '//Check for too large of a size
        If Size > MaxLineLen Then
            '//Check if the last space was too far back
            If i - lastSpace > 12 Then
                '//Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, B, (i - 1) - B))
                B = i - 1
                Size = 0
            Else
                '//Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, B, lastSpace - B))
                B = lastSpace + 1
                
                '//Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = GetTextWidth(theFont, Mid$(Text, lastSpace, i - lastSpace))
            End If
        End If
        
        '//Remainder
        If i = Len(Text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(Text, B, i)
            End If
        End If
    Next
End Sub

'//Update the text to fit on the valid width
Public Function UpdateChatText(ByRef theFont As CustomFont, ByVal Text As String, ByVal MaxWidth As Long) As String
Dim i As Long, X As Long
    
    If GetTextWidth(theFont, Text) > MaxWidth Then
        For i = Len(Text) To 1 Step -1
            X = X + theFont.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
            If X > MaxWidth Then
                UpdateChatText = Right$(Text, Len(Text) - i + 1)
                Exit For
            End If
        Next
    Else
        UpdateChatText = Text
    End If
End Function

Public Function PreviewChatText(ByRef theFont As CustomFont, ByVal Text As String, ByVal MaxWidth As Long) As String
Dim i As Long, X As Long
    
    If GetTextWidth(theFont, Text) > MaxWidth Then
        For i = 1 To Len(Text)
            X = X + theFont.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
            If X > MaxWidth Then
                PreviewChatText = Left$(Text, i)
                Exit For
            End If
        Next
    Else
        PreviewChatText = Text
    End If
End Function

Public Function CensorWord(ByVal SString As String) As String
    CensorWord = String(Len(SString), "*")
End Function

'//This function input the keyascii to the string variable
Public Function InputText(ByVal tempString As String, ByVal KeyAscii As Integer) As String
    '//if KeyAscii is backspace then delete the last letter
    If (KeyAscii = vbKeyBack) Then
        '//Check if string have letters
        If LenB(tempString) > 0 Then tempString = Mid$(tempString, 1, Len(tempString) - 1)
    End If
    '//Check if keyascii is valid key
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
        tempString = tempString & ChrW$(KeyAscii)
    End If
    '//enter
    InputText = tempString
End Function

Public Function isNameLegal(ByVal KeyAscii As Integer, Optional ByVal DisableSpaceBar As Boolean = False) As Boolean
    If DisableSpaceBar Then
        If KeyAscii = 32 Then
            isNameLegal = False
            Exit Function
        End If
    End If
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii = 32) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        isNameLegal = True
    End If
End Function

Public Function isStringLegal(ByVal KeyAscii As Integer, Optional ByVal DisableSpaceBar As Boolean = False) As Boolean
    If DisableSpaceBar Then
        If KeyAscii = 32 Then
            isStringLegal = False
            Exit Function
        End If
    End If
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii = 32) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 33 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 91 And KeyAscii <= 96) Or (KeyAscii >= 123 And KeyAscii <= 126) Then
        isStringLegal = True
    End If
End Function

Public Function CheckNameInput(ByVal Name As String, Optional ByVal HaveStringLimit As Boolean = False, Optional ByVal MaxLimit As Long = 0, Optional ByVal isString As Boolean = False) As Boolean
Dim i As Long, n As Long

    If Not HaveStringLimit Then
        ' Check if name is within the letter limit
        If Len(Name) <= 2 Or Len(Name) >= MaxLimit Then
            CheckNameInput = False
            Exit Function
        End If
    End If
    
    ' Check Legal Asc key
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))
        
        If isString Then
            If Not isStringLegal(n, True) Then
                CheckNameInput = False
                Exit Function
            End If
        Else
            If Not isNameLegal(n, True) Then
                CheckNameInput = False
                Exit Function
            End If
        End If
    Next
    
    CheckNameInput = True
End Function

Public Sub RenderArrayText(ByRef theFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal y As Long, ByVal MaxLineLen As Long, ByVal Colour As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal KeepCentered As Boolean = False)
Dim theArray() As String, i As Long, MaxLine As Long
Dim DrawX As Long

    ' convert the single text to array text
    WordWrap_Array theFont, Text, MaxLineLen, theArray
    ' check how many lines does the array have
    MaxLine = UBound(theArray)
    If MaxLine > 1 Then
        ' if line have more than one then render each line
        For i = 1 To MaxLine
            If KeepCentered Then
                DrawX = X + ((MaxLineLen / 2) - (GetTextWidth(theFont, theArray(i)) / 2))
            Else
                DrawX = X
            End If
            RenderText theFont, theArray(i), DrawX, y, Colour, Alpha
            ' increase the y position by height of the font
            y = y + 16
        Next
    Else
        If KeepCentered Then
            DrawX = X + ((MaxLineLen / 2) - (GetTextWidth(theFont, Text) / 2))
        Else
            DrawX = X
        End If
        ' if line have one then render the text
        RenderText theFont, Text, DrawX, y, Colour, Alpha
    End If
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
    Dim textX As Long, textY As Long
    Dim Color As Long, Name As String

    With Player(Index)
        '//Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .y < TileView.top Or .y > TileView.bottom Then Exit Sub

        Name = Trim$(.Name) + " " + Trim$(.Level)
        Select Case .Access
        Case ACCESS_MODERATOR
            Color = BrightGreen
        Case ACCESS_MAPPER
            Color = BrightBlue
        Case ACCESS_DEVELOPER
            Color = BrightCyan
        Case ACCESS_CREATOR
            Color = BrightRed
        Case Else
            Color = Yellow
        End Select

        Select Case .TempSprite
        Case TEMP_SPRITE_GROUP_BIKE
            '//calc pos
            textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
            If .Sprite < 1 Or .Sprite > Count_Character Then
                textY = ConvertMapY(.y * (TILE_Y)) + .yOffset
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Character(.Sprite)) / 4)
            End If
        Case TEMP_SPRITE_GROUP_DIVE
            '//calc pos
            textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
            If .Sprite < 1 Or .Sprite > Count_Character Then
                textY = ConvertMapY(.y * TILE_Y) + .yOffset
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Character(.Sprite)) / 4)
            End If
        Case TEMP_SPRITE_GROUP_MOUNT
            '//calc pos
            textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
            If .Sprite < 1 Or .Sprite > Count_Character Then
                textY = ConvertMapY(.y * (TILE_Y)) + .yOffset
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Character(.Sprite)) / 2.3)
            End If
        Case Else
            '//calc pos
            textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
            If .Sprite < 1 Or .Sprite > Count_Character Then
                textY = ConvertMapY(.y * TILE_Y) + .yOffset
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Character(.Sprite)) / 4)
            End If
        End Select

        '//Draw name
        If .StealthMode = YES Then
            If Index = MyIndex Then
                RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
                RenderText Font_Default, Name, textX, textY, Color
                '//Status
                If .Status > 0 Then
                    RenderTexture Tex_Misc(Misc_Status), (textX - 2) + (((GetTextWidth(Font_Default, Name) + 8) / 2) - 10), textY + 18 + 2, 0, (.Status - 1) * 8, 20, 8, 20, 8
                End If
            End If
        Else
            RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
            RenderText Font_Default, Name, textX, textY, Color
            '//Status
            If .Status > 0 Then
                RenderTexture Tex_Misc(Misc_Status), (textX - 2) + (((GetTextWidth(Font_Default, Name) + 8) / 2) - 10), textY + 18 + 2, 0, (.Status - 1) * 8, 20, 8, 20, 8
            End If
        End If
    End With
End Sub

Public Sub DrawNpcName(ByVal MapNpcNum As Long)
Dim textX As Long, textY As Long
Dim Color As Long, Name As String
    
    With MapNpc(MapNpcNum)
        If .Num <= 0 Then Exit Sub
        
        '//Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .y < TileView.top Or .y > TileView.bottom Then Exit Sub
    
        Color = White
        Name = Trim$(Npc(.Num).Name)
        
        '//calc pos
        textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
        If Npc(.Num).Sprite < 1 Or Npc(.Num).Sprite > Count_Character Then
            textY = ConvertMapY(.y * TILE_Y) + .yOffset
        Else
            textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Character(Npc(.Num).Sprite)) / 4)
        End If
        
        '//Draw name
        RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
        RenderText Font_Default, Name, textX, textY, Color
    End With
End Sub

Public Sub DrawPokemonName(ByVal MapPokeNum As Long)
Dim textX As Long, textY As Long
Dim Color As Long, Name As String

    
    With MapPokemon(MapPokeNum)
        If .Num <= 0 Then Exit Sub
        
        '//Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .y < TileView.top Or .y > TileView.bottom Then Exit Sub
        
        If PlayerPokedex(.Num).Scanned = NO Then Exit Sub
            
        Color = White
        Name = Trim$(Pokemon(.Num).Name)
        
        '//calc pos
        textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
        If Pokemon(.Num).Sprite < 1 Or Pokemon(.Num).Sprite > Count_Pokemon Then
            textY = ConvertMapY(.y * TILE_Y) + .yOffset
        Else
            If Pokemon(.Num).ScaleSprite = YES Then
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - ((GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) + ConvertInverse(Pokemon(.Num).NameOffSetY)))
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) / 4 + ConvertInverse(Pokemon(.Num).NameOffSetY))
            End If
        End If
        
        '//Draw name
        RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
        RenderText Font_Default, Name, textX, textY, Color
        '//Status
        If .Status > 0 Then
            RenderTexture Tex_Misc(Misc_Status), (textX - 2) + (((GetTextWidth(Font_Default, Name) + 8) / 2) - 10), textY + 18 + 2, 0, (.Status - 1) * 8, 20, 8, 20, 8
        End If
    End With
End Sub

Private Function ConvertInverse(ByVal value As Integer) As Integer
    If value > 0 Then
        ConvertInverse = -value
    ElseIf value < 0 Then
        ConvertInverse = Math.Abs(value)
    End If
End Function

Public Sub DrawPlayerPokemonName(ByVal Index As Long)
Dim textX As Long, textY As Long
Dim Color As Long, Name As String
    
    With PlayerPokemon(Index)
        If .Num <= 0 Then Exit Sub
        
        '//Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .y < TileView.top Or .y > TileView.bottom Then Exit Sub
    
        Color = White
        Name = Trim$(Pokemon(.Num).Name)
        
        '//calc pos
        textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
        If Pokemon(.Num).Sprite < 1 Or Pokemon(.Num).Sprite > Count_Pokemon Then
            textY = ConvertMapY(.y * TILE_Y) + .yOffset
        Else
            If Pokemon(.Num).ScaleSprite = YES Then
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - ((GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) + ConvertInverse(Pokemon(.Num).NameOffSetY)))
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) / 4 + ConvertInverse(Pokemon(.Num).NameOffSetY))
            End If
        End If
        
        '//Draw name
        If Player(Index).StealthMode = NO Then
            RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
            RenderText Font_Default, Name, textX, textY, Color
            '//Status
            If .Status > 0 Then
                RenderTexture Tex_Misc(Misc_Status), (textX - 2) + (((GetTextWidth(Font_Default, Name) + 8) / 2) - 10), textY + 18 + 2, 0, (.Status - 1) * 8, 20, 8, 20, 8
            End If
        End If
    End With
End Sub

Public Sub DrawNpcPokemonName(ByVal MapPokeNum As Long)
Dim textX As Long, textY As Long
Dim Color As Long, Name As String
    
    With MapNpcPokemon(MapPokeNum)
        If .Num <= 0 Then Exit Sub
        
        '//Check if Player is within screen area
        If .X < TileView.Left Or .X > TileView.Right Then Exit Sub
        If .y < TileView.top Or .y > TileView.bottom Then Exit Sub
    
        Color = White
        Name = Trim$(Pokemon(.Num).Name)
        
        '//calc pos
        textX = ConvertMapX(.X * TILE_X) + .xOffset + ((TILE_X \ 2) - ((GetTextWidth(Font_Default, Name) \ 2) + 2))
        If Pokemon(.Num).Sprite < 1 Or Pokemon(.Num).Sprite > Count_Pokemon Then
            textY = ConvertMapY(.y * TILE_Y) + .yOffset
        Else
            If Pokemon(.Num).ScaleSprite = YES Then
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - ((GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) + ConvertInverse(Pokemon(.Num).NameOffSetY)))
            Else
                textY = ConvertMapY(.y * TILE_Y) + .yOffset - (GetPicHeight(Tex_Pokemon(Pokemon(.Num).Sprite)) / 4 + ConvertInverse(Pokemon(.Num).NameOffSetY))
            End If
        End If
        
        '//Draw name
        RenderTexture Tex_System(gSystemEnum.UserInterface), textX - 2, textY, 0, 8, GetTextWidth(Font_Default, Name) + 8, 18, 1, 1, D3DColorARGB(100, 0, 0, 0)
        RenderText Font_Default, Name, textX, textY, Color
        '//Status
        If .Status > 0 Then
            RenderTexture Tex_Misc(Misc_Status), (textX - 2) + (((GetTextWidth(Font_Default, Name) + 8) / 2) - 10), textY + 18 + 2, 0, (.Status - 1) * 8, 20, 8, 20, 8
        End If
    End With
End Sub

' *************
' ** Chatbox **
' *************
Public Sub RenderChatTextBuffer()
Dim theFont As CustomFont

    '//Changing font here
    theFont = Font_Default
    
    '//Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    D3DDevice.SetTexture 0, theFont.Texture
    CurrentTexture = -1
    
    If ChatArrayUbound > 0 Then
        D3DDevice.SetStreamSource 0, ChatVBS, FVF_Size
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        D3DDevice.SetStreamSource 0, ChatVB, FVF_Size
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim y As Single
Dim Y2 As Single
Dim i As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long
Dim theFont As CustomFont

    On Error Resume Next

    '//Changing font here
    theFont = Font_Default
    
    '//set the offset of each line
    yOffset = 14

    '//Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    '//Get the number of characters in all the visible buffer
    Size = 0

    For LoopC = (Chunk * ChatBufferChunk) - (MaxChatLine - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).Text)
    Next
    
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) '//Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    '//Set the base position
    X = GUI(GuiEnum.GUI_CHATBOX).X + ChatOffsetX
    y = GUI(GuiEnum.GUI_CHATBOX).y + ChatOffsetY

    '//Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (MaxChatLine - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        '//Set the temp color
        TempColor = ChatTextBuffer(LoopC).Color
        
        '//Set the Y position to be used
        Y2 = y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        '//Loop through each line if there are line breaks (vbCrLf)
        Count = 0 'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then  'Dont bother with empty strings
            '//Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
                '//Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                
                '//tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
                u = ((Ascii - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
                v = Row * theFont.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * pos))
                    .Color = TempColor
                    .X = (X) + Count
                    .y = (Y2)
                    .tu = u
                    .tv = v
                    .rhw = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * pos))
                    .Color = TempColor
                    .X = (X) + Count
                    .y = (Y2) + theFont.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + theFont.RowFactor
                    .rhw = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * pos))
                    .Color = TempColor
                    .X = (X) + Count + theFont.HeaderInfo.CellWidth
                    .y = (Y2) + theFont.HeaderInfo.CellHeight
                    .tu = u + theFont.ColFactor
                    .tv = v + theFont.RowFactor
                    .rhw = 1
                End With
                
                '//Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * pos)) = ChatVA(0 + (6 * pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * pos))
                    .Color = TempColor
                    .X = (X) + Count + theFont.HeaderInfo.CellWidth
                    .y = (Y2)
                    .tu = u + theFont.ColFactor
                    .tv = v
                    .rhw = 1
                End With
                
                ChatVA(5 + (6 * pos)) = ChatVA(2 + (6 * pos))
                
                '//Update the character we are on
                pos = pos + 1
                
                '//Shift over the the position to render the next character
                Count = Count + theFont.HeaderInfo.CharWidth(Ascii)
                
                '//Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).Color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice Is Nothing Then '//Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = D3DDevice.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_Size * pos * 6, 0, ChatVAS(0)
        Set ChatVB = D3DDevice.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
End Sub

Public Sub AddText(ByVal Text As String, ByVal tColor As Long)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim B As Long
Dim Color As Long
Dim theFont As CustomFont
Dim MaxY As Long

    '//Changing font here
    theFont = Font_Default

    Color = dx8Colour(tColor)

    '//Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
        '//Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        '//Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
            '//If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            '//Add up the size
            Size = Size + theFont.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            '//Check for too large of a size
            If Size > ChatWidth Then
                '//Check if the last space was too far back
                If i - lastSpace > 10 Then
                    '//Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)), Color
                    B = i - 1
                    Size = 0
                Else
                    '//Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)), Color
                    B = lastSpace + 1
                    '//Count all the words we ignored (the ones that weren't printed, but are before "i")
                    Size = GetTextWidth(theFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            '//This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If B <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), B, i), Color
            End If
        Next i
    Next TSLoop
    
    '//Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If theFont.RowPitch = 0 Then Exit Sub
    
    '//update chat scroll
    If ChatScroll > MaxChatLine Then
        ChatScroll = ChatScroll + 1
        '//scrolling up
        If ChatScroll >= totalChatLines Then ChatScroll = totalChatLines
            
        '//Update scrollbar
        MaxY = totalChatLines
        If MaxY < MaxChatLine Then MaxY = MaxChatLine
        chatScrollY = (chatScrollL / (MaxY - 8)) * (ChatScroll - 8)
    Else
        '//reset
        ChatScroll = MaxChatLine: chatScrollY = 0
    End If
    
    '//Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal Text As String, ByVal Color As Long)
Dim LoopC As Long

    '//Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    '//Set the values
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).Color = Color
    
    '//set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub

Public Sub ScrollChatBox(ByVal direction As Byte)
Dim MaxY As Long

    '//do a quick exit if we don't have enough text to scroll
    If totalChatLines < MaxChatLine Then
        ChatScroll = MaxChatLine
        UpdateChatArray
        Exit Sub
    End If
    '//actually scroll
    If direction = 0 Then '//up
        ChatScroll = ChatScroll + 1
    Else '//down
        ChatScroll = ChatScroll - 1
    End If
    '//scrolling down
    If ChatScroll < MaxChatLine Then ChatScroll = MaxChatLine
    '//scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    
    '//Scroll bar
    MaxY = totalChatLines
    If MaxY < MaxChatLine Then MaxY = MaxChatLine
    If ChatScroll = MaxChatLine Then
        chatScrollY = 0
    ElseIf ChatScroll = totalChatLines Then
        chatScrollY = chatScrollL
    Else
        chatScrollY = (chatScrollL / (MaxY - 8)) * (ChatScroll - 8)
    End If
    
    '//update the array
    UpdateChatArray
End Sub

Public Sub ClearChat()
Dim i As Long

    For i = 1 To ChatTextBufferSize
        ChatTextBuffer(i).Text = vbNullString
        ChatTextBuffer(i).Color = 0
    Next
    
    ChatScroll = MaxChatLine
End Sub

'//Map Attribute
Public Function DrawMapAttributes()
Dim X As Long, y As Long
Dim tx As Long, ty As Long

    If frmEditor_Map.optType(2).value = True Then
        '//Render Dark Alpha color on screen to easily notice the attribute tags
        RenderTexture Tex_System(gSystemEnum.UserInterface), 0, 0, 0, 8, Screen_Width, Screen_Height, 1, 1, D3DColorARGB(160, 0, 0, 0)
        
        For X = TileView.Left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, y) Then
                    With Map.Tile(X, y)
                        tx = ((ConvertMapX(X * TILE_X)) - 4) + (TILE_X * 0.5)
                        ty = ((ConvertMapY(y * TILE_Y)) - 7) + (TILE_Y * 0.5)
                        Select Case .Attribute
                            Case MapAttribute.Blocked
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 0, 0)
                                RenderText Font_Default, "B", tx, ty, BrightRed
                            Case MapAttribute.NpcSpawn
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 255, 255)
                                RenderText Font_Default, "N", tx, ty, White
                            Case MapAttribute.NpcAvoid
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 255, 255)
                                RenderText Font_Default, "A", tx, ty, Grey
                            Case MapAttribute.Warp
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 16, 224, 237)
                                RenderText Font_Default, "W", tx, ty, Cyan
                            Case MapAttribute.HealPokemon
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 181, 230, 29)
                                RenderText Font_Default, "H", tx, ty, BrightGreen
                            Case MapAttribute.BothStorage
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 174, 201)
                                RenderText Font_Default, "B", tx, ty, Pink
                            Case MapAttribute.InvStorage
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 174, 201)
                                RenderText Font_Default, "I", tx, ty, Pink
                            Case MapAttribute.PokemonStorage
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 174, 201)
                                RenderText Font_Default, "P", tx, ty, Pink
                            Case MapAttribute.ConvoTile
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 255, 255)
                                RenderText Font_Default, "C", tx, ty, White
                            Case MapAttribute.Slide
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 255, 255, 255)
                                RenderText Font_Default, "S", tx, ty, White
                            Case MapAttribute.Checkpoint
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 16, 224, 237)
                                RenderText Font_Default, "C", tx, ty, Cyan
                            Case MapAttribute.WarpCheckpoint
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 16, 224, 237)
                                RenderText Font_Default, "WC", tx, ty, Cyan
                            Case MapAttribute.FishSpot
                                RenderTexture Tex_System(gSystemEnum.UserInterface), ConvertMapX(X * TILE_X), ConvertMapY(y * TILE_Y), 0, 8, TILE_X, TILE_Y, 1, 1, D3DColorARGB(100, 16, 224, 237)
                                RenderText Font_Default, "F", tx, ty, Green
                        End Select
                    End With
                End If
            Next
        Next
    End If
End Function

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String
Dim X As Long, y As Long
Dim x2 As Long, Y2 As Long
Dim MaxWidth As Long
Dim i As Long
    
    With chatBubble(Index)
        '//Set Default
        X = ConvertMapX(.X * TILE_X) + 16
        y = ConvertMapY(.y * TILE_Y) - 28
        
        '//Got target
        If .targetType = TARGET_TYPE_PLAYER And .target > 0 Then
            '//it's a player
            If IsPlaying(.target) Then
                If Player(.target).Map = Player(MyIndex).Map Then
                    If Player(.target).StealthMode = YES Then
                        Exit Sub
                    End If
                    
                    '//it's on our map - get co-ords
                    X = ConvertMapX((Player(.target).X * TILE_X) + Player(.target).xOffset) + 16
                    y = ConvertMapY((Player(.target).y * TILE_Y) + Player(.target).yOffset) - 28
                End If
            End If
        End If
        
        '//word wrap the text
        WordWrap_Array Font_Default, .Msg, ChatBubbleWidth, theArray
                
        '//find max width
        For i = 1 To UBound(theArray)
            If GetTextWidth(Font_Default, theArray(i)) > MaxWidth Then MaxWidth = GetTextWidth(Font_Default, theArray(i))
        Next
        
        '//Increase size for padding
        MaxWidth = MaxWidth + 5
                
        '//calculate the new position
        x2 = X - (MaxWidth \ 2)
        Y2 = y - (UBound(theArray) * 12)
                
        ' **************
        ' ** Top-Left **
        ' **************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 - 8, Y2 - 8, 0, 0, 8, 8, 8, 8
        ' ***************
        ' ** Top-Right **
        ' ***************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 + MaxWidth, Y2 - 8, 16, 0, 8, 8, 8, 8
        ' *********
        ' ** Top **
        ' *********
        RenderTexture Tex_Misc(Misc_Chatbubble), x2, Y2 - 8, 8, 0, MaxWidth, 8, 8, 8
        ' *****************
        ' ** Bottom-Left **
        ' *****************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 - 8, y, 0, 16, 8, 8, 8, 8
        ' ******************
        ' ** Bottom-Right **
        ' ******************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 + MaxWidth, y, 16, 16, 8, 8, 8, 8
        ' ************
        ' ** Bottom **
        ' ************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2, y, 8, 16, MaxWidth, 8, 8, 8
        ' **********
        ' ** Left **
        ' **********
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 - 8, Y2, 0, 8, 8, (UBound(theArray) * 12), 8, 8
        ' ***********
        ' ** Right **
        ' ***********
        RenderTexture Tex_Misc(Misc_Chatbubble), x2 + MaxWidth, Y2, 16, 8, 8, (UBound(theArray) * 12), 8, 8
        ' ************
        ' ** Center **
        ' ************
        RenderTexture Tex_Misc(Misc_Chatbubble), x2, Y2, 8, 8, MaxWidth, (UBound(theArray) * 12), 8, 8
        ' ***********
        ' ** Point **
        ' ***********
        RenderTexture Tex_Misc(Misc_Chatbubble), X - 5, y, 24, 0, 8, 11, 8, 11
        
        '//render each line centralised
        For i = 1 To UBound(theArray)
            RenderText Font_Default, theArray(i), X - (GetTextWidth(Font_Default, theArray(i)) / 2) - 2, Y2 - 3, .Colour
            Y2 = Y2 + 12
        Next
        '//check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub


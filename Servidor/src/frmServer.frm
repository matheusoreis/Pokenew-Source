VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PokeNew "
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "frmServer.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtLog"
      Tab(0).Control(1)=   "txtCommand"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Administrativo"
      TabPicture(2)   =   "frmServer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkStaffOnly"
      Tab(2).Control(1)=   "frmInfo"
      Tab(2).Control(2)=   "lblCPS"
      Tab(2).Control(3)=   "lblGameTime"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Editores"
      TabPicture(3)   =   "frmServer.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "cmdEditStore"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "editorItem"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.CommandButton editorItem 
         Caption         =   "Edit Item"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdEditStore 
         Caption         =   "Edit Store"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkStaffOnly 
         Caption         =   "Modo Desenvolvedor"
         Height          =   255
         Left            =   -70440
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame frmInfo 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   4
         Top             =   720
         Width           =   6615
         Begin VB.CommandButton cmdShutdown 
            Caption         =   "Desligar Servidor"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Recarregar Pokemons Fish"
            Height          =   375
            Left            =   720
            TabIndex        =   17
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdExp 
            Caption         =   "Ativar"
            Height          =   375
            Left            =   4080
            TabIndex        =   13
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtExpHour 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4200
            TabIndex        =   11
            Text            =   "0"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.HScrollBar scrlExp 
            Height          =   255
            Left            =   3720
            Max             =   5
            Min             =   1
            TabIndex        =   8
            Top             =   960
            Value           =   1
            Width           =   2295
         End
         Begin VB.CommandButton btnReload 
            Caption         =   "Recarregar o Mapa"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   7
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton btnReload 
            Caption         =   "Recarregar os Npc's"
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   6
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CommandButton btnReload 
            Caption         =   "Recarregar os Pokémons"
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   5
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label2 
            Caption         =   "Evento EXP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Horas:"
            Height          =   195
            Left            =   3720
            TabIndex        =   12
            Top             =   1440
            Width           =   465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   3480
            X2              =   3480
            Y1              =   2280
            Y2              =   240
         End
         Begin VB.Label lblExp 
            Caption         =   "Exp: 1"
            Height          =   255
            Left            =   3720
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
      End
      Begin VB.TextBox txtLog 
         Height          =   2295
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   780
         Width           =   6495
      End
      Begin VB.TextBox txtCommand 
         Height          =   380
         Left            =   -74880
         TabIndex        =   1
         Top             =   3060
         Width           =   6495
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4683
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   3617
         EndProperty
      End
      Begin VB.Label lblCPS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "CPS: 0"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73080
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblGameTime 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   390
      End
   End
   Begin MSWinsockLib.Winsock Server_Socket 
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "&PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanChar 
         Caption         =   "Ban Char"
      End
      Begin VB.Menu mnuBanIp 
         Caption         =   "Ban IP"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban Ip & Char"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Access"
      End
      Begin VB.Menu mnuMod 
         Caption         =   "Set Mod"
      End
      Begin VB.Menu mnuMapper 
         Caption         =   "Set Mapper"
      End
      Begin VB.Menu mnuDev 
         Caption         =   "Set Dev"
      End
      Begin VB.Menu mnuOwner 
         Caption         =   "Set Owner"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnReload_Click(Index As Integer)
Dim i As Long
    
    Select Case Index
        Case 0
            Call LoadMaps
            TextAdd frmServer.txtLog, "Os mapas foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i), GetPlayerDir(i)
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
        Case 1
            Call LoadNpcs
            TextAdd frmServer.txtLog, "Os npc's foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        SendNpcs i
                    Else
                        Exit Sub
                    End If
                End If
            Next
        Case 2
            Call LoadSpawns
            TextAdd frmServer.txtLog, "Os pokémons foram atualizados!"
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        SendSpawns i
                    Else
                        Exit Sub
                    End If
                End If
            Next
    End Select
End Sub

Private Sub chkStaffOnly_Click()
Dim i As Long

    If chkStaffOnly.Value = YES Then
        '//Disconnect all non staff members
        If Player_HighIndex > 0 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).UseChar > 0 Then
                        If Player(i, TempPlayer(i).UseChar).Access <= 0 Then
                            Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "You have been disconnected from the server.", White, YES
                                Case LANG_EN: AddAlert i, "You have been disconnected from the server.", White, YES
                                Case LANG_ES: AddAlert i, "You have been disconnected from the server.", White, YES
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End If
End Sub

Private Sub cmdEditStore_Click()
    Call LoadVirtualShop
    frmEditor_Store.Show vbModeless, frmServer
    'Editor_Item.Show
End Sub

Private Sub cmdExp_Click()
    Dim i As Integer

    With EventExp
        If Not .ExpEvent Then
            .ExpEvent = True
            .ExpMultiply = scrlExp
            .ExpSecs = (txtExpHour * 3600)

            cmdExp.Caption = "Desativar"
            scrlExp.Enabled = False
            txtExpHour.Enabled = False

            If Player_HighIndex > 0 Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(i).UseChar > 0 Then
                            If Player(i, TempPlayer(i).UseChar).Access <= 0 Then
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Event Exp Activated.", White
                                Case LANG_EN: AddAlert i, "Event Exp Activated.", White
                                Case LANG_ES: AddAlert i, "Event Exp Activated.", White
                                End Select

                                SendEventInfo i
                            End If
                        End If
                    End If
                Next
            End If

        Else
            .ExpEvent = False
            .ExpMultiply = 0
            .ExpSecs = 0

            cmdExp.Caption = "Ativar"
            scrlExp.Enabled = True
            txtExpHour.Enabled = True

            If Player_HighIndex > 0 Then
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If TempPlayer(i).UseChar > 0 Then
                            If Player(i, TempPlayer(i).UseChar).Access <= 0 Then
                                Select Case TempPlayer(i).CurLanguage
                                Case LANG_PT: AddAlert i, "Event Exp Desactivated.", BrightRed
                                Case LANG_EN: AddAlert i, "Event Exp Desactivated.", BrightRed
                                Case LANG_ES: AddAlert i, "Event Exp Desactivated.", BrightRed
                                End Select

                                SendEventInfo i
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub cmdShutdown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutdown.Caption = "Desligar Servidor"
        SendGlobalMsg "Shutdown canceled.", White
    Else
        isShuttingDown = True
        cmdShutdown.Caption = "Cancelar Desligamento"
        Secs = 180
    End If
End Sub

Private Sub frmUser_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Command1_Click()
    AddPokemonsFishing
End Sub

Private Sub editorItem_Click()
    Dim Index As Long
    
    With frmEditor_Itens
        ' Limpa a lista
        .listIndex.Clear
        
        For Index = 1 To MAX_ITEM
            .listIndex.AddItem Index & ": " & Trim$(Item(Index).Name)
        Next
        
        ' Abrir o Editor
        .Show
        .listIndex.listIndex = 0
    End With
    
    
    Call ItensEditorInit
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
'Set the SortKey to the Index of the ColumnHeader - 1
'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not AppRunning Then Exit Sub
    
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

End Sub

Private Sub scrlExp_Change()
    frmServer.lblExp.Caption = scrlExp & " x"
End Sub

Private Sub Server_Socket_DataArrival(ByVal bytesTotal As Long)
    If IsServerConnected Then Call main_IncomingData(bytesTotal)
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim Count As Byte
Dim i As Long

    ' Check connection
    Count = 0
    For i = 1 To MAX_PLAYER
        If IsConnected(i) Then
            If GetPlayerIP(i) = Socket(Index).RemoteHostIP Then
                Count = Count + 1
                If Count >= 5 Then Exit Sub
            End If
        End If
    Next
    
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
Dim Count As Byte
Dim i As Long

    ' Check connection
    Count = 0
    For i = 1 To MAX_PLAYER
        If IsConnected(i) Then
            If GetPlayerIP(i) = Socket(Index).RemoteHostIP Then
                Count = Count + 1
                If Count >= 5 Then Exit Sub
            End If
        End If
    Next
    
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

' *****************
' ** Form object **
' *****************
Private Sub Form_Unload(Cancel As Integer)
    DestroyServer
End Sub

Sub mnuDisconnect_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If IsConnected(i) Then
        If GetPlayerIP(i) <> vbNullString Then
            CloseSocket i
        End If
    End If

End Sub

Sub mnuBanChar_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index
    
    ' Banir o Character
    If IsPlaying(i) Then
        BanCharacter Player(i, TempPlayer(i).UseChar).Name
        CloseSocket FindPlayer(Player(i, TempPlayer(i).UseChar).Name)
    End If
End Sub

Sub mnuBanIp_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If IsConnected(i) Then
        BanIP GetPlayerIP(i)
        CloseSocket i
    End If

End Sub

Sub mnuBan_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If IsPlaying(i) Then
        BanCharacter Player(i, TempPlayer(i).UseChar).Name
    End If

    If IsConnected(i) Then
        BanIP GetPlayerIP(i)
        CloseSocket i
    End If

End Sub

Sub mnuRemove_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If i > 0 Then
        If IsPlaying(i) Then
            Player(i, TempPlayer(i).UseChar).Access = ACCESS_NONE
            SendPlayerData i
        End If
    End If

End Sub

Sub mnuMod_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If i > 0 Then
        If IsPlaying(i) Then
            Player(i, TempPlayer(i).UseChar).Access = ACCESS_MODERATOR
            SendPlayerData i
        End If
    End If

End Sub

Sub mnuMapper_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If i > 0 Then
        If IsPlaying(i) Then
            Player(i, TempPlayer(i).UseChar).Access = ACCESS_MAPPER
            SendPlayerData i
        End If
    End If

End Sub

Sub mnuDev_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If i > 0 Then
        If IsPlaying(i) Then
            Player(i, TempPlayer(i).UseChar).Access = ACCESS_DEVELOPER
            SendPlayerData i
        End If
    End If

End Sub

Sub mnuOwner_click()
    Dim i As Long
    i = frmServer.lvwInfo.SelectedItem.Index

    If i > 0 Then
        If IsPlaying(i) Then
            Player(i, TempPlayer(i).UseChar).Access = ACCESS_CREATOR
            SendPlayerData i
        End If
    End If

End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
Dim Index As Long
Dim Command() As String
Dim chatMsg As String
Dim CurLanguage As Byte

    If KeyAscii = vbKeyReturn Then
        
        chatMsg = Trim$(txtCommand.Text)
        
        If Left$(chatMsg, 1) = "/" Then
            chatMsg = LCase(Trim$(txtCommand.Text))
            Command = Split(chatMsg, Space(1))
            
            Select Case Command(0)
                Case "/online"
                    TextAdd frmServer.txtLog, "Jogadores Online: " & TotalPlayerOnline
                
                Case "/clear"
                    txtLog.Text = vbNullString
                    
            End Select
            
        Else
        
            If LenB(Trim$(txtCommand.Text)) > 0 Then
                
                Select Case CurLanguage
                    Case LANG_PT: Call SendGlobalMsg("[SERVIDOR]: " + txtCommand.Text, White)
                    Case LANG_EN: Call SendGlobalMsg("[SERVER]: " + txtCommand.Text, White)
                    Case LANG_ES: Call SendGlobalMsg("[SERVIDOR]: " + txtCommand.Text, White)
                End Select

                For Index = 1 To MAX_PLAYER
                    If IsPlaying(Index) Then
                        If TempPlayer(Index).UseChar > 0 Then
                            Select Case CurLanguage
                                Case LANG_PT: AddAlert Index, "Servidor:" + txtCommand.Text, White
                                Case LANG_EN: AddAlert Index, "Server:" + txtCommand.Text, White
                                Case LANG_ES: AddAlert Index, "Servidor:" + txtCommand.Text, White
                            End Select
                        End If
                    End If
                Next
                
            End If
            
        End If

        KeyAscii = 0
        txtCommand.Text = vbNullString
    End If
End Sub

Private Sub txtExpHour_Change()
    If Not IsNumeric(txtExpHour) Then
        txtExpHour = 0
    End If
    
    If txtExpHour < 0 Then
        txtExpHour = 0
    End If
    
    If txtExpHour >= 99 Then
        txtExpHour = 48
    End If
End Sub

Private Sub txtLog_GotFocus()
    txtCommand.SetFocus
    DoEvents
End Sub

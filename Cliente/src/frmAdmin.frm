VERSION 5.00
Begin VB.Form frmAdmin 
   Caption         =   "Painel Administrativo"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Tools"
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   4095
      Begin VB.CommandButton Command10 
         Caption         =   "Map Report"
         Height          =   375
         Left            =   2520
         TabIndex        =   65
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
         Height          =   435
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkIsShiny 
         Caption         =   "Is Shiny?"
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cmbSpawn 
         Height          =   315
         ItemData        =   "frmAdmin.frx":0000
         Left            =   720
         List            =   "frmAdmin.frx":0002
         TabIndex        =   39
         Text            =   "Select"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Spawn"
         Height          =   435
         Left            =   1800
         TabIndex        =   38
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Invisivel"
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Loc"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Teleportar"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Spawn:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Administrator"
      Height          =   5895
      Left            =   4320
      TabIndex        =   9
      Top             =   0
      Width           =   4095
      Begin VB.ComboBox cmbBall 
         Height          =   315
         Left            =   2040
         TabIndex        =   63
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ComboBox cmbNature 
         Height          =   315
         Left            =   120
         TabIndex        =   62
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox chkAIv 
         Caption         =   "IVFull"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox chkAShiny 
         Caption         =   "IsShiny"
         Height          =   255
         Left            =   2040
         TabIndex        =   60
         Top             =   2880
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Caption         =   "Money e Cash"
         Height          =   1335
         Left            =   120
         TabIndex        =   49
         Top             =   4080
         Width           =   3855
         Begin VB.CommandButton Command9 
            Caption         =   "Enviar"
            Height          =   255
            Left            =   1200
            TabIndex        =   59
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtCash 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2520
            TabIndex        =   57
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optMoney 
            Caption         =   "Money"
            Height          =   255
            Left            =   840
            TabIndex        =   55
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optCash 
            Caption         =   "Cash"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Request"
            Height          =   255
            Left            =   2640
            TabIndex        =   53
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtCName 
            Height          =   285
            Left            =   720
            TabIndex        =   51
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "ADD:"
            Height          =   255
            Left            =   2040
            TabIndex        =   58
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblCash 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Player Cash: 0"
            Height          =   195
            Left            =   1680
            TabIndex        =   52
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.TextBox txtAccess 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   720
         TabIndex        =   46
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Acesso"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Unmut"
         Height          =   255
         Left            =   3360
         TabIndex        =   44
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Mute"
         Height          =   255
         Left            =   2760
         TabIndex        =   43
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtAItem 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   37
         Text            =   "0"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtAPoke 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Text            =   "0"
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton optGiveItem 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1920
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton optGivePoke 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2520
         Width           =   255
      End
      Begin VB.HScrollBar scrlAPoke 
         Height          =   255
         Left            =   360
         Min             =   1
         TabIndex        =   32
         Top             =   2520
         Value           =   1
         Width           =   1815
      End
      Begin VB.TextBox txtBName 
         Height          =   285
         Left            =   720
         TabIndex        =   24
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   360
         Min             =   1
         TabIndex        =   17
         Top             =   1920
         Value           =   1
         Width           =   1815
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Entregar"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   15
         Text            =   "1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdAtt 
         Caption         =   "Obter"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Puxar"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Acces:"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Left            =   2280
         TabIndex        =   56
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Find:"
         Height          =   195
         Left            =   2280
         TabIndex        =   48
         Top             =   2280
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pokemon Reborn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   47
         Top             =   5400
         Width           =   3060
      End
      Begin VB.Label lblAPoke 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give Poke: No Name"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3960
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   465
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblAItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Give Item: No Name"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-Quantidade-"
         Height          =   195
         Left            =   3000
         TabIndex        =   18
         Top             =   2040
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdEditNpc 
         Caption         =   "Editar Npc"
         Height          =   495
         Left            =   2040
         TabIndex        =   21
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdEditMove 
         Caption         =   "Editar Move"
         Height          =   495
         Left            =   2040
         TabIndex        =   20
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdConv 
         Caption         =   "Editar Conversas"
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdMap 
         Caption         =   "Editar Mapa"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdItem 
         Caption         =   "Editar Itens"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdSpawn 
         Caption         =   "Editar Spawn"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdEditPokemon 
         Caption         =   "Editar Pokemons"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdShop 
         Caption         =   "Editar Lojas"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton CmdAnimation 
         Caption         =   "Editar animações"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdQuest 
         Caption         =   "Editar Quest's"
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Top             =   1440
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdABan_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione um nome antes.", BrightRed
            Exit Sub
        End If

        SendBanPlayer txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub cmdAKick_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione um nome antes.", BrightRed
            Exit Sub
        End If

        SendKickPlayer txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub CmdAnimation_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditAnimation
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub cmdASpawn_Click()
    If optGiveItem.value Then
        If Player(MyIndex).Access >= ACCESS_CREATOR Then
            If txtBName = vbNullString Then
                AddText "Adicione um nome.", BrightRed
                Exit Sub
            End If

            If Not IsNumeric(txtAmount) Then
                AddText "Quantidade sem valor numerico.", BrightRed
                Exit Sub
            End If

            If scrlAItem = 0 Then
                AddText "Adicione um item com ID maior que zero.", BrightRed
                Exit Sub
            End If

            SendGiveItemTo txtBName, scrlAItem, txtAmount
        Else
            AddText "Invalid command!", BrightRed
            Exit Sub
        End If
    Else
        If Player(MyIndex).Access >= ACCESS_CREATOR Then
            If txtBName = vbNullString Then
                AddText "Adicione um nome", BrightRed
                Exit Sub
            End If

            If Not IsNumeric(txtAmount) Then
                AddText "Quantidade não numerica.", BrightRed
                Exit Sub
            End If

            SendGivePokemonTo txtBName, scrlAPoke, CLng(txtAmount), chkAShiny, chkAIv, cmbNature.ListIndex - 1, cmbBall.ListIndex
        Else
            AddText "Invalid command!", BrightRed
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAtt_Click()
    If optGiveItem Then
        If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
            If scrlAItem = 0 Then
                AddText "Adicione um item com ID maior que zero.", BrightRed
                Exit Sub
            End If

            If Not IsNumeric(txtAmount) Then
                AddText "Adicione valor numerico à quantidade.", BrightRed
                Exit Sub
            End If

            SendGetItem scrlAItem, txtAmount
        Else
            AddText "Invalid command!", BrightRed
            Exit Sub
        End If
    Else
        If Player(MyIndex).Access >= ACCESS_CREATOR Then

            If Not IsNumeric(txtAmount) Then
                AddText "Level não numerico.", BrightRed
                Exit Sub
            End If

            If scrlAPoke = 0 Then
                AddText "Adicione um ID maior que zero.", brighred
                Exit Sub
            End If
            
            
            SendGivePokemonTo Trim$(Player(MyIndex).Name), scrlAPoke, CLng(txtAmount), chkAShiny, chkAIv, cmbNature.ListIndex - 1, cmbBall.ListIndex
        Else
            AddText "Invalid command!", BrightRed
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAWarp_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAMap = vbNullString Then
            AddText "Adicione o número do mapa antes", BrightRed
            Exit Sub
        End If

        If Not IsNumeric(txtAMap) Then
            AddText "Somente numeros no mapa", BrightRed
            Exit Sub
        End If

        If PlayerPokemon(MyIndex).Num > 0 Then
            AddText "Tire seu pokemon do mapa antes", BrightRed
            Exit Sub
        End If

        SendWarpTo Val(txtAMap)
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub cmdAWarp2Me_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione algum nome.", BrightRed
            Exit Sub
        End If

        If PlayerPokemon(MyIndex).Num > 0 Then
            AddText "Retire seu pokemon do mapa.", BrightRed
            Exit Sub
        End If

        SendWarpToMe txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub cmdAWarpMe2_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione algum nome.", BrightRed
            Exit Sub
        End If

        If PlayerPokemon(MyIndex).Num > 0 Then
            AddText "Retire seu pokemon do mapa.", BrightRed
            Exit Sub
        End If

        SendWarpMeTo txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub CmdConv_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditConversation
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub cmdEditMove_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditPokemonMove
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub cmdEditNpc_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditNpc
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdEditPokemon_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditPokemon
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdItem_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditItem
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdMap_Click()
    If Player(MyIndex).Access >= ACCESS_MAPPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditMap
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdQuest_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditQuest
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdShop_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditShop
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub CmdSpawn_Click()
    If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
        If GameSetting.Fullscreen = YES Then
            AddText "You cannot open any editor in fullscreen mode", BrightRed
        Else
            SendRequestEditSpawn
        End If
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub Command1_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        SendStealthMode
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command10_Click()
    If Player(MyIndex).Access < ACCESS_MODERATOR Then Exit Sub
    
    SendMapReport
End Sub

Private Sub Command2_Click()
    If Player(MyIndex).Access >= ACCESS_CREATOR Then
        If txtAccess < 0 Or txtAccess > 5 Then
            AddText "Acessos do 0 ao 5", BrightRed
            Exit Sub
        End If

        If Not IsNumeric(txtAccess) Then
            AddText "Utilize numeros no acesso.", BrightRed
            Exit Sub
        End If

        If txtAName = vbNullString Then
            AddText "Adicione algum nome.", BrightRed
            Exit Sub
        End If

        SendSetAccess txtAName, txtAccess
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command3_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        ShowLoc = Not ShowLoc
    Else
        AddText "Invalid command!", BrightRed
    End If
End Sub

Private Sub Command4_Click()
    If Player(MyIndex).Access >= ACCESS_CREATOR Then
        SendSpawnPokemon Me.cmbSpawn.ListIndex, chkIsShiny
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Me.cmbSpawn.Clear
    Me.cmbSpawn.AddItem "None."
    For i = 1 To MAX_POKEMON
        If Spawn(i).PokeNum > 0 Then
            Me.cmbSpawn.AddItem i & ": " & Trim$(Pokemon(Spawn(i).PokeNum).Name)
        Else
            Me.cmbSpawn.AddItem i & ": Vazio"
        End If
        DoEvents
    Next i
End Sub

Private Sub Command6_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione um nome.", BrightRed
            Exit Sub
        End If

        SendMutePlayer txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command7_Click()
    If Player(MyIndex).Access >= ACCESS_MODERATOR Then
        If txtAName = vbNullString Then
            AddText "Adicione um nome.", BrightRed
            Exit Sub
        End If

        SendUnmutePlayer txtAName
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command8_Click()
    
    If Player(MyIndex).Access >= ACCESS_CREATOR Then
        SendRequestPlayerValue txtCName, optCash
    Else
        AddText "Invalid command!", BrightRed
        Exit Sub
    End If
End Sub

Private Sub Command9_Click()
    SendCashValueTo txtCName, txtCash, optCash
End Sub

Private Sub Form_Load()

    Dim i As Integer

    scrlAItem.max = MAX_ITEM
    scrlAItem.min = 0
    scrlAItem = 0

    scrlAPoke.max = MAX_POKEMON
    scrlAPoke.min = 0
    scrlAPoke = 0

    ' Poke Spawn
    Me.cmbSpawn.Clear
    Me.cmbSpawn.AddItem "None."
    For i = 1 To MAX_POKEMON
        If Spawn(i).PokeNum > 0 Then
            Me.cmbSpawn.AddItem i & ": " & Trim$(Pokemon(Spawn(i).PokeNum).Name)
        Else
            Me.cmbSpawn.AddItem i & ": Vazio"
        End If
        DoEvents
    Next i
    
    ' Poke Nature
    Me.cmbNature.Clear
    Me.cmbNature.AddItem "None."
    For i = 0 To PokemonNature.PokemonNature_Count - 1
        Me.cmbNature.AddItem i & ": " & CheckNatureString(i)
        DoEvents
    Next i
    
    ' Poke Balls
    Me.cmbBall.Clear
    For i = 0 To BallEnum.BallEnum_Count - 1
        Me.cmbBall.AddItem i & ": " & CheckPokeBallString(i)
        DoEvents
    Next i
End Sub

Private Sub optGiveItem_Click()
    If optGivePoke Then
        lblAmount = "-Level-"
        txtAmount = 1
    Else
        lblAmount = "-Quantidade-"
        txtAmount = 1
    End If
End Sub

Private Sub optGivePoke_Click()
    If optGivePoke Then
        lblAmount = "-Level-"
        txtAmount = 1
    Else
        lblAmount = "-Quantidade-"
        txtAmount = 1
    End If
End Sub

Private Sub scrlAItem_Change()
    If scrlAItem > 0 Then
        If LenB(Trim$(Item(scrlAItem).Name)) > 0 Then
            lblAItem = "Give Item: " & scrlAItem & "-" & Trim$(Item(scrlAItem).Name)
        Else
            lblAItem = "Give Item: No Name"
        End If
    Else
        lblAItem = "Give Item: No Name"
    End If
End Sub

Private Sub scrlAPoke_Change()
    If scrlAPoke > 0 Then
        If LenB(Trim$(Pokemon(scrlAPoke).Name)) > 0 Then
            lblAPoke = "Give Poke: " & scrlAPoke & "-" & Trim$(Pokemon(scrlAPoke).Name)
        Else
            lblAPoke = "Give Poke: No Name"
        End If
    Else
        lblAPoke = "Give Poke: No Name"
    End If
End Sub

Private Sub txtAItem_Change()

    Dim Find As String, i As Long

    If Not IsNumeric(txtAItem) Then
        Find = UCase$(Trim$(txtAItem.Text))
        If Len(Find) <= 2 And Not Find = "" Then
            lblAItem = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_ITEM
            If Not Find = "" Then
                If InStr(1, UCase$(Trim$(Item(i).Name)), Find) > 0 Then
                    scrlAItem = i
                    Exit Sub
                End If
            End If
        Next
    Else
        If txtAItem > MAX_ITEM Then
            txtAItem = MAX_ITEM
        ElseIf txtAItem <= 0 Then
            txtAItem = 1
        End If
        scrlAItem = txtAItem
    End If
End Sub

Private Sub txtAPoke_Change()
    
    Dim Find As String, i As Long

    If Not IsNumeric(txtAPoke) Then
        Find = UCase$(Trim$(txtAPoke.Text))
        If Len(Find) <= 2 And Not Find = "" Then
            lblAPoke = "Adicione mais letras."
            Exit Sub
        End If

        For i = 1 To MAX_POKEMON
            If Not Find = "" Then
                If InStr(1, UCase$(Trim$(Pokemon(i).Name)), Find) > 0 Then
                    scrlAPoke = i
                    Exit Sub
                End If
            End If
        Next
    Else
        If txtAPoke > MAX_POKEMON Then
            txtAPoke = MAX_POKEMON
        ElseIf txtAPoke <= 0 Then
            txtAPoke = 1
        End If
        scrlAPoke = txtAPoke
    End If
End Sub





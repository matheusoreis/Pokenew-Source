VERSION 5.00
Begin VB.Form frmEditor_Pokemon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Editor"
   ClientHeight    =   11430
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   25380
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   762
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1692
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frameIndex 
      Caption         =   "Lista de pokemons"
      Height          =   6135
      Left            =   10080
      TabIndex        =   97
      Top             =   0
      Width           =   3495
      Begin VB.ListBox listIndex 
         Height          =   4350
         ItemData        =   "frmEditor_Pokemon.frx":0000
         Left            =   120
         List            =   "frmEditor_Pokemon.frx":0002
         TabIndex        =   100
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton buttonCopy 
         Caption         =   "Copiar"
         Height          =   615
         Left            =   120
         TabIndex        =   99
         Top             =   4755
         Width           =   3255
      End
      Begin VB.CommandButton buttonPaste 
         Caption         =   "Colar"
         Height          =   615
         Left            =   120
         TabIndex        =   98
         Top             =   5400
         Width           =   3255
      End
   End
   Begin VB.Frame frameProperties 
      Caption         =   "Propriedades"
      Height          =   6135
      Left            =   13680
      TabIndex        =   88
      Top             =   0
      Width           =   11655
      Begin VB.Frame frameStatus 
         Caption         =   "Status do pokemon"
         Height          =   2175
         Left            =   120
         TabIndex        =   110
         Top             =   2880
         Width           =   5415
         Begin VB.HScrollBar HScroll3 
            Height          =   315
            Left            =   2760
            Max             =   0
            TabIndex        =   122
            Top             =   1680
            Width           =   2535
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   315
            Left            =   2760
            Max             =   0
            TabIndex        =   120
            Top             =   1080
            Width           =   2535
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   315
            Left            =   2760
            Max             =   0
            TabIndex        =   118
            Top             =   480
            Width           =   2535
         End
         Begin VB.HScrollBar scrollDefense 
            Height          =   315
            Left            =   120
            Max             =   0
            TabIndex        =   116
            Top             =   1680
            Width           =   2535
         End
         Begin VB.HScrollBar scrollAttack 
            Height          =   315
            Left            =   120
            Max             =   0
            TabIndex        =   114
            Top             =   1080
            Width           =   2535
         End
         Begin VB.HScrollBar scrollHp 
            Height          =   315
            Left            =   120
            Max             =   0
            TabIndex        =   112
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label30 
            Caption         =   "Velocidade de Ataque"
            Height          =   255
            Left            =   2760
            TabIndex        =   121
            Top             =   1440
            Width           =   2520
         End
         Begin VB.Label labelSpeedDefense 
            Caption         =   "Esquiva"
            Height          =   255
            Left            =   2760
            TabIndex        =   119
            Top             =   840
            Width           =   2520
         End
         Begin VB.Label labelSpeed 
            Caption         =   "Velocidade do pokémon"
            Height          =   255
            Left            =   2760
            TabIndex        =   117
            Top             =   240
            Width           =   2520
         End
         Begin VB.Label labelDefense 
            Caption         =   "Defesa do pokémon"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1440
            Width           =   2520
         End
         Begin VB.Label labelAttack 
            Caption         =   "Ataque do pokémon"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   840
            Width           =   2520
         End
         Begin VB.Label labelHp 
            Caption         =   "Vida do pokémon"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   240
            Width           =   2520
         End
      End
      Begin VB.Frame frameDetails 
         Caption         =   "Detalhes"
         Height          =   2655
         Left            =   120
         TabIndex        =   93
         Top             =   240
         Width           =   5415
         Begin VB.CheckBox checkStackItemDetails 
            Caption         =   "Pokémon lendário?"
            Height          =   375
            Left            =   120
            TabIndex        =   109
            Top             =   2160
            Width           =   5175
         End
         Begin VB.TextBox textSound 
            Height          =   320
            Left            =   2760
            TabIndex        =   108
            Top             =   1800
            Width           =   2535
         End
         Begin VB.ComboBox comboTypeBehaviour 
            Height          =   315
            ItemData        =   "frmEditor_Pokemon.frx":0004
            Left            =   120
            List            =   "frmEditor_Pokemon.frx":0023
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   1800
            Width           =   2535
         End
         Begin VB.VScrollBar scrollNamePosition 
            Height          =   975
            Left            =   5000
            Max             =   200
            Min             =   -200
            TabIndex        =   104
            Top             =   480
            Width           =   315
         End
         Begin VB.HScrollBar scrollPokemonSprite 
            Height          =   315
            Left            =   120
            Max             =   0
            TabIndex        =   102
            Top             =   1125
            Width           =   3735
         End
         Begin VB.PictureBox picturePokemonSprite 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   3960
            ScaleHeight     =   66.065
            ScaleMode       =   0  'User
            ScaleWidth      =   68.266
            TabIndex        =   101
            Top             =   480
            Width           =   960
         End
         Begin VB.TextBox textShopName 
            Height          =   320
            Left            =   120
            TabIndex        =   94
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label labelSound 
            Caption         =   "Som:"
            Height          =   255
            Left            =   2760
            TabIndex        =   107
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label labelPositionName 
            Caption         =   "Y: "
            Height          =   255
            Left            =   3960
            TabIndex        =   105
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label labelPokemonSprite 
            Caption         =   "item: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label labelPokemonName 
            Caption         =   "Nome do pokémon"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   3720
         End
         Begin VB.Label labelBehaviour 
            Caption         =   "Comportamento:"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   1560
            Width           =   2520
         End
      End
      Begin VB.Frame frameShop 
         Caption         =   "Lista de itens:"
         Height          =   3495
         Left            =   7320
         TabIndex        =   89
         Top             =   2520
         Width           =   3975
         Begin VB.ListBox listShopItens 
            Height          =   2010
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   3735
         End
         Begin VB.PictureBox pictureShopItemName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   3360
            ScaleHeight     =   32
            ScaleMode       =   0  'User
            ScaleWidth      =   32
            TabIndex        =   90
            Top             =   2880
            Width           =   480
         End
         Begin VB.Label labelShopItens 
            Caption         =   "Lista dos Itens:"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   320
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Moveset"
      Height          =   2175
      Left            =   120
      TabIndex        =   81
      Top             =   9000
      Width           =   4575
      Begin VB.ListBox lstMoveset 
         Height          =   1035
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtMoveLevel 
         Height          =   285
         Left            =   3600
         TabIndex        =   85
         Text            =   "0"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbMoveNum 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1800
         TabIndex        =   83
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   255
         Left            =   3480
         TabIndex        =   82
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblMovesetNum 
         Caption         =   "Move: "
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Item Drop"
      Height          =   1815
      Left            =   4800
      TabIndex        =   74
      Top             =   9000
      Width           =   4575
      Begin VB.ListBox lstItemDrop 
         Height          =   840
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   4335
      End
      Begin VB.ComboBox cmbItemNum 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtItemDropRate 
         Height          =   285
         Left            =   1800
         TabIndex        =   76
         Text            =   "0"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtItemSearch 
         Height          =   285
         Left            =   1200
         TabIndex        =   75
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Rarity 0% - 100%"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Index"
      Height          =   9015
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdIndexSearch 
         Caption         =   "Find"
         Height          =   255
         Left            =   2040
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIndexSearch 
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstIndex 
         Height          =   8250
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   9015
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkLendary 
         Caption         =   "Lendary?"
         Height          =   255
         Left            =   4920
         TabIndex        =   73
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1440
         Width           =   3495
      End
      Begin VB.VScrollBar scrlOffSetY 
         Height          =   1215
         Left            =   4800
         Max             =   200
         Min             =   -200
         TabIndex        =   70
         Top             =   120
         Width           =   135
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   2880
         TabIndex        =   67
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame fraEvolve 
         Caption         =   "Evolve - 1"
         Height          =   1815
         Left            =   120
         TabIndex        =   55
         Top             =   5400
         Width           =   5775
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   4680
            TabIndex        =   68
            Top             =   720
            Width           =   855
         End
         Begin VB.HScrollBar scrlEvolveCondition 
            Height          =   255
            Left            =   3000
            Max             =   0
            TabIndex        =   63
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtEvolveLevel 
            Height          =   285
            Left            =   1440
            TabIndex        =   60
            Text            =   "0"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtEvolveConditionData 
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Text            =   "0"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.HScrollBar scrlEvolve 
            Height          =   255
            Left            =   2400
            Max             =   0
            TabIndex        =   57
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlEvolveIndex 
            Height          =   255
            Left            =   120
            Max             =   1
            Min             =   1
            TabIndex        =   56
            Top             =   240
            Value           =   1
            Width           =   5535
         End
         Begin VB.Label lblEvolveCondition 
            Caption         =   "Condition: None"
            Height          =   255
            Left            =   3000
            TabIndex        =   64
            Top             =   1080
            Width           =   2655
         End
         Begin VB.Label Label21 
            Caption         =   "Evolve Level:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Condition Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblEvolve 
            Caption         =   "Evolve To: None"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   5295
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5640
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1560
         Max             =   10
         TabIndex        =   53
         Top             =   8640
         Width           =   4335
      End
      Begin VB.TextBox txtSpecies 
         Height          =   285
         Left            =   1560
         TabIndex        =   51
         Top             =   8280
         Width           =   4335
      End
      Begin VB.TextBox txtPokedexEntry 
         Height          =   855
         Left            =   1560
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   7320
         Width           =   4335
      End
      Begin VB.TextBox txtWeight 
         Height          =   285
         Left            =   4440
         TabIndex        =   48
         Text            =   "0"
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   4440
         TabIndex        =   45
         Text            =   "0"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.ComboBox cmbGrowthRate 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox txtBaseExp 
         Height          =   285
         Left            =   1440
         TabIndex        =   42
         Text            =   "0"
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtEVYeildVal 
         Height          =   285
         Left            =   4440
         TabIndex        =   40
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox cmbEVYeildType 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3960
         Width           =   1455
      End
      Begin VB.ComboBox cmbEggGroup 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox txtEggCycle 
         Height          =   285
         Left            =   1440
         TabIndex        =   33
         Text            =   "0"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtFemaleRate 
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Text            =   "0"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtCatchRate 
         Height          =   285
         Left            =   4320
         TabIndex        =   29
         Text            =   "0"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Type"
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   2775
         Begin VB.ComboBox cmbSecondaryType 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cmbPrimaryType 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Secondary:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Primary:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Base Stats"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   5775
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   6
            Left            =   4800
            TabIndex        =   23
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   5
            Left            =   2760
            TabIndex        =   21
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   19
            Text            =   "0"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   3
            Left            =   4800
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   2
            Left            =   2760
            TabIndex        =   15
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Spd:"
            Height          =   255
            Left            =   4080
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "SpDef:"
            Height          =   255
            Left            =   2040
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "SpAtk:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Def:"
            Height          =   255
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Atk:"
            Height          =   255
            Left            =   2040
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "HP:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox chkScale 
         Caption         =   "Scale"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   5160
         ScaleHeight     =   66.065
         ScaleMode       =   0  'User
         ScaleWidth      =   68.266
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   0
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label23 
         Caption         =   "Cries:"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblOffSetY 
         Caption         =   "Name OffSetY: 0"
         Height          =   255
         Left            =   3120
         TabIndex        =   69
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblRange 
         Caption         =   "Range: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   8640
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Species:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   8280
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Pokedex Entry:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Weight:"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Height:"
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Growth Rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Base Exp:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "EV Yeild Val:"
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "EV Yeild Type:"
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Egg Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Egg Cycle:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Female Rate:"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Catch Rate:"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Behaviour:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit(Esc)"
      End
   End
End
Attribute VB_Name = "frmEditor_Pokemon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


